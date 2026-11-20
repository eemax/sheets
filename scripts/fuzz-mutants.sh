#!/bin/bash
# scripts/fuzz-mutants.sh OUT [--only NAME]: checks the fuzzer finds each of
# wave 1's bugs, put back into sheets (internal/fuzz/mutants/NAME.patch).
#
# It copies the working tree as it is (tracked and untracked files, ignored
# ones left out: uncommitted work included; never HEAD, a branch or a
# worktree of this checkout) into a temporary folder: once as it is, the
# base, whose fuzz test binary it builds, and for each mutant once more, with
# its patch applied. It runs the fuzzer's long mode on the patched copy
# until the mutant's finding, for at most 4 minutes (SHEETS_FUZZ_UNTIL=new,
# SHEETS_FUZZ_TIME=4m): the first finding known.txt does not list (by its
# signature, or as a bug it names) whose sequence the base makes with no
# finding (SHEETS_FUZZ_BASE: a finding the code makes without the mutant is
# not the mutant's, whatever the seed). A mutant is found when the run stops
# at such a finding, shrunk.
#
# It writes OUT/NAME.log for each (the mutant, the seed, the repro, the time
# to the finding, the run's log) and OUT/summary.txt, and exits 0 only if
# every mutant is found. When uv cannot make long mode's pools (no uv, or no
# network for openpyxl and XlsxWriter), each mutant runs the default mode on
# the committed pool instead (seed 1, at most 2,000 edits, with the same stop
# and the same base), and the logs say so.
#
#   SHEETS_FUZZ_MUTANT_SEED   the seed (default 7)
#   SHEETS_FUZZ_MUTANT_TIME   the cap (default 4m)
set -euo pipefail
root=$(cd "$(dirname "$0")/.." && pwd)
usage() { echo "usage: scripts/fuzz-mutants.sh OUT [--only NAME]" >&2; exit 2; }
[ $# -ge 1 ] || usage
out=$1
shift
only=""
while [ $# -gt 0 ]; do
  case $1 in
    --only) [ $# -ge 2 ] || usage; only=$2; shift 2 ;;
    *) usage ;;
  esac
done
seed=${SHEETS_FUZZ_MUTANT_SEED:-7}
cap=${SHEETS_FUZZ_MUTANT_TIME:-4m}
mutants="$root/internal/fuzz/mutants"
if [ -n "$only" ] && [ ! -f "$mutants/$only.patch" ]; then
  echo "no mutant $only (have: $(cd "$mutants" && ls *.patch | sed 's/\.patch$//' | tr '\n' ' '))" >&2
  exit 2
fi
mkdir -p "$out"
out=$(cd "$out" && pwd)
tmp=$(mktemp -d "${TMPDIR:-/tmp}/sheets-mutants.XXXXXX")
trap 'rm -rf "$tmp"' EXIT

# copy_tree DIR: the working tree as it is, into DIR.
copy_tree() {
  (cd "$root" && git ls-files -co --exclude-standard -z |
    while IFS= read -r -d '' f; do [ -e "$f" ] && printf '%s\0' "$f"; done |
    tar --null -T - -cf -) | tar -xf - -C "$1"
}

# The base: the working tree without a mutant, its fuzz test binary built, to
# replay the sequence of each finding a mutant's run does not know.
base="$tmp/base"
mkdir -p "$base"
copy_tree "$base"
if ! (cd "$base" && go test -c -o internal/fuzz/fuzz.test ./internal/fuzz) >"$tmp/base.log" 2>&1; then
  echo "the working tree's fuzz tests do not build: $(tail -3 "$tmp/base.log")" >&2
  exit 2
fi

# Long mode needs fresh pools: can uv make one?
mode=long
probe=$(cd "$root" && uv run --no-project --with openpyxl==3.1.5 --with xlsxwriter==3.3.1 \
  python scripts/fuzzgen.py --seed 1000 --count 2 --pool-only "$tmp/probe" 2>&1) || mode=default
if [ $mode = default ]; then
  echo "uv cannot make long mode's pools ($(echo "$probe" | tail -1)): each mutant runs the default mode on the committed pool" >&2
fi

found=0 missed=0
summary="$out/summary.txt"
{
  echo "fuzz-mutants $(date -u +%Y-%m-%dT%H:%M:%SZ), from $root ($(git -C "$root" rev-parse --short HEAD), working tree as it is)"
  if [ $mode = long ]; then
    echo "mode: long, seed $seed, cap $cap, stopping at the mutant's finding: the first known.txt does not list whose sequence the tree without the mutant makes with no finding"
  else
    echo "mode: default (uv could not make pools), seed 1, at most 2,000 edits on the committed pool, with the same stop"
  fi
} > "$summary"
for patch in "$mutants"/*.patch; do
  name=$(basename "$patch" .patch)
  [ -z "$only" ] || [ "$name" = "$only" ] || continue
  dir="$tmp/$name"
  mkdir -p "$dir"
  copy_tree "$dir"
  log="$out/$name.log"
  {
    echo "mutant $name ($(date -u +%Y-%m-%dT%H:%M:%SZ))"
    sed -n '/^--- a\//q;p' "$patch"
  } > "$log"
  if ! (cd "$dir" && git apply "$patch") >> "$log" 2>&1; then
    echo "result: NOT RUN: the patch does not apply to the working tree" | tee -a "$log"
    echo "$name: not run (the patch does not apply)" >> "$summary"
    missed=$((missed + 1))
    continue
  fi
  if [ $mode = long ]; then
    echo "run: SHEETS_FUZZ=long SHEETS_FUZZ_SEED=$seed SHEETS_FUZZ_TIME=$cap SHEETS_FUZZ_UNTIL=new SHEETS_FUZZ_BASE=BASE go test ./internal/fuzz -run 'TestFuzz\$' -count=1 -v" >> "$log"
  else
    echo "run: SHEETS_FUZZ_UNTIL=new SHEETS_FUZZ_BASE=BASE go test ./internal/fuzz -run 'TestFuzz\$' -count=1 -v (default mode: uv could not make pools)" >> "$log"
  fi
  echo "(BASE: a copy of the working tree without the mutant, its fuzz test binary built)" >> "$log"
  echo >> "$log"
  start=$(date +%s)
  set +e
  if [ $mode = long ]; then
    (cd "$dir" && SHEETS_FUZZ=long SHEETS_FUZZ_SEED=$seed SHEETS_FUZZ_TIME=$cap SHEETS_FUZZ_UNTIL=new SHEETS_FUZZ_BASE="$base" \
      go test ./internal/fuzz -run 'TestFuzz$' -count=1 -v -timeout 20m) >> "$log" 2>&1
  else
    (cd "$dir" && SHEETS_FUZZ_UNTIL=new SHEETS_FUZZ_BASE="$base" \
      go test ./internal/fuzz -run 'TestFuzz$' -count=1 -v -timeout 20m) >> "$log" 2>&1
  fi
  rc=$?
  set -e
  secs=$(($(date +%s) - start))
  rm -rf "$dir"
  # Found: the run stopped at the mutant's finding (one known.txt does not
  # list, whose sequence the base makes with no finding), and shrank it.
  stop=$(grep -m1 "the mutant's finding: " "$log" || true)
  sig=$(echo "$stop" | sed -n "s/.*the mutant's finding: \(.*\), after \([0-9]*\)s: .*/\1/p")
  after=$(echo "$stop" | sed -n "s/.*the mutant's finding: .*, after \([0-9]*\)s: .*/\1/p")
  shrunk=$(grep -B1 "^ *finding $sig\$" "$log" | grep -m1 "a finding not in known.txt, shrunk" || true)
  if [ $rc -ne 0 ] && [ -n "$sig" ] && [ -n "$shrunk" ]; then
    repro=$(echo "$shrunk" | sed -n 's/.*(replay: \(.*\)):$/\1/p')
    line="$name: found, $sig, after ${after}s (the run took ${secs}s with shrinking); repro: $repro"
    found=$((found + 1))
  else
    line="$name: NOT FOUND (exit $rc, ${secs}s)"
    missed=$((missed + 1))
  fi
  echo "result: $line" >> "$log"
  echo "$line" | tee -a "$summary"
  others=$(grep -c "and the tree without the mutant makes its sequence with a finding too" "$log" || true)
  if [ "$others" -gt 0 ]; then
    echo "  ($others findings known.txt does not list were the code's, not the mutant's: the base makes their sequences with a finding too; see $name.log)" | tee -a "$summary"
  fi
done
echo "$found found, $missed not" | tee -a "$summary"
[ "$missed" -eq 0 ] && [ "$found" -gt 0 ]
