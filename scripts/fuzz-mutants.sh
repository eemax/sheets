#!/bin/bash
# scripts/fuzz-mutants.sh OUT [--only NAME]: checks the fuzzer finds each of
# wave 1's bugs, put back into sheets (internal/fuzz/mutants/NAME.patch).
#
# For each mutant it copies the working tree as it is (tracked and untracked
# files, ignored ones left out: uncommitted work included; never HEAD, a
# branch or a worktree of this checkout) into a temporary folder, applies the
# patch there, and runs the fuzzer's long mode on it until its first finding
# known.txt does not list, for at most 4 minutes (SHEETS_FUZZ_UNTIL=new,
# SHEETS_FUZZ_TIME=4m). The seed is long mode's run in known.txt (7), so the
# findings of the code as it is are all listed, and a new one is the mutant's.
# A mutant is found when the run stops at such a finding, shrunk.
#
# It writes OUT/NAME.log for each (the mutant, the seed, the repro, the time
# to the finding, the run's log) and OUT/summary.txt, and exits 0 only if
# every mutant is found. When uv cannot make long mode's pools (no uv, or no
# network for openpyxl and XlsxWriter), each mutant runs the default mode on
# the committed pool instead (seed 1, 2,000 edits, its findings all in
# known.txt), and the logs say so.
#
#   SHEETS_FUZZ_MUTANT_SEED   the seed (default 7: long mode's in known.txt)
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
    echo "mode: long, seed $seed, cap $cap, stopping at the first finding known.txt does not list"
  else
    echo "mode: default (uv could not make pools), seed 1, 2,000 edits on the committed pool"
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
    echo "run: SHEETS_FUZZ=long SHEETS_FUZZ_SEED=$seed SHEETS_FUZZ_TIME=$cap SHEETS_FUZZ_UNTIL=new go test ./internal/fuzz -run 'TestFuzz\$' -count=1 -v" >> "$log"
  else
    echo "run: go test ./internal/fuzz -run 'TestFuzz\$' -count=1 -v (default mode: uv could not make pools)" >> "$log"
  fi
  echo >> "$log"
  start=$(date +%s)
  set +e
  if [ $mode = long ]; then
    (cd "$dir" && SHEETS_FUZZ=long SHEETS_FUZZ_SEED=$seed SHEETS_FUZZ_TIME=$cap SHEETS_FUZZ_UNTIL=new \
      go test ./internal/fuzz -run 'TestFuzz$' -count=1 -v -timeout 20m) >> "$log" 2>&1
  else
    (cd "$dir" && go test ./internal/fuzz -run 'TestFuzz$' -count=1 -v -timeout 20m) >> "$log" 2>&1
  fi
  rc=$?
  set -e
  secs=$(($(date +%s) - start))
  rm -rf "$dir"
  # Found: the run failed on a finding known.txt does not list, shrunk.
  first=$(grep -m1 "a finding not in known.txt, shrunk" "$log" || true)
  if [ $rc -ne 0 ] && [ -n "$first" ]; then
    sig=$(grep -m1 -A1 "a finding not in known.txt, shrunk" "$log" | sed -n 's/^ *finding //p')
    after=$(echo "$first" | sed -n 's/.*shrunk, after \([0-9]*s\).*/\1/p')
    repro=$(echo "$first" | sed -n 's/.*(replay: \(.*\)):$/\1/p')
    line="$name: found, $sig, after $after (the run took ${secs}s with shrinking); repro: $repro"
    found=$((found + 1))
  else
    line="$name: NOT FOUND (exit $rc, ${secs}s)"
    missed=$((missed + 1))
  fi
  echo "result: $line" >> "$log"
  echo "$line" | tee -a "$summary"
done
echo "$found found, $missed not" | tee -a "$summary"
[ "$missed" -eq 0 ] && [ "$found" -gt 0 ]
