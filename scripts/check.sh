#!/bin/bash
# sheets' checks for `portal check` (.portal-checks.yaml): `scripts/check.sh fast|full`.
# The last line is the total that `count:` reads: tests passed, subtests included, summed over
# every `go test` the mode runs.
set -euo pipefail
cd "$(dirname "$0")/.."
mode=${1:-full}

stream_pkgs="./cmd/... ./internal/xlsx/... ./internal/cli/..."

# -json for an exact count; the summary prints each package's line and a failing test's output.
summarize() {
  python3 -I -c '
import json, sys
label = sys.argv[1]
out, passed, failed = {}, 0, []
for line in sys.stdin:
    try:
        e = json.loads(line)
    except ValueError:
        print(line, end="")  # build errors arrive as plain text
        continue
    key = (e.get("Package"), e.get("Test"))
    a = e.get("Action")
    if a == "output":
        out.setdefault(key, []).append(e.get("Output", ""))
    elif a in ("pass", "fail", "skip"):
        if e.get("Test"):
            if a == "pass":
                passed += 1
            elif a == "fail":
                failed.append(key)
                sys.stdout.write("".join(out.get(key, [])))
            out.pop(key, None)
        else:
            text = out.pop(key, [])
            summary = [l for l in text if l.startswith(("ok", "FAIL", "?", "---"))]
            if a == "fail" and not any(k[0] == key[0] for k in failed):
                sys.stdout.write("".join(text))  # a package failing outside a test (a panic, a build)
            else:
                sys.stdout.write("".join(summary) or f"{a} {key[0]}\n")
for k in failed:
    print(f"FAILED {label} {k[0]} {k[1]}")
print(f"{label}: {passed} passed, {len(failed)} failed")
sys.exit(1 if failed else 0)
' "$1"
}

total=0
rc=0
work=$(mktemp -d)
pids=()
labels=()
# A stage still running when this ends (an interrupt, a timeout) is stopped.
stop() {
  local p
  for p in "${pids[@]:-}"; do
    # (set -e is on here: a stage that has ended is no error)
    if [ -n "$p" ]; then
      pkill -P "$p" 2>/dev/null || true
      kill "$p" 2>/dev/null || true
    fi
  done
  rm -rf "$work"
}
trap stop EXIT
trap 'exit 130' INT TERM
# start <label> <go test args…>: one go test in the background, its output kept
# for show, in the order the stages were started. The stages share no state
# (go's build cache takes concurrent builds, each test works in temp dirs), and
# the longest of them, not their sum, is the wall time: check full took 125 to
# 137 s with the stages one after the other (57 s plain, 4 s stream, 66 s
# -race), the 110 s bar's reason (fuzzer-long's audit, f3.3).
start() {
  local label=$1; shift
  echo "== $label: go test $*"
  (
    set +e
    go test -json "$@" | summarize "$label" > "$work/$label.out"
    echo "${PIPESTATUS[1]}" > "$work/$label.rc"
  ) &
  pids+=($!)
  labels+=("$label")
}
# show: wait for the stages, print each one's output in the order started, and
# count its tests.
show() {
  local i label n
  for i in "${!pids[@]}"; do
    label=${labels[$i]}
    wait "${pids[$i]}" || true
    cat "$work/$label.out"
    n=$(tail -1 "$work/$label.out" | sed -n 's/^.*: \([0-9]*\) passed.*/\1/p')
    total=$((total + ${n:-0}))
    [ "$(cat "$work/$label.rc")" = "0" ] || rc=1
  done
}

case $mode in
  fast)    # go's own test cache stays on: a working loop reruns only what changed
    go vet ./...
    start plain ./...
    show
    ;;
  full)    # AGENTS.md's checks before calling work done, on every package
    go vet ./...
    start plain -count=1 ./...
    SHEETS_STREAM=always start stream -count=1 $stream_pkgs
    start race -race -count=1 -timeout 30m ./...
    show
    ;;
  *) echo "usage: scripts/check.sh fast|full" >&2; exit 2 ;;
esac

if [ "$rc" -eq 0 ]; then
  echo "check: $total tests passed"
else
  echo "check: failures above"
fi
exit $rc
