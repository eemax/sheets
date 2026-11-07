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
# run <label> <go test args…>: one go test, counted.
run() {
  local label=$1; shift
  echo "== $label: go test $*"
  local log; log=$(mktemp)
  set +e
  go test -json "$@" | summarize "$label" | tee "$log"
  local r=${PIPESTATUS[1]}
  set -e
  local n; n=$(tail -1 "$log" | sed -n 's/^.*: \([0-9]*\) passed.*/\1/p')
  rm -f "$log"
  total=$((total + ${n:-0}))
  [ "$r" -eq 0 ] || rc=1
}

case $mode in
  fast)    # go's own test cache stays on: a working loop reruns only what changed
    go vet ./...
    run plain ./...
    ;;
  full)    # AGENTS.md's checks before calling work done, on every package
    go vet ./...
    run plain -count=1 ./...
    SHEETS_STREAM=always run stream -count=1 $stream_pkgs
    run race -race -count=1 -timeout 30m ./...
    ;;
  *) echo "usage: scripts/check.sh fast|full" >&2; exit 2 ;;
esac

if [ "$rc" -eq 0 ]; then
  echo "check: $total tests passed"
else
  echo "check: failures above"
fi
exit $rc
