# Exit status

0 on success. 1 on errors (a message goes to stderr, or `{"error": ...}`
with `--json`), when `fmt --check` finds changes, when `lint` finds
problems at `--fail-on` level, and when `diff --exit-code` finds
differences.
