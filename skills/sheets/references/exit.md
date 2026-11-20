# Exit status

0 on success. 1 on errors (a message goes to stderr, or `{"error": ...}`
with `--json`), when `fmt --check` finds changes, when `check` or `lint`
finds problems at `--fail-on` level, when `check --against` finds values
unexplained or features lost, and when `diff --exit-code` finds
differences.
