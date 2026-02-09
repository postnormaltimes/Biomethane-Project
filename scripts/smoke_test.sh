
#!/usr/bin/env bash
set -euo pipefail

root_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$root_dir"

export PYTHONPATH="${PYTHONPATH:-}:$root_dir/src"

if command -v dcf >/dev/null 2>&1; then
  dcf_cmd=(dcf)
else
  dcf_cmd=(python -m dcf_ui_cli.cli)
fi

printf "\n==> ${dcf_cmd[*]} --help\n"
"${dcf_cmd[@]}" --help >/dev/null

printf "\n==> ${dcf_cmd[*]} run --input examples/example_input.yaml\n"
"${dcf_cmd[@]}" run --input examples/example_input.yaml >/dev/null

printf "\n==> ${dcf_cmd[*]} biometano report --input case_files/biometano_case.yaml\n"
"${dcf_cmd[@]}" biometano report --input case_files/biometano_case.yaml >/dev/null

printf "\n==> ${dcf_cmd[*]} biometano sens --input case_files/biometano_case.yaml\n"
"${dcf_cmd[@]}" biometano sens --input case_files/biometano_case.yaml >/dev/null

echo "\nSmoke test completed successfully."
