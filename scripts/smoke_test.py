from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


def _run(cmd: list[str], env: dict[str, str]) -> None:
    result = subprocess.run(
        cmd,
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        env=env,
    )
    if result.returncode != 0:
        print(result.stdout)
        raise SystemExit(result.returncode)


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    env = os.environ.copy()
    env["PYTHONPATH"] = str(root / "src")

    _run([sys.executable, "-m", "dcf_ui_cli.cli", "--help"], env)
    _run(
        [sys.executable, "-m", "dcf_ui_cli.cli", "run", "--input", "examples/example_input.yaml", "--quiet"],
        env,
    )
    _run(
        [
            sys.executable,
            "-m",
            "dcf_ui_cli.cli",
            "biometano",
            "report",
            "--input",
            "case_files/biometano_case.yaml",
            "--output",
            "output/biometano_smoke",
            "--xlsx-mode",
            "formulas",
        ],
        env,
    )
    _run(
        [
            sys.executable,
            "-m",
            "dcf_ui_cli.cli",
            "export",
            "--input",
            "case_files/biometano_case.yaml",
            "--output",
            "output/biometano_smoke/biometano_export.xlsx",
            "--format",
            "xlsx",
            "--xlsx-mode",
            "formulas",
        ],
        env,
    )


if __name__ == "__main__":
    main()
