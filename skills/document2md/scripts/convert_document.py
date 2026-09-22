#!/usr/bin/env python3
"""Entry point for document2md.

On first use this creates a virtual environment inside the skill folder
(`<skill>/.venv`), installs the requirements into it, and re-executes itself
with that interpreter. Every later run finds the same environment and starts
straight away, so callers never have to locate or activate it. Pass
``--no-bootstrap`` to stay in the current interpreter.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPTS_DIR.parent
VENV_DIR = SKILL_DIR / ".venv"
REQUIREMENTS = SCRIPTS_DIR / "requirements.txt"
GUARD = "DOCUMENT2MD_BOOTSTRAPPED"
REQUIRED_MODULES = ("docx", "pptx", "openpyxl", "docling", "pypdfium2")


def _venv_python(venv_dir: Path) -> Path:
    if os.name == "nt":
        return venv_dir / "Scripts" / "python.exe"
    return venv_dir / "bin" / "python"


def _dependencies_available() -> bool:
    from importlib.util import find_spec

    try:
        return all(find_spec(name) is not None for name in REQUIRED_MODULES)
    except (ImportError, ValueError):
        return False


def _create_environment() -> None:
    import venv

    print(f"document2md: creating {VENV_DIR} (first run only)...", file=sys.stderr)
    venv.EnvBuilder(with_pip=True, symlinks=os.name != "nt").create(VENV_DIR)
    subprocess.run(
        [str(_venv_python(VENV_DIR)), "-m", "pip", "install", "--quiet",
         "--upgrade", "pip"],
        check=True,
    )
    print("document2md: installing dependencies...", file=sys.stderr)
    subprocess.run(
        [str(_venv_python(VENV_DIR)), "-m", "pip", "install", "--quiet",
         "-r", str(REQUIREMENTS)],
        check=True,
    )


def _bootstrap(argv: list[str]) -> None:
    """Re-execute in the skill's own environment, creating it when missing."""
    interpreter = _venv_python(VENV_DIR)
    if not interpreter.exists():
        try:
            _create_environment()
        except (OSError, subprocess.CalledProcessError) as error:
            print(
                f"ERROR: could not create {VENV_DIR}: {error}\n"
                f"Install {REQUIREMENTS} into an environment of your own and rerun "
                "with --no-bootstrap.",
                file=sys.stderr,
            )
            raise SystemExit(2)
    environment = dict(os.environ, **{GUARD: "1"})
    completed = subprocess.run([str(interpreter), str(Path(__file__).resolve()), *argv], env=environment)
    _exit(completed.returncode)


def _exit(code: int) -> None:
    """Leave immediately, after flushing.

    Native OCR and inference libraries can abort during interpreter shutdown
    (`recursive_mutex lock failed`), which would turn a finished conversion
    into a crash exit code. The output is already on disk at this point.
    """
    sys.stdout.flush()
    sys.stderr.flush()
    os._exit(code)


def main() -> None:
    argv = sys.argv[1:]
    if "--no-bootstrap" not in argv and not os.environ.get(GUARD) and not _dependencies_available():
        _bootstrap(argv)

    sys.path.insert(0, str(SCRIPTS_DIR))
    from document2md.cli import main as cli_main

    _exit(cli_main(argv))


if __name__ == "__main__":
    main()
