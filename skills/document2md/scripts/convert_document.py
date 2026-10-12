#!/usr/bin/env python3
"""Entry point for document2md.

On first use this creates a virtual environment inside the skill folder
(`<skill>/.venv`), installs the requirements into it, and re-executes itself
with that interpreter. Every later run finds the same environment and starts
straight away, so callers never have to locate or activate it. Pass
``--no-bootstrap`` to stay in the current interpreter.

The optional Chandra PDF backend (``--pdf-backend chandra``) is installed into
that same environment the first time it is requested, so the default install
does not carry it.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPTS_DIR.parent
VENV_DIR = SKILL_DIR / ".venv"
REQUIREMENTS = SCRIPTS_DIR / "requirements.txt"
CHANDRA_REQUIREMENTS = SCRIPTS_DIR / "requirements-chandra.txt"
GUARD = "DOCUMENT2MD_BOOTSTRAPPED"
REQUIRED_MODULES = ("docx", "pptx", "openpyxl", "docling", "pypdfium2")
MIN_PYTHON = (3, 11)
CANDIDATE_PYTHONS = ("python3.12", "python3.13", "python3.11", "python3.14", "python3", "python")


def _venv_python(venv_dir: Path) -> Path:
    if os.name == "nt":
        return venv_dir / "Scripts" / "python.exe"
    return venv_dir / "bin" / "python"


def _python_version(interpreter: str | Path) -> tuple[int, int] | None:
    try:
        completed = subprocess.run(
            [str(interpreter), "-c", "import sys; print(*sys.version_info[:2])"],
            capture_output=True,
            text=True,
            timeout=30,
        )
    except (OSError, subprocess.SubprocessError):
        return None
    try:
        major, minor = (completed.stdout or "").split()
        return int(major), int(minor)
    except ValueError:
        return None


def _find_base_python() -> str | None:
    """Return an interpreter new enough to build the environment."""
    if sys.version_info[:2] >= MIN_PYTHON:
        return sys.executable
    for name in CANDIDATE_PYTHONS:
        found = shutil.which(name)
        if found:
            version = _python_version(found)
            if version and version >= MIN_PYTHON:
                return found
    uv = shutil.which("uv")
    if uv:
        wanted = ">={}.{}".format(*MIN_PYTHON)
        completed = subprocess.run([uv, "python", "find", wanted], capture_output=True, text=True)
        found = completed.stdout.strip()
        if completed.returncode == 0 and found:
            return found
    return None


def _dependencies_available() -> bool:
    from importlib.util import find_spec

    if sys.version_info[:2] < MIN_PYTHON:
        return False

    try:
        return all(find_spec(name) is not None for name in REQUIRED_MODULES)
    except (ImportError, ValueError):
        return False


def _create_environment() -> None:
    base = _find_base_python()
    if base is None:
        raise OSError(
            "document2md needs Python {}.{} or newer and none was found on PATH "
            "(install one, e.g. `uv python install 3.12`).".format(*MIN_PYTHON)
        )
    print(f"document2md: creating {VENV_DIR} with {base} (first run only)...", file=sys.stderr)
    subprocess.run([base, "-m", "venv", str(VENV_DIR)], check=True)
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
    version = _python_version(interpreter) if interpreter.exists() else None
    if version is not None and version < MIN_PYTHON:
        # An earlier run built the environment with a Python that is too old.
        print(f"document2md: rebuilding {VENV_DIR}, which uses an unsupported Python...", file=sys.stderr)
        shutil.rmtree(VENV_DIR)
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


def _wants_chandra(argv: list[str]) -> bool:
    for index, arg in enumerate(argv):
        if arg == "--pdf-backend=chandra":
            return True
        if arg == "--pdf-backend" and argv[index + 1:index + 2] == ["chandra"]:
            return True
    return False


def _in_skill_environment() -> bool:
    try:
        return Path(sys.prefix).resolve() == VENV_DIR.resolve()
    except OSError:
        return False


def _ensure_chandra(argv: list[str]) -> None:
    """Install the optional Chandra backend into the skill environment on first use."""
    from importlib.util import find_spec

    if "--no-bootstrap" in argv or not _wants_chandra(argv):
        return
    if find_spec("chandra") is not None or not _in_skill_environment():
        return
    print("document2md: installing the Chandra PDF backend (first use only)...", file=sys.stderr)
    completed = subprocess.run(
        [sys.executable, "-m", "pip", "install", "--quiet", "-r", str(CHANDRA_REQUIREMENTS)]
    )
    if completed.returncode != 0:
        print(
            f"ERROR: could not install {CHANDRA_REQUIREMENTS} into {VENV_DIR}.",
            file=sys.stderr,
        )
        _exit(2)


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

    if sys.version_info[:2] < MIN_PYTHON:
        print(
            "ERROR: document2md needs Python {}.{} or newer; this is {}.{}.".format(
                *MIN_PYTHON, *sys.version_info[:2]
            ),
            file=sys.stderr,
        )
        _exit(2)

    _ensure_chandra(argv)
    sys.path.insert(0, str(SCRIPTS_DIR))
    from document2md.cli import main as cli_main

    _exit(cli_main(argv))


if __name__ == "__main__":
    main()
