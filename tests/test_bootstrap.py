"""The entry point manages its own environment.

The first run creates `<skill>/.venv` and re-executes into it, so later callers
never have to know where the environment lives.
"""

from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path

import pytest


ROOT = Path(__file__).resolve().parents[1]
SKILL_SCRIPT = ROOT / "skills" / "document2md" / "scripts" / "convert_document.py"


@pytest.fixture()
def entry_point():
    spec = importlib.util.spec_from_file_location("document2md_entry", SKILL_SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_venv_lives_inside_the_skill_folder(entry_point) -> None:
    assert entry_point.VENV_DIR == SKILL_SCRIPT.parent.parent / ".venv"
    assert entry_point.VENV_DIR.name == ".venv"
    assert entry_point._venv_python(Path("/tmp/env")).parent.name in {"bin", "Scripts"}


def test_first_run_creates_the_environment_then_re_executes(entry_point, monkeypatch, tmp_path) -> None:
    created = []
    launched = []

    monkeypatch.setattr(entry_point, "VENV_DIR", tmp_path / ".venv")
    monkeypatch.setattr(entry_point, "_dependencies_available", lambda: False)
    monkeypatch.setattr(entry_point, "_create_environment", lambda: created.append(True))
    monkeypatch.setattr(entry_point, "_exit", lambda code: (_ for _ in ()).throw(SystemExit(code)))
    monkeypatch.setattr(
        entry_point.subprocess,
        "run",
        lambda command, env=None, **kwargs: launched.append((command, env)) or subprocess.CompletedProcess(command, 0),
    )
    monkeypatch.setattr(sys, "argv", ["convert_document.py", "file.pdf"])
    monkeypatch.delenv(entry_point.GUARD, raising=False)

    with pytest.raises(SystemExit) as exit_info:
        entry_point.main()

    assert exit_info.value.code == 0
    assert created == [True]
    command, env = launched[0]
    assert command[0] == str(entry_point._venv_python(tmp_path / ".venv"))
    assert command[-1] == "file.pdf"
    assert env[entry_point.GUARD] == "1"


def test_an_existing_environment_is_reused_without_reinstalling(entry_point, monkeypatch, tmp_path) -> None:
    interpreter = entry_point._venv_python(tmp_path / ".venv")
    interpreter.parent.mkdir(parents=True)
    interpreter.write_text("#!/bin/sh\n")

    def fail() -> None:
        raise AssertionError("the environment must not be rebuilt when it already exists")

    monkeypatch.setattr(entry_point, "VENV_DIR", tmp_path / ".venv")
    monkeypatch.setattr(entry_point, "_dependencies_available", lambda: False)
    monkeypatch.setattr(entry_point, "_create_environment", fail)
    monkeypatch.setattr(entry_point, "_exit", lambda code: (_ for _ in ()).throw(SystemExit(code)))
    monkeypatch.setattr(
        entry_point.subprocess,
        "run",
        lambda command, env=None, **kwargs: subprocess.CompletedProcess(command, 3),
    )
    monkeypatch.setattr(sys, "argv", ["convert_document.py", "file.pdf"])
    monkeypatch.delenv(entry_point.GUARD, raising=False)

    with pytest.raises(SystemExit) as exit_info:
        entry_point.main()
    assert exit_info.value.code == 3


def test_bootstrapping_is_skipped_when_the_dependencies_are_already_importable(
    entry_point, monkeypatch
) -> None:
    def fail(argv) -> None:
        raise AssertionError("must not bootstrap when the current interpreter already works")

    monkeypatch.setattr(entry_point, "_dependencies_available", lambda: True)
    monkeypatch.setattr(entry_point, "_bootstrap", fail)
    monkeypatch.setattr(entry_point, "_exit", lambda code: (_ for _ in ()).throw(SystemExit(code)))
    monkeypatch.setattr(sys, "argv", ["convert_document.py", "--help"])

    with pytest.raises(SystemExit):
        entry_point.main()


def test_no_bootstrap_flag_keeps_the_current_interpreter(entry_point, monkeypatch) -> None:
    def fail(argv) -> None:
        raise AssertionError("--no-bootstrap must never re-execute")

    monkeypatch.setattr(entry_point, "_dependencies_available", lambda: False)
    monkeypatch.setattr(entry_point, "_bootstrap", fail)
    monkeypatch.setattr(entry_point, "_exit", lambda code: (_ for _ in ()).throw(SystemExit(code)))
    monkeypatch.setattr(sys, "argv", ["convert_document.py", "--no-bootstrap", "--help"])

    with pytest.raises(SystemExit):
        entry_point.main()


def test_an_environment_built_with_an_old_python_is_rebuilt(entry_point, monkeypatch, tmp_path) -> None:
    venv_dir = tmp_path / ".venv"
    interpreter = entry_point._venv_python(venv_dir)
    interpreter.parent.mkdir(parents=True)
    interpreter.write_text("#!/bin/sh\n")
    created = []

    def create() -> None:
        assert not venv_dir.exists(), "the stale environment must be removed first"
        interpreter.parent.mkdir(parents=True)
        interpreter.write_text("#!/bin/sh\n")
        created.append(True)

    monkeypatch.setattr(entry_point, "VENV_DIR", venv_dir)
    monkeypatch.setattr(entry_point, "_python_version", lambda path: (3, 9))
    monkeypatch.setattr(entry_point, "_create_environment", create)
    monkeypatch.setattr(entry_point, "_exit", lambda code: (_ for _ in ()).throw(SystemExit(code)))
    monkeypatch.setattr(
        entry_point.subprocess,
        "run",
        lambda command, env=None, **kwargs: subprocess.CompletedProcess(command, 0),
    )

    with pytest.raises(SystemExit):
        entry_point._bootstrap(["file.pdf"])
    assert created == [True]


def test_an_environment_whose_version_cannot_be_read_is_kept(entry_point, monkeypatch, tmp_path) -> None:
    venv_dir = tmp_path / ".venv"
    interpreter = entry_point._venv_python(venv_dir)
    interpreter.parent.mkdir(parents=True)
    interpreter.write_text("#!/bin/sh\n")

    def fail() -> None:
        raise AssertionError("an environment must not be deleted on an unreadable version")

    monkeypatch.setattr(entry_point, "VENV_DIR", venv_dir)
    monkeypatch.setattr(entry_point, "_python_version", lambda path: None)
    monkeypatch.setattr(entry_point, "_create_environment", fail)
    monkeypatch.setattr(entry_point, "_exit", lambda code: (_ for _ in ()).throw(SystemExit(code)))
    monkeypatch.setattr(
        entry_point.subprocess,
        "run",
        lambda command, env=None, **kwargs: subprocess.CompletedProcess(command, 0),
    )

    with pytest.raises(SystemExit):
        entry_point._bootstrap(["file.pdf"])
    assert interpreter.exists()


def test_an_old_interpreter_looks_for_a_newer_python_on_path(entry_point, monkeypatch) -> None:
    versions = {"/opt/py39": (3, 9), "/opt/py312": (3, 12)}
    on_path = {"python3.12": "/opt/py312", "python3": "/opt/py39"}

    monkeypatch.setattr(sys, "version_info", (3, 9, 6, "final", 0))
    monkeypatch.setattr(entry_point.shutil, "which", lambda name: on_path.get(name))
    monkeypatch.setattr(entry_point, "_python_version", lambda path: versions.get(path))

    assert entry_point._find_base_python() == "/opt/py312"
    assert entry_point._dependencies_available() is False


def test_a_new_enough_interpreter_builds_the_environment_itself(entry_point, monkeypatch) -> None:
    monkeypatch.setattr(sys, "version_info", (3, 12, 0, "final", 0))
    assert entry_point._find_base_python() == sys.executable
