"""End-to-end packaging, clean installation, doctor, and Streamlit runtime smoke tests."""

from __future__ import annotations

import subprocess
import sys
import tomllib
from pathlib import Path

from click.testing import CliRunner
from streamlit.testing.v1 import AppTest

from ada_platform.cli.main import cli
import ada_platform.dashboard.app as dashboard_app

REPO_ROOT = Path(__file__).resolve().parents[1]


class TestCleanInstallAndRuntime:
    def test_pyproject_declares_all_runtime_dependencies(self):
        """Every third-party runtime dependency must be declared in pyproject.toml."""
        pyproject_path = REPO_ROOT / "pyproject.toml"
        data = tomllib.loads(pyproject_path.read_text(encoding="utf-8"))
        deps = [d.lower() for d in data["project"]["dependencies"]]

        for required_pkg in (
            "python-docx",
            "lxml",
            "openpyxl",
            "click",
            "streamlit",
            "pymupdf",
        ):
            assert any(
                d.startswith(required_pkg) for d in deps
            ), f"Missing runtime dependency '{required_pkg}' in pyproject.toml"

        scripts = data["project"].get("scripts", {})
        assert scripts.get("ada-platform") == "ada_platform.cli.main:cli"

    def test_wheel_builds_and_installs_into_isolated_venv(self, tmp_path):
        """Build a wheel, install into a brand-new isolated venv, and run `ada-platform doctor`."""
        import shutil

        uv_bin = shutil.which("uv")
        assert uv_bin is not None, "uv must be available to test clean venv installation"

        wheel_dir = tmp_path / "dist"
        wheel_dir.mkdir()
        clean_venv = tmp_path / "clean_venv"

        # 1. Build wheel from source (offline-first with online fallback for fresh CI sandboxes)
        build_cmd = [
            uv_bin,
            "build",
            "--wheel",
            "--offline",
            "--out-dir",
            str(wheel_dir),
            str(REPO_ROOT),
        ]
        build_res = subprocess.run(
            build_cmd,
            capture_output=True,
            text=True,
            check=False,
        )
        if build_res.returncode != 0:
            build_res = subprocess.run(
                [c for c in build_cmd if c != "--offline"],
                capture_output=True,
                text=True,
                check=False,
            )
        assert build_res.returncode == 0, f"Wheel build failed:\n{build_res.stderr}"
        wheels = list(wheel_dir.glob("ada_platform-*.whl"))
        assert len(wheels) == 1

        # 2. Create a clean virtualenv and install the wheel
        venv_res = subprocess.run(
            [uv_bin, "venv", str(clean_venv)],
            capture_output=True,
            text=True,
            check=False,
        )
        assert venv_res.returncode == 0, f"venv creation failed:\n{venv_res.stderr}"

        if sys.platform == "win32":
            clean_py = clean_venv / "Scripts" / "python.exe"
            installed_cli = clean_venv / "Scripts" / "ada-platform.exe"
        else:
            clean_py = clean_venv / "bin" / "python"
            installed_cli = clean_venv / "bin" / "ada-platform"

        install_cmd = [
            uv_bin,
            "pip",
            "install",
            "--python",
            str(clean_py),
            "--offline",
            str(wheels[0]),
        ]
        install_res = subprocess.run(
            install_cmd,
            capture_output=True,
            text=True,
            check=False,
        )
        if install_res.returncode != 0:
            install_res = subprocess.run(
                [c for c in install_cmd if c != "--offline"],
                capture_output=True,
                text=True,
                check=False,
            )
        assert install_res.returncode == 0, f"Wheel install failed:\n{install_res.stderr}"

        # 3. Run the installed `ada-platform doctor` and `ada-platform ui --check` from tmp_path (outside repo)
        assert installed_cli.is_file()
        doc_res = subprocess.run(
            [str(installed_cli), "doctor"],
            cwd=str(tmp_path),
            capture_output=True,
            text=True,
            check=False,
        )
        assert doc_res.returncode == 0, f"doctor failed:\n{doc_res.stdout}\n{doc_res.stderr}"
        assert "RESULT: READY" in doc_res.stdout

        ui_res = subprocess.run(
            [str(installed_cli), "ui", "--check"],
            cwd=str(tmp_path),
            capture_output=True,
            text=True,
            check=False,
        )
        assert ui_res.returncode == 0, f"clean-venv ui --check failed:\n{ui_res.stdout}\n{ui_res.stderr}"
        assert "UI_CHECK_PASS" in ui_res.stdout

    def test_doctor_cli_command(self):
        """`ada-platform doctor` verifies Python modules, UI script, and lualatex."""
        runner = CliRunner()
        res = runner.invoke(cli, ["doctor"])
        assert res.exit_code == 0, res.output
        assert "RESULT: READY" in res.output

    def test_doctor_cli_windows_mode(self, monkeypatch):
        """When running on Windows (sys.platform == 'win32'), `doctor` checks PowerShell COM readiness."""
        import shutil
        import ada_platform.cli.main as cli_main

        monkeypatch.setattr(cli_main.sys, "platform", "win32")
        monkeypatch.setattr(
            shutil,
            "which",
            lambda cmd: (
                r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
                if cmd == "powershell"
                else None
            ),
        )
        runner = CliRunner()
        res = runner.invoke(cli, ["doctor"])
        assert res.exit_code == 0, res.output
        assert "Windows Microsoft Word COM via PowerShell" in res.output
        assert "RESULT: READY" in res.output

        # Also verify fail-closed behavior if PowerShell is missing on Windows
        monkeypatch.setattr(shutil, "which", lambda _cmd: None)
        res_fail = runner.invoke(cli, ["doctor"])
        assert res_fail.exit_code == 1
        assert "RESULT: FAILED" in res_fail.output

    def test_streamlit_dashboard_renders_cleanly_via_apptest(self):
        """Headless AppTest verifies `app.py` renders initial UI with zero uncaught exceptions."""
        app_file = Path(dashboard_app.__file__).resolve()
        at = AppTest.from_file(str(app_file), default_timeout=15)
        at.run()
        assert not at.exception, f"Streamlit AppTest raised exception: {at.exception}"
        assert len(at.title) == 1
        assert "ADA Document & Tagged PDF Compliance Tool" in at.title[0].value

    def test_cli_ui_check_boots_real_http_streamlit_server(self):
        """`ada-platform ui --check` boots a real Streamlit HTTP server and checks /_stcore/health."""
        runner = CliRunner()
        res = runner.invoke(cli, ["ui", "--check"])
        assert res.exit_code == 0, res.output
        assert "UI_CHECK_PASS" in res.output
