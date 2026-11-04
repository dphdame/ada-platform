@echo off
setlocal
cd /d "%~dp0"
if exist "pyproject.toml" (
    set "REPO_ROOT=%CD%"
) else (
    cd /d "%~dp0.."
    set "REPO_ROOT=%CD%"
)

echo ============================================================
echo   ADA Document ^& Tagged PDF Compliance Platform - Installer
echo ============================================================

where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python 3.11+ was not found on your PATH.
    echo Please install Python 3.11 or newer from https://www.python.org/downloads/windows/
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo [1/4] Creating isolated Python virtual environment (.venv)...
    python -m venv .venv
    if %ERRORLEVEL% neq 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
)

echo [2/4] Installing ADA Compliance Platform and dependencies...
if exist "dist" (
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -e .
) else (
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -e .
)
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Package installation failed.
    pause
    exit /b 1
)

echo [3/4] Running environment ^& dependency check (ada-platform doctor)...
".venv\Scripts\ada-platform.exe" doctor
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Environment check failed.
    pause
    exit /b 1
)

echo [4/4] Verifying Streamlit server health (ada-platform ui --check)...
".venv\Scripts\ada-platform.exe" ui --check
if %ERRORLEVEL% neq 0 (
    echo [ERROR] UI server self-check failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Installation Verified! Launching Interactive Web App...
echo ============================================================
".venv\Scripts\ada-platform.exe" ui
pause
