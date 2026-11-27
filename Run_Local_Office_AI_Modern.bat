@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "VENV_PY=%~dp0venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo [ERROR] Project virtual environment Python was not found:
  echo         %VENV_PY%
  echo This launcher never uses the global "py" or a system Python.
  echo Restore F:\Local-Office-AI\venv or point VENV_PY to it.
  pause
  exit /b 1
)

echo ============================================================
echo  Local Office AI - Modern Desktop Interface (PySide6)
echo  Project root: %~dp0
echo  Python:       %VENV_PY%
echo ============================================================
echo.

"%VENV_PY%" -u "%~dp0run_ui.py" %*
set "EXITCODE=%ERRORLEVEL%"

echo.
if not "%EXITCODE%"=="0" (
  echo Interface closed with errors. Exit code: %EXITCODE%
  echo Check that PySide6 is installed in the project venv and see:
  echo   %~dp0logs\local_office_ai.log
  pause
) 
exit /b %EXITCODE%
