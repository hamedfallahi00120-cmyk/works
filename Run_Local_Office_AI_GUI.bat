@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "VENV_PY=%~dp0venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo [ERROR] Project virtual environment Python was not found:
  echo         %VENV_PY%
  echo This launcher never uses the global "py" or a system Python.
  pause
  exit /b 1
)

"%VENV_PY%" -u "%~dp0run_engine.py" --gui
set "EXITCODE=%ERRORLEVEL%"
if not "%EXITCODE%"=="0" (
  echo.
  echo [ERROR] GUI could not start. Exit code: %EXITCODE%
  echo Check that Tkinter is available in the project venv and see:
  echo   %~dp0logs\local_office_ai.log
  pause
)
exit /b %EXITCODE%
