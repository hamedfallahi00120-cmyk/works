@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "VENV_PY=%~dp0venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo.
  echo [ERROR] Project virtual environment Python was not found:
  echo         %VENV_PY%
  echo.
  echo This launcher never uses the global "py" or a system Python.
  echo Restore the existing F:\Local-Office-AI\venv or point VENV_PY to it.
  echo.
  pause
  exit /b 1
)

echo ============================================================
echo  Local Office AI - Excel Intelligence Engine
echo  Project root: %~dp0
echo  Python:       %VENV_PY%
echo ============================================================
echo.

"%VENV_PY%" -u "%~dp0run_engine.py" %*
set "EXITCODE=%ERRORLEVEL%"

echo.
if not "%EXITCODE%"=="0" (
  echo Program finished with errors. Exit code: %EXITCODE%
  echo See the log folder: %~dp0logs
) else (
  echo Program finished successfully.
)
echo.
pause
exit /b %EXITCODE%
