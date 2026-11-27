@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "VENV_PY=%~dp0venv\Scripts\python.exe"
if not exist "%VENV_PY%" (
  echo [ERROR] Project virtual environment Python was not found:
  echo         %VENV_PY%
  pause
  exit /b 1
)

"%VENV_PY%" -c "import pytest" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] pytest is not installed in the project venv.
  echo         Daily use does not need pytest. To run tests only, install it with:
  echo         venv\Scripts\python.exe -m pip install pytest
  pause
  exit /b 2
)

"%VENV_PY%" -m pytest tests -o addopts="" -q
set "EXITCODE=%ERRORLEVEL%"
echo.
pause
exit /b %EXITCODE%
