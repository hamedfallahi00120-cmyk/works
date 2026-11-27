@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ================================================
echo Local Office Assistant V4.2
echo Excel Audit + GPU Monitor + ComfyUI Image Studio
echo ================================================
where py >nul 2>nul
if errorlevel 1 (
  echo Python launcher was not found. Install Python 3.x from python.org.
  pause
  exit /b 1
)
py -m pip show openpyxl >nul 2>nul
if errorlevel 1 (
  echo Installing openpyxl...
  py -m pip install openpyxl
  if errorlevel 1 goto :fail
)
py -m pip show psutil >nul 2>nul
if errorlevel 1 (
  echo Installing psutil...
  py -m pip install psutil
  if errorlevel 1 goto :fail
)
py -u "%~dp0Ollama_Excel_Assistant_V4_2.py"
echo.
echo V4.2 closed. Press any key to close this window.
pause >nul
exit /b 0
:fail
echo A required Python package could not be installed.
pause
exit /b 1
