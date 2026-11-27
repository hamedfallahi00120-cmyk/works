@echo off
setlocal enabledelayedexpansion
rem Copies this update over an existing Local-Office-AI project.
rem Usage: Deploy_Update.bat F:\Local-Office-AI [--test]
rem Copies code, tests, docs and launchers only.
rem Never touches or overwrites venv, data, logs, reports, reference or personal files.

set "SRC=%~dp0"
set "DST=%~1"
if "%DST%"=="" set "DST=F:\Local-Office-AI"
for %%I in ("%DST%") do set "DST=%%~fI"
if "%DST:~-1%"=="\" set "DST=%DST:~0,-1%"

if not exist "%DST%" (
    echo [ERROR] project folder not found: %DST%
    exit /b 1
)
if not exist "%DST%\venv\Scripts\python.exe" (
    echo [WARN] venv not found at %DST%\venv\Scripts\python.exe
    echo        the update still copies code, but the project may not run until venv exists.
)

echo Source: %SRC%
echo Target: %DST%
echo.

rem ---- update this updater on the target (unless already running from it) --
if /I not "%~dp0"=="%DST%\" (
    copy /Y "%~f0" "%DST%\Deploy_Update.bat" >nul
    if errorlevel 1 (
        echo [WARN] could not replace Deploy_Update.bat on the target.
        echo        Copy it manually before running the new deployment tests.
    )
)

rem ---- code folders (never user data) -------------------------------------
for %%D in (src tests docs legacy) do (
    if exist "%SRC%%%D" (
        echo [+] folder %%D
        xcopy "%SRC%%%D" "%DST%\%%D" /E /I /Y /Q >nul
    )
)

rem ---- top-level code, config templates and launchers ---------------------
for %%F in (run_engine.py run_ui.py pytest.ini requirements.txt requirements-documents.txt README.md AGENTS.md DEPLOY.md .gitignore) do (
    if exist "%SRC%%%F" (
        echo [+] file %%F
        copy /Y "%SRC%%%F" "%DST%\%%F" >nul
    )
)
for %%F in (Run_Local_Office_AI.bat Run_Local_Office_AI_GUI.bat Run_Local_Office_AI_Modern.bat Run_Tests.bat) do (
    if exist "%SRC%%%F" (
        echo [+] launcher %%F
        copy /Y "%SRC%%%F" "%DST%\%%F" >nul
    )
)

rem ---- config defaults: create only when missing, never overwrite user values
if exist "%SRC%config" (
    if not exist "%DST%\config" mkdir "%DST%\config"
    for %%C in (settings.json synonyms.json ui_settings.json) do (
        if exist "%SRC%config\%%C" (
            if not exist "%DST%\config\%%C" (
                echo [+] new config %%C
                copy /Y "%SRC%config\%%C" "%DST%\config\%%C" >nul
            ) else (
                echo [=] kept existing config %%C
            )
        )
    )
)

echo.
echo Skipped by design: venv, data\input, data\output, data\processed, logs,
echo reports, reference and every personal file already on the target.
echo Nothing was deleted.
echo.

rem ---- post-copy verification: the modern entry point must exist ----------
set "MISSING=0"
for %%F in (run_engine.py run_ui.py Run_Local_Office_AI_Modern.bat src\local_office_ai\ui\desktop\app.py src\local_office_ai\ui\desktop\main_window.py) do (
    if not exist "%DST%\%%F" (
        echo [ERROR] missing after copy: %%F
        set "MISSING=1"
    )
)
if "!MISSING!"=="1" (
    echo [ERROR] deployment is incomplete; fix the manifest and run again.
    exit /b 1
)
echo Entry points verified: run_ui.py and Run_Local_Office_AI_Modern.bat.

if /I "%~2"=="--test" (
    echo Running test suite with the project venv...
    "%DST%\venv\Scripts\python.exe" -c "import pytest" >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] pytest is not installed in the target venv; tests were not run.
        echo         Daily use does not need pytest. Install it only if you want tests.
        exit /b 2
    )
    pushd "%DST%"
    venv\Scripts\python.exe -m pytest tests -o addopts="" -q
    set "RC=!ERRORLEVEL!"
    popd
    if not "!RC!"=="0" (
        echo [ERROR] tests failed with code !RC!
        exit /b !RC!
    )
    echo Tests passed.
)

exit /b 0
