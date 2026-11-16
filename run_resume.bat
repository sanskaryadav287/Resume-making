@echo off
title Sanskar Yadav - 3D Animated Resume
color 0b
echo =============================================================================
echo   SANSKAR YADAV - 3D ANIMATED INTERACTIVE RESUME LAUNCHER
echo   Data Science Student ^| Python ^& ML ^| AWS Cloud Practitioner
echo =============================================================================
echo.

cd /d "%~dp0"

where python >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [*] Launching 3D Resume via Python...
    python launch_resume.py
    goto end
)

where python3.13 >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [*] Launching 3D Resume via Python 3.13...
    python3.13 launch_resume.py
    goto end
)

where py >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    echo [*] Launching 3D Resume via Python Launcher (py)...
    py launch_resume.py
    goto end
)

echo [*] Python not detected in system PATH. Opening 3D Animated Resume directly in browser...
start "" "index.html"

:end
pause
