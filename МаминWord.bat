@echo off
rem Only ASCII here: cmd.exe misreads .bat files that contain Cyrillic text.
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
  echo Python not found. Install Python 3 from https://www.python.org
  echo and tick "Add python.exe to PATH" during installation.
  pause
  exit /b 1
)
python launcher.py
pause
