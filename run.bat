@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title School Management System - Server

echo.
echo ============================================================
echo  School Management System
echo ============================================================
echo  Folder: %CD%
echo.

where python >nul 2>&1
if errorlevel 1 (
  echo Python was not found in PATH.
  echo Please install Python 3 and enable "Add Python to PATH".
  pause
  exit /b 1
)

REM Stop only the process currently listening on TCP 8000.
for /f "tokens=5" %%P in ('netstat -ano ^| findstr /R /C:":8000 .*LISTENING"') do (
  echo Stopping old process on port 8000: PID %%P
  taskkill /PID %%P /F >nul 2>&1
)

REM Repair the database using the script in THIS folder, never System32.
python "%~dp0repair_login.py"
if errorlevel 1 (
  echo.
  echo Login repair failed.
  pause
  exit /b 1
)

echo.
echo Server is starting. Other computers on the SAME WiFi/LAN can use:
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4 Address"') do echo    http://%%a:8000/
echo This computer: http://127.0.0.1:8000/
echo.
echo Keep this window open. Closing it stops the server.
echo.

python -m uvicorn main:app --host 0.0.0.0 --port 8000
if errorlevel 1 (
  echo.
  echo Server stopped or could not start.
  echo Check whether another program is using port 8000.
  pause
)
endlocal
