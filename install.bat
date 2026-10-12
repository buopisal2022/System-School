@echo off
cd /d "%~dp0"
title School Management System - Install
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo INSTALL FAILED. Please make sure Python is installed and added to PATH.
  pause
  exit /b 1
)
python "%~dp0schema_setup.py"
python "%~dp0repair_login.py"
if errorlevel 1 pause
pause
