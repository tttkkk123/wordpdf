@echo off
setlocal
cd /d "%~dp0"
title Word-PDF Converter

set "PY=%~dp0.venv\Scripts\python.exe"
set "PYW=%~dp0.venv\Scripts\pythonw.exe"

if /i "%~1"=="/debug" goto :debug
if not exist "%PY%" goto :setup
if not exist "%PYW%" set "PYW=%PY%"
goto :run

:setup
echo.
echo   [1/3] Creating virtual environment ...
set "BASEPY="
where py >nul 2>nul
if not errorlevel 1 set "BASEPY=py -3"
if not defined BASEPY (
    where python >nul 2>nul
    if not errorlevel 1 set "BASEPY=python"
)
if not defined BASEPY goto :nopython
%BASEPY% -m venv .venv
if errorlevel 1 goto :fail
echo   [2/3] Installing dependencies (first run: 1-3 minutes) ...
"%PY%" -m pip install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple >nul 2>nul
"%PY%" -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 goto :fail
echo   [3/3] Ready. Starting ...
if not exist "%PYW%" set "PYW=%PY%"
goto :run

:run
"%PY%" -c "import tkinter, docx" >nul 2>nul
if errorlevel 1 goto :broken
start "" "%PYW%" "%~dp0run_gui.pyw"
exit /b 0

:debug
"%PY%" -m wordpdf.gui
pause
exit /b 0

:broken
echo.
echo   [!] Environment incomplete: tkinter or python-docx is missing.
echo       Delete the .venv folder and run this file again.
echo       If it still fails, use the official Python from python.org.
echo.
pause
exit /b 1

:nopython
echo.
echo   [!] Python not found. Please install Python 3.9+ first
echo       and check "Add python.exe to PATH" during setup:
echo       https://www.python.org/downloads/
echo.
pause
exit /b 1

:fail
echo.
echo   [!] Setup failed. Please send the messages above for help.
echo.
pause
exit /b 1
