@echo off
chcp 65001 >nul 2>&1
title NordFox Module Manager
setlocal

cd /d "%~dp0"
set "PY_EXE="
set "USING_VENV=0"

if exist ".venv\Scripts\python.exe" (
    .venv\Scripts\python.exe -V >nul 2>&1
    if not errorlevel 1 (
        set "PY_EXE=.venv\Scripts\python.exe"
        set "USING_VENV=1"
    )
)

if not defined PY_EXE (
    where py >nul 2>&1
    if not errorlevel 1 (
        echo Local virtual environment not found. Creating .venv...
        py -3.12 -m venv .venv >nul 2>&1
        if errorlevel 1 py -m venv .venv >nul 2>&1
        if exist ".venv\Scripts\python.exe" (
            set "PY_EXE=.venv\Scripts\python.exe"
            set "USING_VENV=1"
            echo .venv created successfully.
            echo.
        )
    )
)

if not defined PY_EXE (
    where py >nul 2>&1
    if not errorlevel 1 set "PY_EXE=py"
)

if not defined PY_EXE (
    where python >nul 2>&1
    if not errorlevel 1 set "PY_EXE=python"
)

echo ========================================
echo   NordFox Module Manager
echo ========================================
echo.

if not defined PY_EXE (
    echo [ERROR] Python not found.
    echo Install Python 3.11+ from python.org
    echo or create local virtual environment: py -m venv .venv
    pause
    exit /b 1
)

if not exist "requirements.txt" (
    echo [ERROR] requirements.txt not found in project root.
    pause
    exit /b 1
)

echo Checking dependencies...
%PY_EXE% -c "import PyQt6, win32com.client, PIL, segno, requests, openpyxl, flask, pypdf" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing missing dependencies...
    %PY_EXE% -m pip --version >nul 2>&1
    if %errorlevel% neq 0 (
        %PY_EXE% -m ensurepip --upgrade >nul 2>&1
    )
    %PY_EXE% -m pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo.
        echo [ERROR] Failed to install dependencies.
        pause
        exit /b 1
    )
)

echo Starting application...
%PY_EXE% main.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Application error
    pause
)

