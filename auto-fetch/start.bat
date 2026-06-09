@echo off
title Tesla Fleet Viewer — Auto Fetch
echo.
echo   ========================================
echo   Tesla Fleet Viewer — Auto Fetch Setup
echo   ========================================
echo.

cd /d "%~dp0"

:: Check Node.js
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERRORE] Node.js non trovato!
    echo Scaricalo da: https://nodejs.org/
    echo.
    pause
    exit /b 1
)

:: Install dependencies if needed
if not exist "node_modules" (
    echo [SETUP] Installazione dipendenze...
    call npm install
    echo.
)

:: Check if first run (no chrome profile)
if not exist "chrome-profile" (
    echo [LOGIN] Primo avvio — apro il browser per login SSO.
    echo Fai login su ZipLabs, poi CHIUDI il browser.
    echo.
    node fetch.js --login-only
) else (
    echo [START] Avvio fetch automatico ogni 5 minuti...
    echo.
    node fetch.js
)

pause
