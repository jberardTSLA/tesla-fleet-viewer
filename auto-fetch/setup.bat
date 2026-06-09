@echo off
title Tesla Fleet Viewer — Setup Auto-Fetch
echo.
echo   ==========================================
echo    TESLA FLEET VIEWER — Setup Auto-Fetch
echo   ==========================================
echo.

:: Check Node.js
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [!] Node.js non trovato. Scaricalo da https://nodejs.org/
    echo     Installa Node.js e rilancia questo script.
    echo.
    start https://nodejs.org/
    pause
    exit /b 1
)

:: Create folder
set "INSTALL_DIR=%USERPROFILE%\TeslaFleetAutoFetch"
echo [1/5] Creo cartella %INSTALL_DIR%...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

:: Download files from GitHub
echo [2/5] Scarico script da GitHub...
powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/jberardTSLA/tesla-fleet-viewer/main/auto-fetch/fetch.js' -OutFile '%INSTALL_DIR%\fetch.js'"
powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/jberardTSLA/tesla-fleet-viewer/main/auto-fetch/package.json' -OutFile '%INSTALL_DIR%\package.json'"
powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/jberardTSLA/tesla-fleet-viewer/main/auto-fetch/start.bat' -OutFile '%INSTALL_DIR%\start.bat'"

:: Install dependencies
echo [3/5] Installazione dipendenze npm...
cd /d "%INSTALL_DIR%"
call npm install

:: Create desktop shortcut
echo [4/5] Creo shortcut sul desktop...
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%USERPROFILE%\Desktop\Tesla Fleet Auto-Fetch.lnk'); $s.TargetPath = '%INSTALL_DIR%\start.bat'; $s.WorkingDirectory = '%INSTALL_DIR%'; $s.IconLocation = 'C:\Windows\System32\shell32.dll,172'; $s.Description = 'Avvia auto-fetch ZipLabs per Fleet Viewer'; $s.Save()"

echo.
echo   ==========================================
echo    SETUP COMPLETATO!
echo   ==========================================
echo.
echo   Icona "Tesla Fleet Auto-Fetch" creata sul desktop.
echo   Doppio click per avviare.
echo.
echo [5/5] Avvio primo login SSO...
echo   Fai login su ZipLabs nel browser che si apre,
echo   poi CHIUDI il browser.
echo.
pause
node fetch.js --login-only
