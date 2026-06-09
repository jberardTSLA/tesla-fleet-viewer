@echo off
chcp 65001 >nul
title Tesla Fleet Viewer - Setup Auto-Fetch
echo.
echo   ==========================================
echo    TESLA FLEET VIEWER - Setup Auto-Fetch
echo   ==========================================
echo.

:: Check Node.js in PATH
where node >nul 2>nul
if %errorlevel% equ 0 goto :node_found

:: Check common Node.js paths
if exist "C:\Program Files\nodejs\node.exe" (
    set "PATH=C:\Program Files\nodejs;%PATH%"
    goto :node_found
)
if exist "C:\Program Files (x86)\nodejs\node.exe" (
    set "PATH=C:\Program Files (x86)\nodejs;%PATH%"
    goto :node_found
)
if exist "%APPDATA%\..\Local\Programs\nodejs\node.exe" (
    set "PATH=%APPDATA%\..\Local\Programs\nodejs;%PATH%"
    goto :node_found
)
if exist "%USERPROFILE%\AppData\Local\Programs\Git\cmd\node.exe" (
    set "PATH=%USERPROFILE%\AppData\Local\Programs\Git\cmd;%PATH%"
    goto :node_found
)

:: Node not found - offer to install
echo [!] Node.js non trovato sul tuo PC.
echo.
echo     Vuoi installarlo automaticamente? (richiede connessione internet)
echo.
choice /c SN /m "Installare Node.js? (S/N)"
if %errorlevel% equ 2 (
    echo.
    echo Scaricalo manualmente da: https://nodejs.org/
    start https://nodejs.org/
    pause
    exit /b 1
)

:: Download and install Node.js
echo.
echo [INSTALL] Scarico Node.js...
powershell -Command "Invoke-WebRequest -Uri 'https://nodejs.org/dist/v20.11.1/node-v20.11.1-x64.msi' -OutFile '%TEMP%\node_install.msi'"
echo [INSTALL] Installazione in corso (potrebbe chiedere permessi admin)...
msiexec /i "%TEMP%\node_install.msi" /qn
set "PATH=C:\Program Files\nodejs;%PATH%"
timeout /t 5 >nul

:: Verify
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [!] Installazione fallita. Scarica manualmente da https://nodejs.org/
    start https://nodejs.org/
    pause
    exit /b 1
)

:node_found
echo [OK] Node.js trovato: 
node --version
echo.

:: Create folder
set "INSTALL_DIR=%USERPROFILE%\TeslaFleetAutoFetch"
echo [1/5] Creo cartella %INSTALL_DIR%...
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

:: Download files from GitHub
echo [2/5] Scarico script da GitHub...
powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/jberardTSLA/tesla-fleet-viewer/main/auto-fetch/fetch.js' -OutFile '%INSTALL_DIR%\fetch.js'" 2>nul
powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/jberardTSLA/tesla-fleet-viewer/main/auto-fetch/package.json' -OutFile '%INSTALL_DIR%\package.json'" 2>nul
powershell -Command "Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/jberardTSLA/tesla-fleet-viewer/main/auto-fetch/start.bat' -OutFile '%INSTALL_DIR%\start.bat'" 2>nul

:: Install dependencies
echo [3/5] Installazione Puppeteer (puo richiedere qualche minuto)...
cd /d "%INSTALL_DIR%"
call npm install 2>nul

:: Create desktop shortcut
echo [4/5] Creo shortcut sul desktop...
powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%USERPROFILE%\Desktop\Tesla Fleet Auto-Fetch.lnk'); $s.TargetPath = '%INSTALL_DIR%\start.bat'; $s.WorkingDirectory = '%INSTALL_DIR%'; $s.IconLocation = 'C:\Windows\System32\shell32.dll,172'; $s.Description = 'Avvia auto-fetch ZipLabs per Fleet Viewer'; $s.Save()"

echo.
echo   ==========================================
echo    SETUP COMPLETATO!
echo   ==========================================
echo.
echo   Icona "Tesla Fleet Auto-Fetch" sul desktop.
echo   Doppio click per avviare il fetch automatico.
echo.
echo [5/5] Avvio primo login SSO...
echo   Fai login su ZipLabs nel browser che si apre,
echo   poi CHIUDI il browser quando sei dentro.
echo.
pause
node fetch.js --login-only
