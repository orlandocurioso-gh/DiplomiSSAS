@echo off
SETLOCAL

REM ----------------------------------------------------
REM 1. AVVIA IL SERVER FLASK (SENZA ICONA SULLA BARRA)
REM ----------------------------------------------------

SET FLASK_APP_PATH="C:\Users\484972\Documents\Coding\Python\Sapienza\Diplomi\app_altaformazione.py"

REM Sostituito 'start /min cmd /c' con PowerShell per NASCONDERE l'icona.
powershell -WindowStyle Hidden -Command "Start-Process python -ArgumentList '-u', %FLASK_APP_PATH% -Wait:$false"

REM ----------------------------------------------------
REM 2. ATTENDI L'AVVIO DEL SERVER
REM ----------------------------------------------------

timeout /t 5 /nobreak >nul

REM ----------------------------------------------------
REM 3. APRI LA PAGINA NEL BROWSER A TUTTO SCHERMO
REM ----------------------------------------------------

SET CHROME_PATH="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
SET TARGET_URL="http://127.0.0.1:5000/"

REM Usiamo l'opzione più robusta per l'apertura a schermo intero senza schede.
start "" %CHROME_PATH% --app=%TARGET_URL% --kiosk

exit