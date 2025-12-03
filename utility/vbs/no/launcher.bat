REM ----------------------------------------------------
REM 1. AVVIA IL SERVER FLASK (TOTALMENTE INVISIBILE)
REM ----------------------------------------------------

SET FLASK_APP_PATH="C:\Users\484972\Documents\Coding\Python\Sapienza\Diplomi\app_altaformazione.py"
SET TEMP_VBS_FILE="%TEMP%\startflask.vbs"

REM 1. Crea il contenuto del VBScript in un file temporaneo Txt
echo Set WshShell = CreateObject("WScript.Shell")> "%TEMP%\flask_content.txt"
echo WshShell.Run "python ""%FLASK_APP_PATH%""", 0, false>> "%TEMP%\flask_content.txt"

REM 2. Rinomina il file Txt in VBS (più robusto contro i problemi di reindirizzamento)
copy /y "%TEMP%\flask_content.txt" %TEMP_VBS_FILE% >nul

REM 3. Rimuovi l'attributo di sola lettura nel caso fosse rimasto
attrib -r %TEMP_VBS_FILE% 2>nul

REM 4. Esegui il VBScript
start /b wscript %TEMP_VBS_FILE%

REM ----------------------------------------------------
REM 1. AVVIA IL SERVER FLASK (TOTALMENTE INVISIBILE)
REM ----------------------------------------------------

SET FLASK_APP_PATH="C:\Users\484972\Documents\Coding\Python\Sapienza\Diplomi\app_altaformazione.py"
SET TEMP_VBS_FILE="%TEMP%\startflask.vbs"

REM 1. Crea il contenuto del VBScript in un file temporaneo Txt
echo Set WshShell = CreateObject("WScript.Shell")> "%TEMP%\flask_content.txt"
echo WshShell.Run "python ""%FLASK_APP_PATH%""", 0, false>> "%TEMP%\flask_content.txt"

REM 2. Rinomina il file Txt in VBS (più robusto contro i problemi di reindirizzamento)
copy /y "%TEMP%\flask_content.txt" %TEMP_VBS_FILE% >nul

REM 3. Rimuovi l'attributo di sola lettura nel caso fosse rimasto
attrib -r %TEMP_VBS_FILE% 2>nul

REM 4. Esegui il VBScript
start /b wscript %TEMP_VBS_FILE%

REM ----------------------------------------------------