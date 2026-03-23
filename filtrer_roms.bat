@echo off
:: ============================================================
:: filtrer_roms.bat
:: Glisser-déposer un fichier .xlsx sur ce .bat pour lancer
:: filtrer_roms.py automatiquement.
::
:: Prérequis : Python 3 installé et dans le PATH
:: filtrer_roms.py doit être dans le même dossier que ce .bat
:: ============================================================

:: Vérifier qu'un fichier a été déposé
if "%~1"=="" (
    echo.
    echo  Utilisation : Glisser-deposer un fichier .xlsx sur ce script.
    echo.
    pause
    exit /b
)

:: Vérifier que c'est bien un .xlsx
if /i not "%~x1"==".xlsx" (
    echo.
    echo  Erreur : Le fichier "%~1" n'est pas un .xlsx
    echo.
    pause
    exit /b
)

:: Dossier du .bat = dossier du script Python
set SCRIPT_DIR=%~dp0
set PY_SCRIPT=%SCRIPT_DIR%filtrer_roms.py

:: Vérifier que filtrer_roms.py est présent
if not exist "%PY_SCRIPT%" (
    echo.
    echo  Erreur : filtrer_roms.py introuvable dans %SCRIPT_DIR%
    echo.
    pause
    exit /b
)

:: Dossier de sortie = même dossier que le xlsx déposé, sous-dossier "download_lists"
set OUTPUT_DIR=%~dp1download_lists

echo.
echo  Fichier  : %~1
echo  Sortie   : %OUTPUT_DIR%
echo.

python "%PY_SCRIPT%" "%~1" "%OUTPUT_DIR%"

echo.
pause
