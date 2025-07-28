@echo off
echo ===================================
echo Pushing to MSGLib...
echo ===================================

echo.
echo Changing directory to MSGLib...
cd /d "H:\My Drive\_RPG\GitHub\TurboApp\MSGLib"

REM Check if the .clasp.json file exists before proceeding.
if not exist ".clasp.json" (
    echo ERROR: .clasp.json not found in the target directory!
    echo PATH: H:\My Drive\_RPG\GitHub\TurboApp\MSGLib
    pause
    exit /b
)

echo Executing 'clasp push' for project: MSGLib
clasp push

echo.
echo --- Push to MSGLib Complete ---