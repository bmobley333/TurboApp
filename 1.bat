@echo off
:Menu
cls
echo ===================================
echo  Google Apps Script Project Manager
echo ===================================
echo.
echo  1. PUSH to MSGLib
echo  2. PULL from MSGLib
echo.
echo  3. PULL from Turbo
echo  4. PUSH to Turbo
echo.
echo  5. Exit
echo.
echo ===================================
echo.

CHOICE /C 12345 /M "Enter your choice: "

if errorlevel 5 goto :Exit
if errorlevel 4 goto :PushTurbo
if errorlevel 3 goto :PullTurbo
if errorlevel 2 goto :PullMSGLib
if errorlevel 1 goto :PushMSGLib

REM --- MSGLib PULL Section ---
:PullMSGLib
echo.
echo ---------------------------------------------------
echo Changing directory to MSGLib...
cd /d "H:\My Drive\_RPG\GitHub\TurboApp\MSGLib"

if not exist ".clasp.json" (
    echo ERROR: .clasp.json not found in the target directory!
    echo PATH: H:\My Drive\_RPG\GitHub\TurboApp\MSGLib
    pause
    goto :Menu
)

echo Executing 'clasp PULL' for project: MSGLib
clasp pull
echo ---------------------------------------------------
echo.
pause
goto :Menu

REM --- MSGLib PUSH Section ---
:PushMSGLib
echo.
echo ---------------------------------------------------
echo Changing directory to MSGLib...
cd /d "H:\My Drive\_RPG\GitHub\TurboApp\MSGLib"

if not exist ".clasp.json" (
    echo ERROR: .clasp.json not found in the target directory!
    echo PATH: H:\My Drive\_RPG\GitHub\TurboApp\MSGLib
    pause
    goto :Menu
)

echo WARNING: You are about to PUSH to the MSGLib project.
echo This will OVERWRITE remote files.
CHOICE /C YN /M "Are you sure? (Y/N)"
if errorlevel 2 (
    echo Push cancelled.
    pause
    goto :Menu
)

echo Executing 'clasp PUSH' for project: MSGLib
clasp push
echo ---------------------------------------------------
echo.
pause
goto :Menu

REM --- Turbo PULL Section ---
:PullTurbo
echo.
echo ---------------------------------------------------
echo Changing directory to Turbo...
cd /d "H:\My Drive\_RPG\GitHub\TurboApp\Turbo"

if not exist ".clasp.json" (
    echo ERROR: .clasp.json not found in the target directory!
    echo PATH: H:\My Drive\_RPG\GitHub\TurboApp\Turbo
    pause
    goto :Menu
)

echo Executing 'clasp PULL' for project: Turbo
clasp pull
echo ---------------------------------------------------
echo.
pause
goto :Menu

REM --- Turbo PUSH Section ---
:PushTurbo
echo.
echo ---------------------------------------------------
echo Changing directory to Turbo...
cd /d "H:\My Drive\_RPG\GitHub\TurboApp\Turbo"

if not exist ".clasp.json" (
    echo ERROR: .clasp.json not found in the target directory!
    echo PATH: H:\My Drive\_RPG\GitHub\TurboApp\Turbo
    pause
    goto :Menu
)

echo WARNING: You are about to PUSH to the Turbo project.
echo This will OVERWRITE remote files.
CHOICE /C YN /M "Are you sure? (Y/N)"
if errorlevel 2 (
    echo Push cancelled.
    pause
    goto :Menu
)

echo Executing 'clasp PUSH' for project: Turbo
clasp push
echo ---------------------------------------------------
echo.
pause
goto :Menu

:Exit
exit /b