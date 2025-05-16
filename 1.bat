@echo off
REM Combined Batch Script:
REM 1. Pushes specific Turbo project files to Google Apps Script using clasp.
REM    (Assumes appsscript.json in .\Turbo controls which files are pushed via filePushOrder)
REM 2. If push is successful, performs cleanup and copies files to Text backup folders.
REM Uses CALL to invoke clasp commands to prevent premature script termination.

echo [DEBUG] Combined Script started.

REM --- PART 1: Clasp Push Operation ---

REM Set the project directory relative to this batch file's location.
SET "PROJECT_DIR=.\Turbo"
echo [DEBUG] PROJECT_DIR set to: %PROJECT_DIR%
if %errorlevel% neq 0 (
    echo [ERROR] Failed to SET PROJECT_DIR. Errorlevel: %errorlevel%
    goto :EarlyExitError
)

echo ============================================================
echo  PART 1: Pushing Turbo Project Files to Google Apps Script
echo ============================================================
echo.
echo Project Directory: %PROJECT_DIR%
echo Files to be pushed are controlled by 'filePushOrder' in
echo   %PROJECT_DIR%\appsscript.json
echo.

echo [DEBUG] About to check clasp version using CALL...
CALL clasp --version
SET CLASP_VERSION_EXIT_CODE=%errorlevel%
echo [DEBUG] Immediate errorlevel after 'CALL clasp --version' is: %CLASP_VERSION_EXIT_CODE%
echo [DEBUG] Now checking the IF condition for clasp version...

if %CLASP_VERSION_EXIT_CODE% neq 0 (
    echo [ERROR] "CALL clasp --version" command failed or clasp not found. Errorlevel: %CLASP_VERSION_EXIT_CODE%
    echo         Please install clasp globally: npm install -g @google/clasp
    goto :EarlyExitError
) else (
    echo [INFO] clasp installation found. (Version check successful, errorlevel: %CLASP_VERSION_EXIT_CODE%)
)
echo [DEBUG] Successfully passed the clasp version IF/ELSE block.
echo.

echo [DEBUG] About to change directory to project root for clasp...
REM Navigate to the project directory. The /D switch is used to change the current drive as well if necessary.
echo [STEP 1.1] Changing directory to %PROJECT_DIR%...
cd /D "%PROJECT_DIR%"
SET CD_EXIT_CODE=%errorlevel%
echo [DEBUG] Immediate errorlevel after 'cd' is: %CD_EXIT_CODE%
if %CD_EXIT_CODE% neq 0 (
    echo [ERROR] Failed to change directory to %PROJECT_DIR%. Errorlevel: %CD_EXIT_CODE%
    echo         Current directory before attempting cd: %~dp0
    echo         Target directory was: %PROJECT_DIR%
    echo         Ensure the path is correct relative to the batch file and the directory exists.
    goto :EarlyExitError
)
echo [INFO] Current directory successfully changed to: %cd%
echo.

REM Perform the push operation.
echo [STEP 1.2] Pushing files to Google Apps Script with force (-f) using CALL...
CALL clasp push -f
SET CLASP_PUSH_EXIT_CODE=%errorlevel%
echo [DEBUG] Immediate errorlevel after 'CALL clasp push -f' is: %CLASP_PUSH_EXIT_CODE%

if %CLASP_PUSH_EXIT_CODE% equ 0 (
    echo.
    echo [SUCCESS] Successfully pushed files to Google Apps Script.
    goto :ContinueAfterPushCheck
) else (
    echo.
    echo [ERROR] "CALL clasp push -f" failed. Errorlevel: %CLASP_PUSH_EXIT_CODE%. See output above for details.
    echo         Common issues:
    echo           - Not logged into clasp ^(run "clasp login"^).
    echo           - Incorrect Script ID in .clasp.json.
    echo           - Issues with appsscript.json or file permissions.
    echo.
    echo [INFO] Skipping Part 2 (Cleanup and Copy) due to push failure.
    goto :EarlyExitError
)

:ContinueAfterPushCheck
echo.

REM --- PART 2: Cleanup and File Copy Operations ---
REM This part only runs if clasp push was successful.

echo ============================================================
echo  PART 2: Performing Cleanup and File Copy Operations
echo ============================================================
echo.

REM At this point, we are still inside the %PROJECT_DIR% (.\Turbo)

echo [STEP 2.1] Current directory for initial cleanup: %cd%
echo [DEBUG] Listing directory contents (Turbo)...
dir
echo.

echo [STEP 2.2] Deleting files starting with '_' in %PROJECT_DIR%...
del /Q _*.*
if %errorlevel% neq 0 (
    echo [WARNING] No files found matching _*.* or error during deletion.
) else (
    echo [INFO] Deleted files starting with '_'.
)
echo.

echo [STEP 2.3] Deleting files starting with '=' in %PROJECT_DIR%...
del /Q =*.*
if %errorlevel% neq 0 (
    echo [WARNING] No files found matching =*.* or error during deletion.
) else (
    echo [INFO] Deleted files starting with '='.
)
echo.

echo [STEP 2.4] Changing directory back to the script's root folder (%~dp0)...
cd /D "%~dp0"
SET CD_BACK_EXIT_CODE=%errorlevel%
if %CD_BACK_EXIT_CODE% neq 0 (
    echo [ERROR] Failed to change directory back to script root. Errorlevel: %CD_BACK_EXIT_CODE%
    echo         Current directory: %cd%
    goto :EarlyExitError
)
echo [INFO] Current directory successfully changed to: %cd%
echo.

echo [STEP 2.5] Deleting contents of .\Text\Turbo\...
del /Q .\Text\Turbo\*.*
if %errorlevel% neq 0 (
    echo [WARNING] No files found in .\Text\Turbo\ or error during deletion.
) else (
    echo [INFO] Deleted contents of .\Text\Turbo\.
)
echo.

echo [STEP 2.6] Deleting contents of .\Text\MSGLib\...
del /Q .\Text\MSGLib\*.*
if %errorlevel% neq 0 (
    echo [WARNING] No files found in .\Text\MSGLib\ or error during deletion.
) else (
    echo [INFO] Deleted contents of .\Text\MSGLib\.
)
echo.

echo [STEP 2.7] Copying .\Turbo\*.* to .\Text\Turbo\ as *.txt...
copy .\Turbo\*.* .\Text\Turbo\*.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to copy files from .\Turbo\ to .\Text\Turbo\. Errorlevel: %errorlevel%
    goto :EarlyExitError
) else (
    echo [INFO] Copied files from .\Turbo\ to .\Text\Turbo\.
)
echo.

echo [STEP 2.8] Copying .\MSGLib\*.* to .\Text\MSGLib\ as *.txt...
copy .\MSGLib\*.* .\Text\MSGLib\*.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to copy files from .\MSGLib\ to .\Text\MSGLib\. Errorlevel: %errorlevel%
    goto :EarlyExitError
) else (
    echo [INFO] Copied files from .\MSGLib\ to .\Text\MSGLib\.
)
echo.

REM PowerShell commands for zipping are still commented out as per your original 2.bat
:: powershell Compress-Archive -Path "Turbo\*" -DestinationPath "Text\Turbo\TurboFiles.zip" -Force
:: powershell Compress-Archive -Path "MSGLib\*" -DestinationPath "Text\MSGLib\MSGLibFiles.zip" -Force

echo [SUCCESS] Cleanup and copy operations completed.
goto :NormalEnd

:EarlyExitError
echo [FATAL ERROR] Script aborted due to a critical error.
echo Please check the messages above.

:NormalEnd
echo ============================================================
echo  Combined Script finished.
echo ============================================================
<<<<<<< HEAD
pause
=======
pause
>>>>>>> tryAuthentication
