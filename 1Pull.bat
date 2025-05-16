@echo off
echo Pulling Turbo from GAS...
clasp pull

cd Turbo
if %ERRORLEVEL% neq 0 (
  echo ERROR: Failed to change directory to Turbo.
  pause
  exit /b %ERRORLEVEL%
)

echo Deleting files starting with _ in %CD%...
for %%F in (_*) do (
    echo Deleting "%%F"
    del "%%F"
)

echo Deleting files starting with = in %CD%...
for %%F in (=*) do (
    echo Deleting "%%F"
    del "%%F"
)

cd ..
<<<<<<< HEAD

pause
=======
>>>>>>> tryAuthentication
