cd Turbo
dir
del _*.*
del =*.*
cd ..

del /Q .\Text\Turbo\*.*
del /Q .\Text\MSGLib\*.*
copy .\Turbo\*.* .\Text\Turbo\*.txt
copy .\MSGLib\*.* .\Text\MSGLib\*.txt
:: powershell Compress-Archive -Path "Turbo\*" -DestinationPath "Text\Turbo\TurboFiles.zip" -Force
:: powershell Compress-Archive -Path "MSGLib\*" -DestinationPath "Text\MSGLib\MSGLibFiles.zip" -Force
