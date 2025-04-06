copy .\Turbo\*.* .\Text\Turbo\*.txt
copy .\MSGLib\*.* .\Text\MSGLib\*.txt
tar -c -a -f "Text\Turbo\TurboFiles.zip" "Turbo"
tar -c -a -f "Text\MSGLib\MSGLibFiles.zip" "MSGLib"