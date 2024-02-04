@echo off
:: set PSScript=%~dpn0.ps1
set PSScript=%~dp0dec2bin.ps1

:: powershell.exe -NoExit -File "%PSScript%" %*
powershell.exe -File "%PSScript%" %*