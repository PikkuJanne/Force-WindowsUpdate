@echo off
REM Simple launcher so users can just one-click the .cmd file.
REM -ExecutionPolicy Bypass: ignore local policy for this run
REM -NoExit: keep the PowerShell window open after the script ends
REM -NoLogo / -NoProfile: cleaner output

powershell.exe ^
  -ExecutionPolicy Bypass ^
  -NoLogo ^
  -NoProfile ^
  -NoExit ^
  -File "%~dp0Force-WindowsUpdate.ps1"