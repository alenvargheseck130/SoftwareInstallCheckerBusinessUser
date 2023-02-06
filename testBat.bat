@echo off

PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& {Start-Process PowerShell -ArgumentList '-NoProfile -ExecutionPolicy Bypass -File "%CD%\SoftwareInstallChecker.ps1"'-Verb RunAs} 
::"C:\Users\%USERNAME%\Desktop\SoftwareInstallCheckerBusinessUser\SoftwareInstallChecker.ps1"' -Verb RunAs}"
echo task ended
pause
exit