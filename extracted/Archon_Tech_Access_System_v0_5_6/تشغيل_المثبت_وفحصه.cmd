@echo off
setlocal
cd /d "%~dp0"
echo ================================================
echo Archon Tech Access System - Check and Install
echo ================================================
echo.
echo [1/2] Checking Windows PowerShell scripts...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tests\Verify-Windows-PowerShell.ps1"
if errorlevel 1 goto :failed
echo.
echo [2/2] Starting the Archon Tech installer...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\Install_Archon_Tech.ps1"
if errorlevel 1 goto :failed
echo.
echo Installer process finished. Review the messages above.
goto :done
:failed
echo.
echo The check or installer failed. Do not rerun repeatedly; keep this window open and share the error text.
:done
echo.
pause
endlocal
