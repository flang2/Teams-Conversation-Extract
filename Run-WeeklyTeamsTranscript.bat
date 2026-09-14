@echo off
setlocal
title Weekly Teams Transcript Export


set "RUNNER=%~dp0tools\Invoke-WeeklyTeamsTranscript.ps1"


if not exist "%RUNNER%" (
    echo ERROR: The transcript runner was not found:
    echo %RUNNER%
    echo.
    echo Keep this BAT file in the repository root with the tools folder.
    pause
    exit /b 1
)


where pwsh.exe >nul 2>&1
if not errorlevel 1 (
    set "POWERSHELL_EXE=pwsh.exe"
) else (
    set "POWERSHELL_EXE=powershell.exe"
)


echo Starting the Weekly Teams Transcript export...
echo.
"%POWERSHELL_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%RUNNER%"
set "EXPORT_RESULT=%ERRORLEVEL%"


echo.
if "%EXPORT_RESULT%"=="0" (
    echo SUCCESS: The Weekly Teams Transcript was created.
) else (
    echo FAILED: The Weekly Teams Transcript was not created.
    echo Review the error above and Teams-Chat-Transcript-LastRun.log.
)
echo.
pause
exit /b %EXPORT_RESULT%