@echo off
rem echo. %ERRORLEVEL%
powershell.exe -NoProfile -ExecutionPolicy unrestricted -command ".\Build-Catfish.ps1 %1 %2;exit $LASTEXITCODE"
rem echo.%ERRORLEVEL%

if %ERRORLEVEL% == 0 goto OK
echo ##teamcity[buildStatus status='FAILURE' text='{build.status.text} in compilation']
exit /B %ERRORLEVEL%

:OK