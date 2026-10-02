@echo off
title Stop FIRS Dashboard

echo.
echo  Stopping FIRS e-Invoice Dashboard...

set FOUND=0
for /f "tokens=2" %%p in ('wmic process where "name='python.exe' and commandline like '%%dashboard.py%%'" get processid ^| findstr /r "[0-9]"') do (
    echo  Killing PID %%p...
    taskkill /PID %%p /F > nul 2>&1
    set FOUND=1
)

if "%FOUND%"=="1" (
    echo  [OK]  Dashboard stopped.
) else (
    echo  [INFO] Dashboard was not running.
)
echo.
