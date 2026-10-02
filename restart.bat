@echo off
title Restart FIRS Dashboard

echo.
echo  Keystone Bank FIRS e-Invoice Dashboard — Restart
echo  ==================================================

REM Stop any running instance
echo  Stopping existing instance...
set STOPPED=0
for /f "tokens=2" %%p in ('wmic process where "name='python.exe' and commandline like '%%dashboard.py%%'" get processid ^| findstr /r "[0-9]"') do (
    taskkill /PID %%p /F > nul 2>&1
    echo  Killed PID %%p
    set STOPPED=1
)
if "%STOPPED%"=="0" echo  (No running instance found)

timeout /t 2 /nobreak > nul

REM Rotate logs (keep last 1000 lines each)
echo  Rotating logs...
for %%f in (dashboard.log dashboard_out.log dashboard_err.log invoice_scheduler.log scheduler_error.log) do (
    if exist "E:\EINVOICING_AGENT\logs\%%f" (
        powershell -Command "Get-Content 'E:\EINVOICING_AGENT\logs\%%f' -Tail 1000 | Set-Content 'E:\EINVOICING_AGENT\logs\%%f' -Encoding utf8" 2>nul
    )
)

REM Start fresh
echo  Starting dashboard...
start "" /B "E:\EINVOICING_AGENT\venv\Scripts\python.exe" "E:\EINVOICING_AGENT\dashboard.py" ^
    >> "E:\EINVOICING_AGENT\logs\dashboard_out.log" 2>> "E:\EINVOICING_AGENT\logs\dashboard_err.log"

timeout /t 9 /nobreak > nul

REM Verify
netstat -ano | findstr ":5443" | findstr "LISTENING" > nul
if %errorlevel%==0 (
    echo  [OK]  Dashboard is running: https://10.40.24.41:5443
) else (
    echo  [FAIL] Port 5443 not listening. Check dashboard_err.log for details.
)
echo.
