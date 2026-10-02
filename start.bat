@echo off
title Keystone Bank FIRS e-Invoice Dashboard

echo.
echo  Keystone Bank FIRS e-Invoice Dashboard
echo  ========================================

REM Check if already running
for /f "tokens=2" %%p in ('wmic process where "name='python.exe' and commandline like '%%dashboard.py%%'" get processid ^| findstr /r "[0-9]"') do (
    echo  [WARN] Dashboard already running on PID %%p. Run restart.bat instead.
    echo.
    pause
    exit /b 1
)

echo  Starting dashboard...
start "" /B "E:\EINVOICING_AGENT\venv\Scripts\python.exe" "E:\EINVOICING_AGENT\dashboard.py" ^
    >> "E:\EINVOICING_AGENT\logs\dashboard_out.log" 2>> "E:\EINVOICING_AGENT\logs\dashboard_err.log"

timeout /t 8 /nobreak > nul

REM Verify it's listening
netstat -ano | findstr ":5443" | findstr "LISTENING" > nul
if %errorlevel%==0 (
    echo  [OK]  Dashboard is running: https://10.40.24.41:5443
) else (
    echo  [FAIL] Port 5443 not listening. Check dashboard_err.log for errors.
)
echo.
