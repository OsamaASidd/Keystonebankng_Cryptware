@echo off
echo Installing Keystone Bank FIRS certificate into Trusted Root store...
certutil -addstore -f "Root" "e:\EINVOICING_AGENT\ssl\server.crt"
if %ERRORLEVEL% == 0 (
    echo.
    echo SUCCESS - Certificate installed. Restart Chrome/Edge then visit:
    echo   https://10.40.24.41:5443
) else (
    echo.
    echo FAILED - Make sure you right-clicked and chose "Run as administrator"
)
pause
