@echo off
title RDC Incentive Calculator - Watchdog

set APP_DIR=d:\Dipjyoti Doimari\AI\Incentive_Calculator

echo.
echo ============================================================
echo   RDC Batching Incentive Calculator
echo   Starting watchdog on http://localhost:2001
echo   Crash log: %APP_DIR%\server_crash.log
echo   App log:   %APP_DIR%\server.log
echo ============================================================

cd /d "%APP_DIR%"
python watchdog.py

echo.
echo Watchdog exited. Press any key to close.
pause
