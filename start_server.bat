@echo off
title RDC Incentive Calculator - Auto Restart
:restart
echo.
echo ============================================================
echo   RDC Batching Incentive Calculator
echo   Starting server on http://localhost:2001
echo ============================================================
cd /d "d:\Dipjyoti Doimari\AI\Incentive_Calculator"
python app.py
echo.
echo Server stopped. Restarting in 5 seconds... (Press Ctrl+C to quit)
timeout /t 5 /nobreak
goto restart
