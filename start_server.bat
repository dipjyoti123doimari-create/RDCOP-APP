@echo off
title RDC-OPS Server — DO NOT CLOSE THIS WINDOW
color 0A

echo ============================================================
echo   RDC-OPS Server
echo   Running on http://localhost:2001
echo   DO NOT CLOSE THIS WINDOW — closing it stops the server
echo ============================================================
echo.

cd /d "d:\Dipjyoti Doimari\AI\Incentive_Calculator"
set PYTHONIOENCODING=utf-8

:restart
echo [%date% %time%] Starting server...
"C:\Users\Dipjyoti Doimari\AppData\Local\Programs\Python\Python312\python.exe" app.py
echo.
echo [%date% %time%] Server stopped. Restarting in 3 seconds...
timeout /t 3 /nobreak >nul
goto restart
