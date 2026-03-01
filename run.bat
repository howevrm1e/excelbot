@echo off
REM Run Excel Telegram Bot

echo.
echo ======================================
echo  Excel Telegram Bot - Starting
echo ======================================
echo.

python bot.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Bot failed to start!
    echo.
    echo Possible issues:
    echo - Bot token not configured in bot.py
    echo - Dependencies not installed (run setup.bat first)
    echo - Internet connection issue
    echo.
)

pause
