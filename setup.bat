@echo off
REM Excel Telegram Bot Setup Script for Windows

echo.
echo ======================================
echo  Excel Telegram Bot - Setup
echo ======================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH!
    echo Please install Python 3.9 or higher from https://www.python.org/
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

echo [OK] Python is installed
python --version

echo.
echo Installing dependencies from requirements.txt...
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies!
    pause
    exit /b 1
)

echo.
echo [OK] Dependencies installed successfully!
echo.
echo ======================================
echo  Setup Complete!
echo ======================================
echo.
echo Next steps:
echo 1. Get your Telegram bot token:
echo    - Open Telegram and find @BotFather
echo    - Send /newbot and follow the prompts
echo    - Copy the token provided
echo.
echo 2. Open bot.py and replace:
echo    TELEGRAM_BOT_TOKEN = "YOUR_BOT_TOKEN_HERE"
echo    with your actual token
echo.
echo 3. Run the bot:
echo    python bot.py
echo.
pause
