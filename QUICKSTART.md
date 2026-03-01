# 🚀 Quick Start Guide

## 5-Minute Setup

### Step 1: Install Dependencies (Windows)
```bash
setup.bat
```

Or manually:
```bash
pip install -r requirements.txt
```

### Step 2: Get Your Bot Token
1. Open **Telegram** app
2. Search for `@BotFather` and open chat
3. Send `/newbot`
4. Follow the steps (choose a name and username)
5. Copy the token (long string of characters)

### Step 3: Configure Token
1. Open `bot.py` in your text editor
2. Find line 11: `TELEGRAM_BOT_TOKEN = "YOUR_BOT_TOKEN_HERE"`
3. Replace `YOUR_BOT_TOKEN_HERE` with your copied token
4. Save the file

### Step 4: Run the Bot
```bash
python bot.py
```

Or on Windows, double-click:
```
run.bat
```

🎉 **Bot is now running!**

## Start Using Your Bot

1. In Telegram, find your bot (the username you created in Step 2)
2. Send `/start` to see the welcome message
3. Send any text or data - the bot will create an Excel file!

## Example 1: Simple Text
Send:
```
John
Jane
Bob
```

Result: Excel file with one column, three rows

## Example 2: CSV Format
Send:
```
Name,Age,City
John,25,New York
Jane,30,London
```

Result: Excel file with 3 columns, 2 data rows + header

## Commands
- `/start` - Welcome & info
- `/help` - Detailed help
- `/example` - See CSV example

## Need Help?
See `README.md` for troubleshooting and advanced configuration
