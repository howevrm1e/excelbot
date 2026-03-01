# Excel Telegram Bot 📊

A Telegram bot that converts text messages into Excel (.xlsx) files. Simply send text or data to the bot, and it will create and send back a formatted Excel file.

## Features

✅ Convert text to Excel files instantly
✅ Support for simple text format (one line = one row)
✅ Support for CSV (Comma-Separated Values) format
✅ Auto-adjusting column widths
✅ Clean and user-friendly interface
✅ Command-based help system

## Prerequisites

- Python 3.9 or higher
- A Telegram account
- pip (Python package manager)

## Installation & Setup

### Step 1: Install Python Dependencies

```bash
pip install -r requirements.txt
```

### Step 2: Create Your Telegram Bot

1. Open Telegram and search for **@BotFather**
2. Send `/newbot` command
3. Follow the prompts:
   - Enter your bot's name (e.g., "Excel Bot")
   - Enter a username for your bot (must be unique, ending with "bot")
4. BotFather will send you an API token (a long string of characters)

### Step 3: Configure the Bot Token

1. Open `bot.py` file
2. Find this line:
   ```python
   TELEGRAM_BOT_TOKEN = "YOUR_BOT_TOKEN_HERE"
   ```
3. Replace `YOUR_BOT_TOKEN_HERE` with the token you received from BotFather
4. Save the file

**⚠️ IMPORTANT**: Never share your bot token publicly or commit it to version control!

### Step 4: Run the Bot

```bash
python bot.py
```

You should see:
```
🤖 Excel Bot is running...
Press Ctrl+C to stop
```

## How to Use

### Available Commands

- `/start` - Show welcome message
- `/help` - Show help information
- `/example` - Show an example of CSV format

### Sending Data

#### Simple Text Format
Just send any text, each line becomes a new row:
```
John
Jane
Bob
```

#### CSV Format (Structured Data)
Send comma-separated values for better organization:
```
Name,Age,City
John,25,New York
Jane,30,London
Bob,28,Paris
```

#### Multi-line Text
The bot automatically handles:
- Multiple paragraphs
- Lists
- Any text format

### Example Usage

1. Start a chat with your bot on Telegram
2. Send `/start`
3. Send your data (text or CSV)
4. The bot creates an Excel file and sends it back

## File Structure

```
Excel bot/
├── bot.py           # Main bot code
├── requirements.txt # Python dependencies
└── README.md        # This file
```

## Troubleshooting

### Bot Token Not Working
- Double-check that you copied the entire token from BotFather
- Make sure there are no extra spaces before or after the token
- If you lost the token, create a new bot with BotFather

### Bot Not Starting
- Ensure Python 3.9+ is installed: `python --version`
- Verify all dependencies are installed: `pip install -r requirements.txt`
- Check internet connection (bot needs it to connect to Telegram)

### Excel File Not Generated
- Make sure your data doesn't have special formatting issues
- Try the example format: `Name,Age,City` with proper comma separation
- Check that you have write permissions in the directory

## Advanced Configuration

You can modify the bot behavior by editing `bot.py`:

- **Auto-adjust column width**: Currently set to max 50 characters. Change the `adjusted_width = min(max_length + 2, 50)` line
- **File naming**: Change the filename format in the `filename` variable
- **Add new commands**: Follow the pattern of existing command handlers

## Future Enhancements

Possible features to add:
- Multiple sheet support
- Custom column headers
- Data validation and formatting
- Support for images/attachments
- Database integration
- Scheduling tasks

## Support

If you encounter any issues:
1. Check the troubleshooting section above
2. Review error messages in the console
3. Ensure all dependencies are correctly installed
4. Make sure your Telegram token is valid

## License

MIT License - Feel free to modify and use as needed.

---

**Happy Excel creating! 📊**
