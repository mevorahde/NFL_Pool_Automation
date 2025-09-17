# 🏈 NFL Spread Data Scraper & Excel Automation

This script automates the retrieval of NFL game spread data, filters for upcoming matchups, and updates a structured Excel workbook with formatted entries. It includes robust error handling, Gmail-based alerting, and log archiving for long-term reliability.

---

## 📦 Features

- **Web Scraping**: Extracts NFL game data including teams, dates, times, and betting spreads.
- **Timezone Conversion**: Converts UTC timestamps to Pacific Time for accurate filtering.
- **Excel Integration**: Updates or creates sheets with conditional formatting and dynamic row assignment.
- **Error Alerts**: Sends Gmail notifications for critical failures with log file attachments.
- **Log Archiving**: Compresses logs weekly using gzip and optionally clears originals.
- **MatchKey Normalization**: Ensures consistent row mapping across updates.

---

## 🛠️ Setup

### 1. Python Requirements

Install dependencies via pip:

```bash
pip install requests pandas beautifulsoup4 openpyxl python-dotenv pytz pytest 
```

### 2. Environment Configuration
Create a .env file in the root directory with the following values:\
\
file_path="C:\\Path\\To\\Your\\Family Football Pool YYYY.xlsx"\
#Toggle dry run mode (True = simulate without writing to Excel or sending emails)\
DRY_RUN=False\
#Sender email credentials (use app password for Gmail)\
EMAIL_ADDRESS=your_email@gmail.com\
EMAIL_PASSWORD=your_app_password\
#Recipient email for error alerts\
TO_EMAIL_ADDRESS=recipient_email@example.com\
#SMTP configuration (default for Gmail)\
SMTP_SERVER=smtp.gmail.com\
SMTP_PORT=587

🧠 Notes:
- DRY_RUN=False enables full execution including Excel updates and email alerts.
- EMAIL_ADDRESS and EMAIL_PASSWORD are used for Gmail SMTP authentication.
- TO_EMAIL_ADDRESS is the recipient of error alerts.
- file_path should point to your active NFL pool workbook.

🚀 Usage
Run manually or via scheduler:
```bash
python pool.py
```
📧 Email Alerts
Triggered on:
- Data scraping failure
- Excel update failure
- Unhandled exceptions
Each alert includes:
- Error message
- Timestamp
- Attached log file (.log)
- Archived log (.gz)

🧪 Testing & Validation
Pytest suite covers:
- Week Extraction: Validates get_week_number() across edge-case HTML files
- Row Assignment: Confirms weekday-based Excel row logic for Thursday, Friday, and Saturday games
- Abbreviation Mapping: Ensures all team names resolve to valid abbreviations
- Datetime Parsing: Verifies extract_datetime() returns proper datetime objects
- Excel Row Matching: Compares assigned rows against expected values in test_schedule.xlsx
- Mock HTML Structure: Validates that all test HTML files are compatible with the parser
Run tests with:

```bash
pytest tests/

```

📁 File Structure
NFL_Pool_Automation\
├── Family Football Pool YYYY.xlsx\
├──tests/\
│   ├── test_edge_cases.py\
│   ├── mock_html/\
│   │&emsp;&emsp;└── black_friday.html\
│   │&emsp;&emsp;└── christmas_tuesday.html\
│   │&emsp;&emsp;└── christmas_wednesday.html\
│   │&emsp;&emsp;└── friday_game.html\
│   │&emsp;&emsp;└── saturday_tripleheader.html\
│   │&emsp;&emsp;└── thanksgiving.html\
├── logs/\
│&emsp;&emsp;└── WEEKARCHIVELOGS_YYYY_MM_DD.log\
├── .env\
├── pool.py\
├── README.md\



🧠 Best Practices
- Layer enhancements modularly to avoid regressions.
- Always restore from backup before testing new logic.
- Normalize merge keys and log diagnostics before/after merge operations.
- Archive logs weekly and purge older ones to keep the system lean.
- Validate Excel updates with test data before deploying.

📌 Future Enhancements
- More test cases as the script is run over the course of one or many seasons.

🏈 Author
David — Software Engineer, automation enthusiast, and 49ers loyalist.
Focused on building reliable, unattended systems that just work.
