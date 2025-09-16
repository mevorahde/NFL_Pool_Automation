# 🏈 NFL Spread Data Scraper & Excel Automation

Automates the retrieval of NFL game spread data, filters for upcoming matchups, and updates a structured Excel workbook. Includes robust error handling, Gmail-based alerting, and log archiving for long-term reliability.

---

## 📦 Features

- Web scraping of NFL game data (teams, dates, times, spreads)
- Timezone conversion (UTC → Pacific)
- Excel integration with conditional formatting
- Gmail alerts for critical failures (with log attachments)
- Weekly log archiving via gzip
- MatchKey normalization for consistent row mapping

---

## 🛠️ Setup

### 1. Install Dependencies

pip install -r requirements.txt

### 2. Configure Environment Variables

Create a `.env` file in the root directory:

file_path="C:\\Users\\demev\\Version-Control\\NFL_Pool_Automation\\Family Football Pool 2025.xlsx"
DRY_RUN=False
EMAIL_ADDRESS=mevorahd@gmail.com
TO_EMAIL_ADDRESS=mevorahde@hotmail.com
EMAIL_PASSWORD=ocwujqeubmljgesl
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587

🔒 Tip: Use `.env.template` for safe sharing and add `.env` to `.gitignore`.

---

## 🚀 Usage

Run manually or via scheduler:

python pool.py

---

## 📧 Email Alerts

Alerts are triggered on:
- Scraping failures
- Excel update errors
- Unhandled exceptions

Each email includes:
- Error message
- Timestamp
- Attached `.log` file
- Archived `.gz` version

---

## 🧪 Testing & Validation

- Simulate edge-case schedules (Thanksgiving, Friday, Christmas)
- Validate `Excel_Row` assignment logic
- Confirm MatchKey merges and log diagnostics

---

## 📁 Project Structure

NFL_Pool_Automation\
├── Family Football Pool YYYY.xlsx\
├── logs/\
│&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;└── WEEKARCHIVELOGS_YYYY_MM_DD.log\
├── .env\
├── pool.py\
├── README.md\

---

## ✅ Best Practices

- Layer enhancements modularly to avoid regressions
- Restore from backup before testing new logic
- Normalize merge keys and log diagnostics before/after merge
- Archive logs weekly and purge older ones
- Validate Excel updates with test data before deploying

---

## 🔮 Future Enhancements

- Microsoft 365 fallback for email alerts
- Persistent memory for last update timestamp
- GUI dashboard for manual override and status checks
- Weekly summary email with scrape stats and Excel update status

---

## 👤 Author

David — Software Engineer, automation enthusiast, and 49ers loyalist.  
Focused on building reliable, unattended systems that just work.

---

## 📄 .env.template

# .env.template

# Path to the Excel workbook for NFL pool updates
file_path="C:\\Path\\To\\Your\\NFL_Pool_Workbook.xlsx"

# Toggle dry run mode (True = simulate without writing to Excel or sending emails)
DRY_RUN=False

# Sender email credentials (use app password for Gmail)
EMAIL_ADDRESS=your_email@gmail.com
EMAIL_PASSWORD=your_app_password

# Recipient email for error alerts
TO_EMAIL_ADDRESS=recipient_email@example.com

# SMTP configuration (default for Gmail)
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587

---

## 📄 requirements.txt

requests>=2.31.0
pandas>=2.0.0
beautifulsoup4>=4.12.2
openpyxl>=3.1.0
python-dotenv>=1.0.0
pytz>=2023.3
