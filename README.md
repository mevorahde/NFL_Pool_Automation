# 🏈 NFL Pool Automation

Automates weekly NFL betting data collection and updates an Excel workbook with spreads, team info, and game metadata. Designed for unattended execution via Windows Task Scheduler.

---

## 📌 Features

- Scrapes NFL game data from [scoresandodds.com](https://www.scoresandodds.com/nfl)
- Extracts spreads, favorites, underdogs, and game dates
- Dynamically creates or updates Excel sheets by week
- Applies conditional formatting for home teams
- Supports dry-run mode for testing
- Logs execution details to timestamped log files
- Sends email alerts on critical failures (optional)

---

## 🧰 Requirements

- Python 3.11 (recommended for compatibility)
- Packages:
  - `requests`
  - `beautifulsoup4`
  - `openpyxl`
  - `python-dotenv`
  - `pandas`
- Excel workbook with at least one sheet (used as template)
- `.env` file with:
  ```env
  file_path=C:\Path\To\Your\Workbook.xlsx
  DRY_RUN=False