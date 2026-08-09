# NFL Pool Excel Automation

This Python application grew from a personal family football pool: weekly NFL matchups and betting spreads had to be copied into a structured Excel workbook accurately and repeatedly. The project automates that workflow while preserving stable workbook rows as games begin, handling NFL schedule edge cases, and surfacing operational failures through logs and optional email alerts.

The repository demonstrates production-oriented data parsing, timezone-aware filtering, deterministic spreadsheet mapping, defensive validation, and fixture-based testing. It is a portfolio project and should be reviewed and adapted before use with an important workbook.

## Technologies and engineering skills

- Python 3 and `pytest`
- HTTP requests and retry/backoff handling with `requests`
- HTML parsing with Beautiful Soup
- Tabular transformations and validation with `pandas`
- Excel workbook updates and cell formatting with `openpyxl`
- UTC/Pacific timezone conversion with `pytz`
- Environment-based configuration with `python-dotenv`
- SMTP error notifications with TLS
- Deterministic row identity, input validation, and non-mutating DataFrame preparation
- Saved HTML fixtures and tracked workbook data for repeatable edge-case tests

## Architecture and data flow

```text
Scores-and-Odds NFL page
          |
          v
HTTP fetch with retries and timeout
          |
          v
Beautiful Soup event-card parsing
  - week number
  - teams and abbreviations
  - spread and favorite side
  - kickoff timestamp
          |
          v
pandas normalization and validation
  - official NFL name mapping
  - Pacific game-day classification
  - normalized matchup key
          |
          v
Stable Excel-row assignment on the complete schedule
          |
          v
Started-game filtering and current-day locking
          |
          v
openpyxl weekly-sheet update and workbook save
          |
          +--> local log
          +--> SMTP error notification when configured
```

`main()` performs the live scrape, prepares the complete schedule, filters it, and passes the remaining rows to the Excel update path. Parsing, mapping, row assignment, filtering, and workbook updates are separate functions so their responsibilities can be inspected independently.

## Verified features

- Retries failed HTTP requests with a timeout and exponential backoff.
- Parses the active week and game cards from the expected source-site HTML structure.
- Handles finalized spreads and `TBD` spread data.
- Normalizes surrounding whitespace and case for team names.
- Maps supported nicknames and all 32 explicit current NFL full team names; malformed names remain unmapped and trigger existing diagnostics.
- Converts kickoff data to timezone-aware UTC values and derives game days in `America/Los_Angeles`.
- Rejects missing or invalid kickoff timestamps before row assignment.
- Includes Friday games and tested Thanksgiving and Saturday schedules.
- Preserves stable matchup rows when earlier games have started.
- Creates a weekly sheet from the workbook's first worksheet when the requested week sheet does not exist, or updates the existing week sheet.
- Writes favorite, spread, underdog, and abbreviations to the workbook's established columns.
- Applies home-team fills and highlights the latest Sunday game plus Monday games in the designated columns.
- Writes operational messages to `nfl_spread_script.log` and attempts SMTP error notification for several critical failures.

## Stable Excel-row rule

The workbook template reserves row 1 for headers, so game data begins on row 2. Excel rows are assigned from the complete scraped schedule in its existing order:

```text
Excel_Row = complete-schedule position + 2
```

Row assignment happens before games that have already started are filtered out. Consequently, removing an earlier game does not renumber later matchups, and the same matchup retains its `Excel_Row` across repeated runs as long as the complete source schedule and its order remain consistent.

The assignment function validates every kickoff timestamp and returns a new indexed series. Schedule preparation and filtering operate on DataFrame copies so the caller is not unexpectedly given an `Excel_Row` column or converted timestamp column.

## Current locking behavior

Locking is based on the Pacific calendar day when the script runs:

1. Games whose kickoff is at or before the current time are removed from the update DataFrame.
2. Rows are assigned before that filtering, preserving their original workbook positions.
3. During the workbook update, every remaining game whose Pacific `game_day` equals the current Pacific weekday is excluded from clearing and writing.
4. Remaining games on other days are eligible for updates.

This is a whole-day update lock, not a per-game kickoff lock inside the Excel-writing function. For example, a Saturday run locks remaining Saturday games but does not lock Sunday games merely because the run occurs on Saturday.

## Installation

Use a dedicated virtual environment:

```bash
python -m venv .venv
```

Activate the environment using the command appropriate for your shell, then install the repository dependencies:

```bash
python -m pip install -r requirements.txt
```

## Configuration

Create a local `.env` file in the repository root. Use placeholders until you are ready to configure a private environment:

```dotenv
file_path=<managed-automatically-when-the-script-runs>
EMAIL_ADDRESS=<gmail-sender-address>
EMAIL_PASSWORD=<gmail-app-password>
TO_EMAIL_ADDRESS=<notification-recipient-address>
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
```

At startup, the application determines the current Pacific year and ensures that `Family Football Pool YYYY.xlsx` exists beside `pool.py`. If it is missing, the application creates it from `Family Football Pool Template.xlsx`. An existing yearly workbook is never overwritten. The application then updates only the `.env` `file_path` entry to the resolved yearly workbook and uses its first worksheet as the template for new week sheets.

There is no functional or configurable dry-run mode. Running the production entry point can access the live website, write the configured workbook, and attempt email notifications on failures.

## Usage

Before the first production run:

1. Review `Family Football Pool Template.xlsx` and confirm game rows begin at row 2.
2. If the current year's workbook already exists, make a separate backup of it.
3. Review the configured sender and recipient settings.
4. Close applications that may lock the workbook file.

Run the application:

```bash
python pool.py
```

The command performs a live fetch and may overwrite cells in an existing week sheet before saving the same workbook file.

## Tests

Run the complete suite from the repository root:

```bash
python -m pytest -q
```

Current verified result:

```text
42 passed
```

The suite uses saved HTML fixtures and tracked test-workbook data. It covers HTML structure, automatic advancement from an expired slate to the next NFL week, week and datetime parsing, official and malformed team-name mapping, Friday/Thanksgiving/Saturday schedules, deterministic Excel rows, simulated runs at different times, invalid timestamps, caller DataFrame preservation, yearly workbook creation without overwriting existing data, `.env` path updates, and the production connections.

The tests do not exercise every operational workflow. In particular, they do not perform a live source-site request, send SMTP email, or update the `.env`-selected production workbook.

## Project structure

```text
.
|-- pool.py                          # Production scrape, transformation, and Excel update flow
|-- README.md                        # Project documentation
|-- requirements.txt                 # Runtime and test dependencies
|-- .env.template                    # Configuration example; review before copying
|-- Family Football Pool Template.xlsx # Sanitized example workbook template
`-- tests/
    |-- test_excel_row_assignment.py # Stable-row and filtering tests
    |-- test_edge_cases.py           # Schedule, mapping, and tracked-data tests
    |-- test_html_structure.py       # Fixture parser-contract tests
    |-- generate_test_schedule.py    # Legacy test-data generator
    |-- test_runs/test_schedule.xlsx # Tracked schedule-order data
    |-- test_schedule.xlsx           # Legacy duplicate test data
    `-- mock_html/                    # Saved schedule and holiday HTML fixtures
```

## Known limitations and operational cautions

- The scraper depends on the source site's current HTML structure. Markup changes can break parsing or silently reduce the games returned.
- Stable row identity depends on receiving the complete schedule in a consistent order. A missing or reordered source game can change positional mapping.
- Workbook writes are neither atomic nor transactional. A failure during save can leave the workbook unavailable or damaged; keep backups and use a disposable copy for initial validation.
- The workbook update logic assumes a fixed sheet and column layout. It is not a general Excel schema mapper.
- Current-day locking is day-level. It does not implement a configurable lead time or a separate lock policy for tomorrow's games.
- There is no functional dry-run mode.
- `archive_log_file()` exists but is not called by the production flow, so automatic log archiving is not active.
- The application does not create automatic workbook backups.
- Live website, SMTP, and production-workbook integration are not covered by the local test suite.
- No CI workflow is currently included; test execution is local.

## Privacy and security

- Keep `.env` out of version control. The repository's ignore rules cover `.env`, but verify staged files before every commit.
- Use a Gmail app password rather than an account password. Store it only in the local environment file or a suitable secret manager.
- Never paste credentials into source code, tests, logs, screenshots, issues, or documentation.
- Restrict access to the workbook and logs; they may contain participant, matchup, or operational information.
- Rotate an app password immediately if it is exposed.
- Use a dedicated sender account with the minimum access necessary for notifications.

## Source-site disclaimer

This project relies on HTML made available by a third-party sports information site. It is not a supported API integration. Users are responsible for reviewing and complying with the site's terms, robots policy, access limits, and applicable rules. Keep request frequency modest, and expect the parser to require maintenance when the site's markup changes.

## Origin

The project began as automation for a personal family football pool and evolved into a practical exercise in reliable data ingestion, deterministic spreadsheet updates, and regression testing around real NFL scheduling edge cases.
