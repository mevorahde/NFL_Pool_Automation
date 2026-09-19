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
Started-game filtering by kickoff timestamp
          |
          v
Reconcile matchup identities with existing worksheet rows
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

The workbook template reserves row 1 for headers, so game data begins on row 2. Schedule preparation first assigns provisional rows from the complete schedule before filtering started games.

Before writing, `align_excel_rows_to_worksheet()` reconciles each remaining game with the existing worksheet using an order-independent pair of normalized team abbreviations. Existing matchups keep their rows even if the source reorders games or reverses favorite/underdog order. New matchups use available rows; existing identities remain reserved. Duplicate or incomplete identities, duplicate current matchups, and insufficient space cause alignment to fail before the update writes game rows.

The assignment function validates every kickoff timestamp and returns a new indexed series. Schedule preparation and filtering operate on DataFrame copies so the caller is not unexpectedly given an `Excel_Row` column or converted timestamp column.

## Current locking behavior

Updates are filtered by each game's timezone-aware kickoff timestamp:

1. Games whose kickoff is at or before the current time are removed from the update DataFrame.
2. Remaining games are aligned to existing worksheet matchup rows before clearing and writing.
3. Games that have not started remain eligible, including games later on the current day.
4. Existing rows for started games are left untouched by the game-update path. Participant entry values in columns N/O are preserved while applicable highlighting may change.

For example, a Thursday-morning run preserves a completed Wednesday game's row but can update Thursday evening's game. Friday and Saturday runs preserve Wednesday and Thursday games whose kickoff has passed. This is not a whole-day lock; callers using the Excel-writing function directly must first apply the normal filtering pipeline.

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

The GitHub Actions [Tests workflow](.github/workflows/tests.yml) also runs the suite. Consult its latest run for the current test count and result.

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
    |-- test_stable_workbook_rows.py # Existing matchup identity and started-game protection
    |-- test_edge_cases.py           # Schedule, mapping, and tracked-data tests
    |-- test_html_structure.py       # Fixture parser-contract tests
    |-- generate_test_schedule.py    # Legacy test-data generator
    |-- test_runs/test_schedule.xlsx # Tracked schedule-order data
    |-- test_schedule.xlsx           # Legacy duplicate test data
    `-- mock_html/                    # Saved schedule and holiday HTML fixtures
```

## Known limitations and operational cautions

- The scraper depends on the source site's current HTML structure. Markup changes can break parsing or silently reduce the games returned.
- Existing row identity depends on valid, unique team-abbreviation pairs already stored in the worksheet. The tool rejects ambiguous mappings; it does not automatically repair historical workbook corruption.
- Workbook writes are neither atomic nor transactional. A failure during save can leave the workbook unavailable or damaged; keep backups and use a disposable copy for initial validation.
- The workbook update logic assumes a fixed sheet and column layout. It is not a general Excel schema mapper.
- Started-game protection uses kickoff time, not a configurable lead time or a whole-day lock.
- There is no functional dry-run mode.
- `archive_log_file()` exists but is not called by the production flow, so automatic log archiving is not active.
- The application does not create automatic workbook backups.
- Live website, SMTP, and production-workbook integration are not covered by the local test suite.
- CI uses isolated tests and fixtures, not live website, email, or production-workbook checks.

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

## License

The original software and documentation in this repository are available under the [MIT License](LICENSE). Third-party source-site content, trademarks, and tracked factual data remain the property of their respective owners.
