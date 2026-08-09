import os
import requests
import logging
import pandas as pd
from bs4 import BeautifulSoup as Bs, BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from dotenv import dotenv_values, load_dotenv, set_key
from pathlib import Path
from datetime import datetime
import pytz
import sys
import smtplib
from email.message import EmailMessage
import gzip
import shutil
import time
from requests.exceptions import RequestException
from urllib.parse import urljoin, urlparse


# Activate '.env' file from the repository directory even when Task Scheduler
# starts the script from a different working directory.
PROJECT_DIR = Path(__file__).resolve().parent
env_path = PROJECT_DIR / '.env'
load_dotenv(dotenv_path=env_path)

# Logging setup
log_file = "nfl_spread_script.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(log_file),
        logging.StreamHandler(sys.stdout)
    ]
)

# Dry-run toggle
DRY_RUN = False

# The workbook template uses row 1 for headers; game data begins on row 2.
WORKBOOK_HEADER_ROWS = 1

NFL_URL = "https://www.scoresandodds.com/nfl"
MAX_AUTO_WEEK_ADVANCES = 4
WORKBOOK_TEMPLATE_PATH = PROJECT_DIR / "Family Football Pool Template.xlsx"


def _current_pacific_year(now=None):
    pacific = pytz.timezone("America/Los_Angeles")
    timestamp = pd.Timestamp.now(tz=pacific) if now is None else pd.Timestamp(now)
    if timestamp.tzinfo is None:
        timestamp = timestamp.tz_localize(pacific)
    else:
        timestamp = timestamp.tz_convert(pacific)
    return timestamp.year


def ensure_current_year_workbook(
    template_path=WORKBOOK_TEMPLATE_PATH,
    env_path=env_path,
    now=None,
):
    template_path = Path(template_path).resolve()
    env_path = Path(env_path).resolve()
    if not template_path.is_file():
        raise FileNotFoundError(f"Workbook template not found: {template_path}")

    year = _current_pacific_year(now)
    target_path = template_path.with_name(f"Family Football Pool {year}.xlsx")
    if target_path.exists() and not target_path.is_file():
        raise IsADirectoryError(f"Yearly workbook path is not a file: {target_path}")

    if not target_path.exists():
        shutil.copy2(template_path, target_path)
        logging.info(f"Created yearly workbook: {target_path.name}")
    else:
        logging.info(f"Using existing yearly workbook: {target_path.name}")

    env_path.touch(exist_ok=True)
    target_value = str(target_path)
    configured_value = dotenv_values(env_path).get("file_path")
    if configured_value != target_value:
        set_key(str(env_path), "file_path", target_value, quote_mode="auto")
        logging.info(f"Updated .env file_path for the {year} workbook.")

    os.environ["file_path"] = target_value
    return target_path

# NFL team abbreviations
team_abbr = {
    "49ERS": "SF", "BEARS": "CHI", "BENGALS": "CIN", "BILLS": "BUF",
    "BRONCOS": "DEN", "BROWNS": "CLE", "BUCCANEERS": "TB", "CARDINALS": "ARI",
    "CHARGERS": "LAC", "CHIEFS": "KC", "COLTS": "IND", "COMMANDERS": "WAS",
    "COWBOYS": "DAL", "DOLPHINS": "MIA", "EAGLES": "PHI", "FALCONS": "ATL",
    "GIANTS": "NYG", "JAGUARS": "JAC", "JETS": "NYJ", "LIONS": "DET",
    "PACKERS": "GB", "PANTHERS": "CAR", "PATRIOTS": "NE", "RAIDERS": "LV",
    "RAMS": "LAR", "RAVENS": "BAL", "SAINTS": "NO", "SEAHAWKS": "SEA",
    "STEELERS": "PIT", "TEXANS": "HOU", "TITANS": "TEN", "VIKINGS": "MIN"
}

official_team_abbr = {
    "ARIZONA CARDINALS": "ARI",
    "ATLANTA FALCONS": "ATL",
    "BALTIMORE RAVENS": "BAL",
    "BUFFALO BILLS": "BUF",
    "CAROLINA PANTHERS": "CAR",
    "CHICAGO BEARS": "CHI",
    "CINCINNATI BENGALS": "CIN",
    "CLEVELAND BROWNS": "CLE",
    "DALLAS COWBOYS": "DAL",
    "DENVER BRONCOS": "DEN",
    "DETROIT LIONS": "DET",
    "GREEN BAY PACKERS": "GB",
    "HOUSTON TEXANS": "HOU",
    "INDIANAPOLIS COLTS": "IND",
    "JACKSONVILLE JAGUARS": "JAC",
    "KANSAS CITY CHIEFS": "KC",
    "LAS VEGAS RAIDERS": "LV",
    "LOS ANGELES CHARGERS": "LAC",
    "LOS ANGELES RAMS": "LAR",
    "MIAMI DOLPHINS": "MIA",
    "MINNESOTA VIKINGS": "MIN",
    "NEW ENGLAND PATRIOTS": "NE",
    "NEW ORLEANS SAINTS": "NO",
    "NEW YORK GIANTS": "NYG",
    "NEW YORK JETS": "NYJ",
    "PHILADELPHIA EAGLES": "PHI",
    "PITTSBURGH STEELERS": "PIT",
    "SAN FRANCISCO 49ERS": "SF",
    "SEATTLE SEAHAWKS": "SEA",
    "TAMPA BAY BUCCANEERS": "TB",
    "TENNESSEE TITANS": "TEN",
    "WASHINGTON COMMANDERS": "WAS",
}


def send_error_email(subject, body, log_path):
    try:
        msg = EmailMessage()
        msg["From"] = os.getenv("EMAIL_ADDRESS")
        msg["To"] = os.getenv("TO_EMAIL_ADDRESS")
        msg["Subject"] = subject
        msg.set_content(body)

        # Attach log file
        with open(log_path, "rb") as f:
            msg.add_attachment(f.read(), maintype="text", subtype="plain", filename=os.path.basename(log_path))

        with smtplib.SMTP(os.getenv("SMTP_SERVER"), int(os.getenv("SMTP_PORT"))) as server:
            server.starttls()
            server.login(os.getenv("EMAIL_ADDRESS"), os.getenv("EMAIL_PASSWORD"))
            server.send_message(msg)

        logging.info("Error email sent successfully.")
    except Exception as e:
        logging.warning(f"Failed to send error email: {e}")

def send_test_email():
    subject = "NFL Automation Test Email"
    body = "This is a test email to confirm Gmail alert functionality is working."
    try:
        send_error_email(subject, body, log_file)
        logging.info("Test email sent successfully.")
    except Exception as e:
        logging.critical(f"Test email failed: {e}", exc_info=True)


def archive_log_file():
    try:
        log_path = log_file  # e.g., "nfl_spread_script.log"
        timestamp = datetime.now().strftime("%Y-%m-%d")
        archive_name = f"logs/nfl_spread_script_{timestamp}.log.gz"

        os.makedirs("logs", exist_ok=True)

        with open(log_path, "rb") as f_in:
            with gzip.open(archive_name, "wb") as f_out:
                shutil.copyfileobj(f_in, f_out)

        logging.info(f"Archived log to {archive_name}")

        # Optional: clear original log file
        open(log_path, "w").close()
        logging.info("Cleared original log file after archiving.")

    except Exception as e:
        logging.error(f"Failed to archive log file: {e}")
        send_error_email(
            subject="NFL Spread Script: ERROR - Log Archiving Failed",
            body=f"Failed to archive log file:\n{e}",
            log_path=log_path
        )


def fetch_with_retry(url, headers=None, max_retries=3, backoff_factor=2, timeout=10):
    attempt = 0
    while attempt < max_retries:
        try:
            response = requests.get(url, headers=headers, timeout=timeout)
            response.raise_for_status()
            return response
        except RequestException as e:
            attempt += 1
            wait_time = backoff_factor ** attempt
            logging.warning(f"Request failed (attempt {attempt}/{max_retries}): {e}. Retrying in {wait_time}s...")
            time.sleep(wait_time)
    logging.error(f"All {max_retries} attempts failed for URL: {url}")
    raise ConnectionError(f"Failed to fetch data from {url} after {max_retries} retries.")


def get_webpage(url, headers=None):
    try:
        response = fetch_with_retry(url, headers=headers)
        return BeautifulSoup(response.content, "html.parser")
    except Exception as e:
        logging.error(f"Failed to fetch webpage after retries: {e}")
        return None


def get_week_number(soup):
    try:
        return soup.find("div", class_="filters-week-picker") \
            .find("div", class_="selector week-picker-week") \
            .find("li", class_="menu-item active") \
            .find("span", attrs={"data-endpoint": True}).get_text()
    except AttributeError:
        msg = "Week number not found in HTML structure."
        logging.warning(msg)
        send_error_email(
            subject="NFL Spread Script: ERROR - Week Number Missing",
            body=msg,
            log_path=log_file
        )
        return "Unknown"


def extract_team_info(table, side):
    tr = table.find("tr", attrs={"data-side": side})
    name = tr.find("span", class_="team-name").find("a").find("span").get_text().upper()
    abbr = tr.find("span", class_="team-name").find("a", attrs={"data-abbr": True}).get("data-abbr")
    return name, abbr


def extract_spread_and_favorite(table):
    td = table.find("td", attrs={"data-field": "current-spread"})
    if not td:
        return "TBD", None
    span = td.find("span", class_="data-value")
    raw = span.get_text(strip=True) if span else td.get_text(strip=True).split(" ")[0]
    if raw.lower() in ["tbd", "n/a", ""]:
        return "TBD", None
    raw_clean = raw.replace("−", "-").replace("+", "").strip()
    side = td.get("data-side")
    return raw_clean, side


def extract_datetime(table):
    # Try real HTML format first
    span = table.find("span", attrs={"data-value": True})
    if span:
        try:
            return datetime.fromisoformat(span.get("data-value"))
        except Exception:
            pass

    # Fallback to mock HTML format
    try:
        date_str = table.find("div", class_="game-date").get_text(strip=True)
        return datetime.strptime(date_str, "%A, %B %d, %Y")
    except Exception as e:
        logging.warning(f"Date parsing failed: {e}")
        return None



def parse_game_card(table):
    away_name, away_abbr = extract_team_info(table, "away")
    home_name, home_abbr = extract_team_info(table, "home")
    spread, favorite_side = extract_spread_and_favorite(table)
    date_time = extract_datetime(table)
    return [away_name, spread, home_name, away_abbr, home_abbr, home_name.upper(), date_time, favorite_side]


def _extract_game_rows(soup):
    data = []
    finalized_count = 0
    pending_count = 0

    for table in soup.find_all("div", class_="event-card"):
        try:
            row = parse_game_card(table)
            if row[1] == "TBD":
                pending_count += 1
            else:
                finalized_count += 1
            data.append(row)
        except Exception as e:
            logging.warning(f"Failed to parse game card: {e}")

    return data, finalized_count, pending_count


def _all_games_started(data, now=None):
    if not data:
        return False

    kickoffs = pd.to_datetime([row[6] for row in data], errors="coerce", utc=True)
    if kickoffs.isna().any():
        return False

    now_utc = pd.Timestamp.now(tz="UTC") if now is None else pd.Timestamp(now)
    if now_utc.tzinfo is None:
        now_utc = now_utc.tz_localize("UTC")
    else:
        now_utc = now_utc.tz_convert("UTC")

    return not (kickoffs > now_utc).any()


def _next_week_url(soup, current_url):
    picker = soup.find("div", class_="selector week-picker-week")
    if picker is None:
        return None

    menu_items = picker.find_all("li", class_="menu-item")
    active_index = next(
        (
            index
            for index, item in enumerate(menu_items)
            if "active" in item.get("class", [])
        ),
        None,
    )
    if active_index is None:
        return None

    current_origin = urlparse(current_url)
    for item in menu_items[active_index + 1:]:
        endpoint = item.find("span", attrs={"data-endpoint": True})
        if endpoint is None:
            continue

        candidate = urljoin(current_url, endpoint["data-endpoint"])
        candidate_origin = urlparse(candidate)
        if (
            candidate_origin.scheme == current_origin.scheme
            and candidate_origin.netloc == current_origin.netloc
        ):
            return candidate

    return None


def scrape_nfl_data(now=None):
    url = NFL_URL
    soup = get_webpage(url)

    if not soup:
        logging.error("Failed to load NFL page.")
        send_error_email(
            subject="NFL Scraper Error: Page Load Failure",
            body="Failed to load NFL page from scoresandodds.com.",
            log_path=log_file
        )
        return None, "Unknown"

    visited_urls = {url}
    for _ in range(MAX_AUTO_WEEK_ADVANCES + 1):
        try:
            week = get_week_number(soup)
            logging.info(f"Scraping data for Week {week}")
        except Exception as e:
            logging.error(f"Failed to extract week number: {e}", exc_info=True)
            send_error_email(
                subject="NFL Scraper Error: Week Extraction Failed",
                body=f"Error extracting week number:\n{e}",
                log_path=log_file
            )
            return None, "Unknown"

        data, finalized_count, pending_count = _extract_game_rows(soup)
        if not _all_games_started(data, now=now):
            break

        next_url = _next_week_url(soup, url)
        if next_url is None or next_url in visited_urls:
            break

        logging.info(
            f"All games for Week {week} have started; loading the next NFL week."
        )
        next_soup = get_webpage(next_url)
        if not next_soup:
            break

        url = next_url
        soup = next_soup
        visited_urls.add(url)

    if _all_games_started(data, now=now):
        msg = (
            f"All games for Week {week} have started, and no future NFL week "
            "could be loaded. Aborting before the workbook update."
        )
        logging.error(msg)
        send_error_email(
            subject="NFL Scraper Error: No Future Week Available",
            body=msg,
            log_path=log_file
        )
        return None, week

    if not data:
        logging.error("No game data found.")
        send_error_email(
            subject="NFL Scraper Error: No Game Data",
            body="Scraper ran successfully but found no game data.",
            log_path=log_file
        )
        return None, week

    try:
        df = pd.DataFrame(data, columns=[
            "Team1", "Spread", "Team2", "Team1_Abbr", "Team2_Abbr",
            "Home_Team", "UTC_DateTime", "Favorite_Side"
        ])

        # ✅ Inject game_day from UTC_DateTime
        df["game_day"] = pd.to_datetime(df["UTC_DateTime"], errors="coerce").dt.day_name()

        df = apply_team_abbreviations(df)
        logging.info(f"Scraped {len(df)} games: {finalized_count} finalized, {pending_count} pending")
        return df, week
    except Exception as e:
        logging.critical(f"DataFrame construction or abbreviation failed: {e}", exc_info=True)
        send_error_email(
            subject="NFL Scraper Critical Error: DataFrame Failure",
            body=f"Critical failure during DataFrame construction or abbreviation:\n{e}",
            log_path=log_file
        )
        return None, week


def apply_team_abbreviations(df):
    for team_column, abbreviation_column in [
        ("Team1", "Team1_Abbr"),
        ("Team2", "Team2_Abbr"),
    ]:
        normalized_names = df[team_column].astype("string").str.strip().str.upper()
        mapped_abbreviations = normalized_names.map(team_abbr).fillna(
            normalized_names.map(official_team_abbr)
        )

        if abbreviation_column in df.columns:
            existing_abbreviations = (
                df[abbreviation_column].astype("string").str.strip().str.upper()
            )
            mapped_abbreviations = existing_abbreviations.where(
                existing_abbreviations.eq(mapped_abbreviations),
                mapped_abbreviations,
            )

        df[team_column] = normalized_names
        df[abbreviation_column] = mapped_abbreviations

    missing_team1 = df[df["Team1_Abbr"].isna()]["Team1"].unique()
    missing_team2 = df[df["Team2_Abbr"].isna()]["Team2"].unique()
    missing = list(missing_team1) + list(missing_team2)

    if missing:
        logging.warning(f"Missing abbreviations for: {missing}")
        if len(missing) > 3:
            send_error_email(
                subject="NFL Spread Script: ERROR - Abbreviation Mapping",
                body=f"Missing team abbreviations for: {missing}",
                log_path=log_file
            )
    return df


def extract_favorite_underdog(row):
    spread_val = row["Spread"]
    favorite_side = row["Favorite_Side"]
    team1 = row["Team1"]
    team2 = row["Team2"]
    abbr1 = row["Team1_Abbr"]
    abbr2 = row["Team2_Abbr"]

    if spread_val == "TBD" or favorite_side not in ["home", "away"]:
        return "TBD", "TBD", 0.0, None, None

    try:
        spread_float = float(spread_val)
    except ValueError:
        return "TBD", "TBD", 0.0, None, None

    if favorite_side == "home":
        spread_team = team2
        spread_abbr = abbr2
        other_team = team1
        other_abbr = abbr1
    else:
        spread_team = team1
        spread_abbr = abbr1
        other_team = team2
        other_abbr = abbr2

    if spread_float < 0:
        favorite = spread_team
        underdog = other_team
        fav_abbr = spread_abbr
        und_abbr = other_abbr
        spread_display = abs(spread_float)
    else:
        favorite = other_team
        underdog = spread_team
        fav_abbr = other_abbr
        und_abbr = spread_abbr
        spread_display = spread_float

    return favorite, underdog, spread_display, fav_abbr, und_abbr


def filter_games_by_day(df, now=None):
    pacific = pytz.timezone("America/Los_Angeles")
    now = datetime.now(pacific) if now is None else now
    now = pd.Timestamp(now)
    if now.tzinfo is None:
        now = now.tz_localize(pacific)
    now_local = now.tz_convert(pacific)
    dotw = now_local.strftime("%A")
    now = now_local.tz_convert("UTC")

    # ✅ Ensure UTC_DateTime is a datetime object
    df_working = df.copy(deep=True)
    df_working["UTC_DateTime"] = _validated_kickoff_timestamps(df_working)

    # ✅ Filter out games that have already started
    df_filtered = df_working[df_working["UTC_DateTime"] > now].copy()
    excluded = df_working[df_working["UTC_DateTime"] <= now]

    logging.info(f"Excluded {len(excluded)} played games for {dotw}:")
    for _, row in excluded.iterrows():
        game_day = (
            row["Local_DateTime"].strftime("%A")
            if pd.notna(row.get("Local_DateTime"))
            else "Unknown Day"
        )
        logging.info(f"  {row['Team1']} vs {row['Team2']} on {game_day} ({row['UTC_DateTime']})")

    if not df_filtered.empty and "Local_DateTime" in df_filtered.columns:
        earliest_game = df_filtered["Local_DateTime"].min()
        logging.info(f"Earliest remaining game is on {earliest_game.strftime('%A, %Y-%m-%d %I:%M %p')}")
    logging.info(f"Filtered games for {dotw}: {len(df_filtered)} remaining")

    return df_filtered, dotw

def get_local_day(utc_str):
    pacific = pytz.timezone("America/Los_Angeles")
    try:
        dt = datetime.strptime(utc_str, "%Y-%m-%dT%H:%M:%SZ")
        local_dt = dt.replace(tzinfo=pytz.utc).astimezone(pacific)
        return local_dt.strftime("%A")
    except Exception as e:
        logging.warning(f"Failed to parse UTC datetime: {e}")
        send_error_email(
            subject="NFL Spread Script: ERROR - UTC Day Conversion",
            body=f"Failed to convert UTC string: {utc_str}\nError: {e}",
            log_path=log_file
        )
        return "Unknown"



def get_local_datetime(utc_str):
    pacific = pytz.timezone("America/Los_Angeles")
    try:
        dt = datetime.strptime(utc_str, "%Y-%m-%dT%H:%M:%SZ")
        return dt.replace(tzinfo=pytz.utc).astimezone(pacific)
    except Exception as e:
        logging.warning(f"Failed to convert UTC datetime: {e}")
        send_error_email(
            subject="NFL Spread Script: ERROR - UTC Datetime Conversion",
            body=f"Failed to convert UTC string: {utc_str}\nError: {e}",
            log_path=log_file
        )
        return None


def update_excel(wk_number, df_filtered, dotw):
    try:
        file = os.getenv("file_path")
        wb = load_workbook(filename=file)
        all_sheets = wb.sheetnames
        template = wb.worksheets[0]

        home_fill = PatternFill(start_color='F4B084', end_color='F4B084', fill_type='solid')
        clear_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        night_fill = PatternFill(start_color='00B0F0', end_color='00B0F0', fill_type='solid')  # SNF/MNF highlight

        # Create or overwrite sheet
        if wk_number in all_sheets:
            new_wk_sheet = wb[wk_number]
            logging.info(f"Overwriting existing sheet: {wk_number}")
        else:
            template_copy = wb.copy_worksheet(template)
            template_copy.title = wk_number
            new_wk_sheet = wb[wk_number]
            logging.info(f"Created new sheet: {wk_number}")

        # Activate the new sheet
        for sheet in wb:
            sheet.views.sheetView[0].tabSelected = False
        wb.active = new_wk_sheet
        new_wk_sheet.views.sheetView[0].tabSelected = True

        # Defensive check for Excel_Row
        if "Excel_Row" not in df_filtered.columns:
            msg = "Excel_Row column missing from DataFrame. Aborting Excel update."
            logging.critical(msg)
            send_error_email(
                subject="NFL Excel Update Critical Error",
                body=msg,
                log_path=log_file
            )
            return

        dotw = dotw.strip().title()
        locked_game_days = [dotw]
        logging.info(f"Locked game days for today ({dotw}): {locked_game_days}")

        df = df_filtered.copy()
        df = df[df["Excel_Row"].notna()]
        df["Excel_Row"] = df["Excel_Row"].astype(int)
        df["game_day"] = df["game_day"].astype(str).str.strip().str.title()

        # Identify SNF and MNF rows
        sunday_games = df[df["game_day"] == "Sunday"]
        monday_games = df[df["game_day"] == "Monday"]

        latest_sunday = sunday_games["UTC_DateTime"].max() if not sunday_games.empty else None
        snf_rows = sunday_games[sunday_games["UTC_DateTime"] == latest_sunday]["Excel_Row"].tolist()
        mnf_rows = monday_games["Excel_Row"].tolist()
        night_rows = set(snf_rows + mnf_rows)

        # Filter out locked rows before clearing
        df_unlocked = df[~df["game_day"].isin(locked_game_days)]
        rows_to_update = df_unlocked["Excel_Row"].unique()

        # Clear all rows that will be updated
        for row in rows_to_update:
            logging.info(f"Clearing row {row}")
            for col in [3, 4, 5, 9, 11, 14, 15]:
                new_wk_sheet.cell(row=row, column=col).value = None
                new_wk_sheet.cell(row=row, column=col).fill = clear_fill

        # Update each row with FAVORITE vs UNDERDOG
        for _, row in df_unlocked.iterrows():
            try:
                excel_row = int(row["Excel_Row"])
                game_day = row["game_day"]

                logging.debug(f"Row {excel_row}: game_day={game_day}, dotw={dotw}")

                favorite, underdog, spread_val, fav_abbrev, und_abbr = extract_favorite_underdog(row)
                ht = row["Home_Team"]

                logging.info(f"Updating row {excel_row}: {favorite} vs {underdog}, spread {spread_val}")

                new_wk_sheet.cell(row=excel_row, column=3).value = favorite
                new_wk_sheet.cell(row=excel_row, column=4).value = spread_val
                new_wk_sheet.cell(row=excel_row, column=5).value = underdog
                new_wk_sheet.cell(row=excel_row, column=9).value = fav_abbrev
                new_wk_sheet.cell(row=excel_row, column=11).value = und_abbr

                cell_c = new_wk_sheet.cell(row=excel_row, column=3)
                cell_e = new_wk_sheet.cell(row=excel_row, column=5)

                if favorite == ht:
                    cell_c.fill = home_fill
                    cell_e.fill = clear_fill
                elif underdog == ht:
                    cell_e.fill = home_fill
                    cell_c.fill = clear_fill
                else:
                    cell_c.fill = clear_fill
                    cell_e.fill = clear_fill
                    logging.warning(f"Home team '{ht}' not matched in favorite/underdog for row {excel_row}")

                # Apply SNF/MNF highlight
                if excel_row in night_rows:
                    for col in [14, 15]:
                        new_wk_sheet.cell(row=excel_row, column=col).fill = night_fill
                else:
                    for col in [14, 15]:
                        new_wk_sheet.cell(row=excel_row, column=col).fill = clear_fill

            except Exception as e:
                logging.warning(f"Error updating row {row.get('Excel_Row', 'Unknown')}: {e}")

        wb.save(file)
        logging.info(f"Excel updated and saved for {wk_number}")

    except Exception as e:
        logging.critical(f"Excel update failed: {e}", exc_info=True)
        try:
            logging.error(f"Available sheets: {wb.sheetnames}")
        except:
            logging.error("Workbook not loaded—no sheet names available.")

        send_error_email(
            subject="NFL Excel Update Critical Error",
            body=f"Excel update failed:\n{e}",
            log_path=log_file
        )

def verify_matchkey_alignment(df_full, df_filtered):
    full_keys = df_full["MatchKey"].drop_duplicates()
    filtered_keys = df_filtered["MatchKey"].drop_duplicates()
    unmatched = filtered_keys[~filtered_keys.isin(full_keys)]

    if unmatched.empty:
        logging.info("All MatchKeys in df_filtered matched df_full.")
    else:
        logging.warning("Unmatched MatchKeys in df_filtered:")
        for key in unmatched:
            logging.warning(f"  {key}")
        send_error_email(
            subject="NFL Spread Script: ERROR - MatchKey Mismatch",
            body=f"Unmatched MatchKeys found:\n" + "\n".join(unmatched),
            log_path=log_file
        )

def normalize_matchkeys(df):
    df["Team1"] = df["Team1"].astype(str).str.strip().str.upper()
    df["Team2"] = df["Team2"].astype(str).str.strip().str.upper()
    df["MatchKey"] = (df["Team1"] + " vs " + df["Team2"]).str.strip().str.upper()
    return df

def _validated_kickoff_timestamps(df):
    if "UTC_DateTime" not in df.columns:
        raise ValueError("UTC_DateTime column is required for Excel row assignment")

    timestamps = pd.to_datetime(df["UTC_DateTime"], errors="coerce", utc=True)
    invalid = timestamps.isna()
    if invalid.any():
        invalid_rows = ", ".join(str(index) for index in df.index[invalid])
        raise ValueError(
            "UTC_DateTime contains invalid or missing kickoff timestamps at "
            f"DataFrame rows: {invalid_rows}"
        )
    return timestamps


def assign_excel_rows(df, header_rows=WORKBOOK_HEADER_ROWS):
    """Return stable Excel rows for every game in full scraped-schedule order."""
    if isinstance(header_rows, bool) or not isinstance(header_rows, int) or header_rows < 0:
        raise ValueError("header_rows must be a non-negative integer")

    _validated_kickoff_timestamps(df)
    first_game_row = header_rows + 1
    return pd.Series(
        range(first_game_row, first_game_row + len(df)),
        index=df.index,
        dtype="int64",
        name="Excel_Row",
    )


def prepare_schedule_for_excel(df, now=None, header_rows=WORKBOOK_HEADER_ROWS):
    """Assign stable rows from the full schedule, then remove started games."""
    prepared = df.copy(deep=True)
    prepared["Excel_Row"] = assign_excel_rows(prepared, header_rows=header_rows)
    return filter_games_by_day(prepared, now=now)

def main():
    logging.info("Starting NFL pool automation...")

    try:
        ensure_current_year_workbook()

        # ✅ Scrape and normalize
        df_raw, week_label = scrape_nfl_data()

        if df_raw is None or not isinstance(df_raw, pd.DataFrame):
            msg = "Scraping failed or returned invalid data. Aborting pipeline."
            logging.critical(msg)
            send_error_email(
                subject="NFL Automation Critical Error: Scraping Failed",
                body=msg,
                log_path=log_file
            )
            return

        logging.info(f"Scraping data for Week {week_label}")
        logging.info(f"Scraped {len(df_raw)} games")

        # ✅ Localize game_day to Pacific Time
        df_raw["game_day"] = (
            pd.to_datetime(df_raw["UTC_DateTime"], errors="coerce", utc=True)
            .dt.tz_convert("America/Los_Angeles")
            .dt.day_name()
        )

        # ✅ Preview game_day assignments
        logging.info("Preview of game_day assignments:")
        logging.info(df_raw[["Team1", "Team2", "UTC_DateTime", "game_day"]].to_string(index=False))

        df_raw = normalize_matchkeys(df_raw)

        # Assign rows from the complete scraped schedule before filtering so a
        # matchup keeps the same workbook row as earlier games begin.
        df_filtered, dotw = prepare_schedule_for_excel(df_raw)

        # ✅ Confirm Excel_Row exists
        if "Excel_Row" not in df_filtered.columns:
            msg = "Excel_Row missing from filtered DataFrame. Aborting."
            logging.critical(msg)
            send_error_email(
                subject="NFL Automation Critical Error: Excel_Row Missing",
                body=msg,
                log_path=log_file
            )
            return

        # ✅ Confirm all rows have Excel_Row
        unmatched = df_filtered[df_filtered["Excel_Row"].isna()]
        if not unmatched.empty:
            logging.warning(f"Unmatched rows after filtering: {len(unmatched)}")
            for _, row in unmatched.iterrows():
                logging.warning(f"  {row['Team1']} vs {row['Team2']} — MatchKey: {row['MatchKey']}")
        else:
            logging.info("[OK] All filtered games have Excel_Row assigned.")

        # ✅ Preview post-filter
        logging.info("Post-filter preview:")
        preview_cols = ["Team1", "Team2", "MatchKey", "Excel_Row"]
        logging.info(df_filtered[preview_cols].to_string(index=False))

        # ✅ Update Excel
        update_excel(week_label, df_filtered, dotw)

        logging.info("NFL pool automation complete.")

    except Exception as e:
        logging.critical(f"Unhandled exception in main(): {e}", exc_info=True)
        send_error_email(
            subject="NFL Automation Crash",
            body=f"Unhandled exception in main():\n{e}",
            log_path=log_file
        )

if __name__ == "__main__":
    main()
