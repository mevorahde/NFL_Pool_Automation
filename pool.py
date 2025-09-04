import os
import requests
import logging
import pandas as pd
from bs4 import BeautifulSoup as Bs
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from dotenv import load_dotenv
from pathlib import Path

# Activate '.env' file
load_dotenv()
load_dotenv(verbose=True)
env_path = Path('.') / '.env'
load_dotenv(dotenv_path=env_path)

# Logging setup
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Dry-run toggle
DRY_RUN = False

# NFL team abbreviations
team_abbr = {
    "49ERS": "SF", "BEARS": "CHI", "BENGALS": "CIN", "BILLS": "BUF",
    "BRONCOS": "DEN", "BROWNS": "CLE", "BUCCANEERS": "TB", "CARDINALS": "ARI",
    "CHARGERS": "LAC", "CHIEFS": "KC", "COLTS": "IND", "COMMANDERS": "WAS",
    "COWBOYS": "DAL", "DOLPHINS": "MIA", "EAGLES": "PHI", "FALCONS": "ATL",
    "GIANTS": "NYG", "JAGUARS": "JAX", "JETS": "NYJ", "LIONS": "DET",
    "PACKERS": "GB", "PANTHERS": "CAR", "PATRIOTS": "NE", "RAIDERS": "LV",
    "RAMS": "LAR", "RAVENS": "BAL", "SAINTS": "NO", "SEAHAWKS": "SEA",
    "STEELERS": "PIT", "TEXANS": "HOU", "TITANS": "TEN", "VIKINGS": "MIN"
}

def get_webpage(url):
    try:
        r = requests.get(url)
        r.raise_for_status()
        return Bs(r.content, "html.parser")
    except Exception as e:
        logging.error(f"Failed to load webpage: {e}")
        return None

def get_week_number(soup):
    try:
        return soup.find("div", class_="filters-week-picker") \
                   .find("div", class_="selector week-picker-week") \
                   .find("li", class_="menu-item active") \
                   .find("span", attrs={"data-endpoint": True}).get_text()
    except AttributeError:
        logging.warning("Week number not found.")
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
    side = td.get("data-side")  # 'home' or 'away'
    return raw_clean, side

def extract_datetime(table):
    span = table.find("span", attrs={"data-value": True})
    return span.get("data-value") if span else "Unknown"

def parse_game_card(table):
    away_name, away_abbr = extract_team_info(table, "away")
    home_name, home_abbr = extract_team_info(table, "home")
    spread, favorite_side = extract_spread_and_favorite(table)
    date_time = extract_datetime(table)

    return [away_name, spread, home_name, away_abbr, home_abbr, home_name.upper(), date_time, favorite_side]

def scrape_nfl_data():
    url = "https://www.scoresandodds.com/nfl"
    soup = get_webpage(url)
    if not soup:
        logging.error("Failed to load NFL page.")
        return None, "Unknown"

    week = get_week_number(soup)
    logging.info(f"Scraping data for Week {week}")

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

    if not data:
        logging.error("No game data found.")
        return None, week

    df = pd.DataFrame(data, columns=[
        "Team1", "Spread", "Team2", "Team1_Abbr", "Team2_Abbr",
        "Home_Team", "UTC_DateTime", "Favorite_Side"
    ])
    df = apply_team_abbreviations(df)

    logging.info(f"Scraped {len(df)} games: {finalized_count} finalized, {pending_count} pending")
    return df, week

def apply_team_abbreviations(df):
    df["Team1"] = df["Team1"].str.upper()
    df["Team2"] = df["Team2"].str.upper()
    df["Team1_Abbr"] = df["Team1"].map(team_abbr)
    df["Team2_Abbr"] = df["Team2"].map(team_abbr)

    missing_team1 = df[df["Team1_Abbr"].isna()]["Team1"].unique()
    missing_team2 = df[df["Team2_Abbr"].isna()]["Team2"].unique()

    if len(missing_team1) > 0 or len(missing_team2) > 0:
        logging.warning(f"Missing abbreviations for: {missing_team1.tolist() + missing_team2.tolist()}")

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



def update_excel(wk_number, df, dotw):
    try:
        file = os.getenv("file_path")
        wb = load_workbook(filename=file)
        all_sheets = wb.sheetnames
        template = wb.worksheets[0]

        home_fill = PatternFill(start_color='F4B084', end_color='F4B084', fill_type='solid')
        clear_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        num_games = len(df)

        if wk_number in all_sheets:
            new_wk_sheet = wb[wk_number]
            logging.info(f"Overwriting existing sheet: {wk_number}")
            for sheet in wb:
                wb[sheet.title].views.sheetView[0].tabSelected = False
            wb.active = new_wk_sheet

            for r in range(num_games):
                try:
                    row = df.iloc[r]
                    favorite, underdog, spread_val, fav_abbrev, und_abbrev = extract_favorite_underdog(row)
                    ht = row["Home_Team"]

                    new_wk_sheet.cell(row=r + 2, column=3).value = favorite
                    new_wk_sheet.cell(row=r + 2, column=4).value = spread_val
                    new_wk_sheet.cell(row=r + 2, column=5).value = underdog
                    new_wk_sheet.cell(row=r + 2, column=9).value = fav_abbrev
                    new_wk_sheet.cell(row=r + 2, column=11).value = und_abbrev

                    cell_c = new_wk_sheet.cell(row=r + 2, column=3)
                    cell_e = new_wk_sheet.cell(row=r + 2, column=5)

                    if favorite == ht:
                        cell_c.fill = home_fill
                        cell_e.fill = clear_fill
                    elif underdog == ht:
                        cell_e.fill = home_fill
                        cell_c.fill = clear_fill
                    else:
                        cell_c.fill = clear_fill
                        cell_e.fill = clear_fill
                        logging.warning(f"Home team '{ht}' not matched in favorite/underdog for row {r + 2}")

                    for col in [14, 15]:
                        new_wk_sheet.cell(row=r + 2, column=col).fill = clear_fill

                    if DRY_RUN:
                        logging.info(f"[Dry-run] Row {r + 2}: Favorite={favorite}, Spread={spread_val}, Underdog={underdog}, Home={ht}")

                except Exception as e:
                    logging.warning(f"Error updating row {r + 2}: {e}")

            if not DRY_RUN:
                wb.save(file)
                logging.info(f"Excel updated and saved for {wk_number}")
            else:
                logging.info("Dry-run mode: Excel not saved.")
        else:
            template_copy = wb.copy_worksheet(template)
            new_wk_sheet = wb['Template Copy']
            new_wk_sheet.title = wk_number
            logging.info(f"Created new sheet: {wk_number}")
    except Exception as e:
        logging.critical(f"Excel update failed: {e}", exc_info=True)

if __name__ == "__main__":
    df, wk_number = scrape_nfl_data()
    if df is not None and wk_number != "Unknown":
        dotw = "Sunday"  # You can enhance this later using UTC_DateTime
        update_excel(wk_number, df, dotw)
    else:
        logging.error("Scraping failed or week number unavailable. Excel update aborted.")