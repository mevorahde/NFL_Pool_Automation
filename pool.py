import requests
from bs4 import BeautifulSoup as Bs
import pandas as pd
from openpyxl import load_workbook
import os
from dotenv import load_dotenv
from pathlib import Path

# Activate '.env' file
load_dotenv()
load_dotenv(verbose=True)
env_path = Path('.') / '.env'
load_dotenv(dotenv_path=env_path)

# Load the webpage content
url = "https://www.scoresandodds.com/nfl"
r = requests.get(url)

# Convert to a beautiful soup object
webpage = Bs(r.content, "html.parser")

wk_number = webpage.find("div", attrs={"class": "filters-week-picker"}) \
    .find("div", attrs={"class": "selector week-picker-week"}).find("li", attrs={"class": "menu-item active"}) \
    .find("span", attrs={"data-endpoint": True}).get_text()

# Set Column names
headers = ["Team 1", "Spread", "Team 2", "Team 1 Abbreviation", "Team 2 Abbreviation", "Home Team"]
column_names = [c for c in headers]

data = []
for table in webpage.find_all("div", attrs={"class": "event-card"}):
    find_favorite_tm = table.find("td", attrs={"data-field": "current-spread", "data-side": True})
    favorite_tm = find_favorite_tm.get("data-side")
    if favorite_tm == "home":
        tm1_name = table.find("tr", attrs={"data-side": "home"}).find("span", attrs={"class": "team-name"}) \
            .find("a").find("span").get_text()
        home_tm = tm1_name
        spread = table.find("td", attrs={"data-field": "current-spread"}).find("span",
                                                                               attrs={
                                                                                   "class": "data-value"}) \
            .get_text().strip()
        tm2_name = table.find("tr", attrs={"data-side": "away"}).find("span", attrs={"class": "team-name"}) \
            .find("a").find("span").get_text()
        tm1_abbr_field = table.find("tr", attrs={"data-side": "home"}).find("span", attrs={"class": "team-name"}) \
            .find("a", attrs={"data-abbr": True})
        tm1_abbr = tm1_abbr_field.get('data-abbr')
        tm2_abbr_field = table.find("tr", attrs={"data-side": "away"}).find("span", attrs={"class": "team-name"}) \
            .find("a", attrs={"data-abbr": True})
        tm2_abbr = tm2_abbr_field.get("data-abbr")
        row = [tm1_name.upper(), spread, tm2_name.upper(), tm1_abbr, tm2_abbr, home_tm]
        data.append(row)
    else:
        tm1_name = table.find("tr", attrs={"data-side": "away"}).find("span", attrs={"class": "team-name"}) \
            .find("a").find("span").get_text()
        spread = table.find("td", attrs={"data-field": "current-spread"}).find("span",
                                                                               attrs={
                                                                                   "class": "data-value"})
        grab_away_spread = str(find_favorite_tm.get_text().strip())
        away_spread_split = grab_away_spread.split(" ")
        final_away_spread = away_spread_split[0]
        tm2_name = table.find("tr", attrs={"data-side": "home"}).find("span", attrs={"class": "team-name"}) \
            .find("a").find("span").get_text()
        home_tm = tm2_name
        tm1_abbr_field = table.find("tr", attrs={"data-side": "away"}).find("span", attrs={"class": "team-name"}) \
            .find("a", attrs={"data-abbr": True})
        tm1_abbr = tm1_abbr_field.get('data-abbr')
        tm2_abbr_field = table.find("tr", attrs={"data-side": "home"}).find("span", attrs={"class": "team-name"}) \
            .find("a", attrs={"data-abbr": True})
        tm2_abbr = tm2_abbr_field.get("data-abbr")
        row = [tm1_name.upper(), final_away_spread, tm2_name.upper(), tm1_abbr, tm2_abbr, home_tm]
        data.append(row)

df = pd.DataFrame(data, columns=column_names)
# with pd.option_context('display.max_rows', None, 'display.max_columns', None):  # more options can be specified also
#     print(df)

file = os.getenv("file_path")

wb = load_workbook(filename=file)
all_sheets = wb.sheetnames

template = wb.worksheets[0]

if wk_number not in all_sheets:
    wb.create_sheet("{}".format(wk_number))
    template_copy = wb.copy_worksheet(template)
    wb.save(file)

print(all_sheets)
