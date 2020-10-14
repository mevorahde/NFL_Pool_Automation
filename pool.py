import requests
from bs4 import BeautifulSoup as Bs
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
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

data = []
for table in webpage.find_all("div", attrs={"class": "event-card"}):
    find_favorite_tm = table.find("td", attrs={"data-field": "current-spread", "data-side": True})
    favorite_tm = find_favorite_tm.get("data-side")
    if favorite_tm == "home":
        tm1_name = table.find("tr", attrs={"data-side": "home"}).find("span", attrs={"class": "team-name"}) \
            .find("a").find("span").get_text()
        home_tm = tm1_name.upper()
        spread = table.find("td", attrs={"data-field": "current-spread"}) \
            .find("span", attrs={"class": "data-value"}).get_text().strip().replace("-", "")
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
        final_away_spread = away_spread_split[0].replace("-", "")
        tm2_name = table.find("tr", attrs={"data-side": "home"}).find("span", attrs={"class": "team-name"}) \
            .find("a").find("span").get_text()
        home_tm = tm2_name.upper()
        tm1_abbr_field = table.find("tr", attrs={"data-side": "away"}).find("span", attrs={"class": "team-name"}) \
            .find("a", attrs={"data-abbr": True})
        tm1_abbr = tm1_abbr_field.get('data-abbr')
        tm2_abbr_field = table.find("tr", attrs={"data-side": "home"}).find("span", attrs={"class": "team-name"}) \
            .find("a", attrs={"data-abbr": True})
        tm2_abbr = tm2_abbr_field.get("data-abbr")
        row = [tm1_name.upper(), final_away_spread, tm2_name.upper(), tm1_abbr, tm2_abbr, home_tm]
        data.append(row)

# Column Data
favorite_teams = [fav_teams[0] for fav_teams in data]
spreads_string = [spreads[1] for spreads in data]
spreads_int = []
for i in range(0, len(spreads_string)):
    if spreads_string[i] == '':
        spreads_string[i] = 0
    spreads_string[i] = spreads_int.append(float(spreads_string[i]))
underdog_teams = [under_teams[2] for under_teams in data]
fav_abbr = [fav_team_abbr[3] for fav_team_abbr in data]
under_abbr = [under_team_abbr[4] for under_team_abbr in data]
num_games = len(favorite_teams)
home_team = [home_team[5] for home_team in data]

# Excel Info and processes
file = os.getenv("file_path")
wb = load_workbook(filename=file)
all_sheets = wb.sheetnames
template = wb.worksheets[0]
blue_fill = PatternFill(start_color='00B0F0', end_color='00B0F0', fill_type='solid')
home_fill = PatternFill(start_color='F4B084', end_color='F4B084', fill_type='solid')

if wk_number not in all_sheets:
    template_copy = wb.copy_worksheet(template)
    new_wk_sheet = wb['Template Copy']
    new_wk_sheet.title = wk_number

    for r in range(0, num_games - 1):
        ht = home_team[r]
        new_wk_sheet.cell(row=r + 2, column=3).value = favorite_teams[r]
        if favorite_teams[r] == ht:
            new_wk_sheet.cell(row=r + 2, column=3).fill = home_fill
        new_wk_sheet.cell(row=r + 2, column=4).value = spreads_int[r]
        new_wk_sheet.cell(row=r + 2, column=5).value = underdog_teams[r]
        if underdog_teams[r] == ht:
            new_wk_sheet.cell(row=r + 2, column=5).fill = home_fill
        new_wk_sheet.cell(row=r + 2, column=9).value = fav_abbr[r]
        new_wk_sheet.cell(row=r + 2, column=11).value = under_abbr[r]
    new_wk_sheet.cell(row=num_games-1, column=14).fill = blue_fill
    new_wk_sheet.cell(row=num_games-1, column=15).fill = blue_fill
    new_wk_sheet.cell(row=num_games, column=14).fill = blue_fill
    new_wk_sheet.cell(row=num_games, column=15).fill = blue_fill
    if wk_number == "Week 1":
        new_wk_sheet.cell(row=num_games-2, column=14).fill = blue_fill
        new_wk_sheet.cell(row=num_games-2, column=15).fill = blue_fill
        new_wk_sheet.cell(row=num_games-1, column=14).fill = blue_fill
        new_wk_sheet.cell(row=num_games-1, column=15).fill = blue_fill
        new_wk_sheet.cell(row=num_games, column=14).fill = blue_fill
        new_wk_sheet.cell(row=num_games, column=15).fill = blue_fill
    wb.save(file)
