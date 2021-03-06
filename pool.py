import datetime
import logging
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

# Logging configurations
logging.basicConfig(filename='Pool Automation Error Log.log',
                    level=logging.CRITICAL,
                    format='%(asctime)s : %(levelname)s : %(message)s')

# define a Handler which writes INFO messages or higher to the sys.stderr
console = logging.StreamHandler()
console.setLevel(logging.INFO)
# add the handler to the root logger
logging.getLogger('').addHandler(console)

try:
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
            date_time_value = table.find("span", attrs={"data-value": True}).attrs
            date_time_UTC = date_time_value['data-value']
            row = [tm1_name.upper(), spread, tm2_name.upper(), tm1_abbr, tm2_abbr, home_tm, date_time_UTC]
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
            date_time_value = table.find("span", attrs={"data-value": True}).attrs
            date_time_UTC = date_time_value['data-value']
            row = [tm1_name.upper(), final_away_spread, tm2_name.upper(), tm1_abbr, tm2_abbr, home_tm, date_time_UTC]
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
    date_time = [date_and_time[6] for date_and_time in data]

    dotw = []
    game_time = []
    for i in date_time:
        utc_time_remove_t = i.replace("T", " ")
        split_utc_date_from_time = utc_time_remove_t.split(" ")
        dt = split_utc_date_from_time[0]
        year, month, day = (int(x) for x in dt.split('-'))
        dow = datetime.date(year, month, day).weekday()
        dotw.append(dow)
        time = split_utc_date_from_time[1]
        game_time.append(time)

    current_day = datetime.date.today().weekday()
    game_day_of_week = [day for day in dotw]
    num_games_a_thur = [num_games_a_t for num_games_a_t in dotw if num_games_a_t != 4]
    num_games_a_sat = [num_games_a_s for num_games_a_s in dotw if num_games_a_s != 4 or num_games_a_s != 5]
    dotw_adjusted = [5 if day == 6 and time == '00:15:00Z' else day for day, time in zip(dotw, game_time)]
    num_of_thur_games_list = []
    num_pre_sun_games_list = []
    for g in dotw_adjusted:
        if g != 6 and g != 0 and g != 1:
            num_pre_sun_games_list.append(g)
        if g == 4:
            num_of_thur_games_list.append(g)
    num_of_thur_games = len(num_of_thur_games_list)
    num_pre_sun_games = len(num_pre_sun_games_list)
    num_games_after_thur = len(num_games_a_thur)
    num_games_after_sat = len(num_games_a_sat)
    difference_num_of_games_after_thur = num_games - num_games_after_thur
    difference_num_of_games_after_sat = num_games - num_games_after_sat

    # 00:20:00Z PST/UTC - Sunday Night
    # 00:15:00Z PST/UTC - Monday Night
    # Sunday Night games will show as '0' as the day of the week.
    # Monday Night games will show as a '1' as the day of the week.
    # For any other Tues night games (mainly in 2020 for COVID), will show as a '2' as the day of the week.

    # Excel Info and processes
    file = os.getenv("file_path")
    wb = load_workbook(filename=file)
    all_sheets = wb.sheetnames
    template = wb.worksheets[0]
    blue_fill = PatternFill(start_color='00B0F0', end_color='00B0F0', fill_type='solid')
    home_fill = PatternFill(start_color='F4B084', end_color='F4B084', fill_type='solid')
    clear_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')

    if wk_number not in all_sheets:
        template_copy = wb.copy_worksheet(template)
        new_wk_sheet = wb['Template Copy']
        new_wk_sheet.title = wk_number

        for sheet in wb:
            wb[sheet.title].views.sheetView[0].tabSelected = False

        wb.active = new_wk_sheet

        for r in range(0, num_games):
            ht = home_team[r]
            day_of_the_week = dotw[r]
            new_wk_sheet.cell(row=r + 2, column=3).value = favorite_teams[r]
            if favorite_teams[r] == ht:
                new_wk_sheet.cell(row=r + 2, column=3).fill = home_fill
            new_wk_sheet.cell(row=r + 2, column=4).value = spreads_int[r]
            new_wk_sheet.cell(row=r + 2, column=5).value = underdog_teams[r]
            if underdog_teams[r] == ht:
                new_wk_sheet.cell(row=r + 2, column=5).fill = home_fill
            new_wk_sheet.cell(row=r + 2, column=9).value = fav_abbr[r]
            new_wk_sheet.cell(row=r + 2, column=11).value = under_abbr[r]
            if day_of_the_week < 3:
                new_wk_sheet.cell(row=r + 2, column=14).fill = blue_fill
                new_wk_sheet.cell(row=r + 2, column=15).fill = blue_fill
                new_wk_sheet.cell(row=r + 2, column=14).fill = blue_fill
                new_wk_sheet.cell(row=r + 2, column=15).fill = blue_fill
        wb.save(file)
    else:
        wk_sheet = wb[wk_number]
        for sheet in wb:
            wb[sheet.title].views.sheetView[0].tabSelected = False

        wb.active = wk_sheet
        cells = wk_sheet['C3': 'E{}'.format(num_games + num_of_thur_games)]
        if current_day == 4:
            # Compare Website data to existing spreadsheet data
            website_values = []
            cell_values = []
            final_values = []

            for r in range(0, num_games_after_thur + num_of_thur_games):
                game_values = [favorite_teams[r], spreads_int[r], underdog_teams[r], dotw_adjusted[r]]
                website_values.append(game_values)
            for c1, c2, c3 in cells:
                game_values = [c1.value, float(c2.value), c3.value]
                cell_values.append(game_values)

            for web, cell in zip(website_values, cell_values):
                if web != cell:
                    final_values.append(web)
                else:
                    final_values.append(cell)

            day_of_week = [day[3] for day in final_values]
            favorite_teams_final = [fav_teams[0] for fav_teams in final_values]
            spreads_string = [spreads[1] for spreads in final_values]
            spreads_int = []
            for i in range(0, len(spreads_string)):
                if spreads_string[i] == '':
                    spreads_string[i] = 0
                spreads_string[i] = spreads_int.append(float(spreads_string[i]))
            underdog_teams_final = [under_teams[2] for under_teams in final_values]

            # KEEP NEXT 3 LINES FOR TESTINGS PURPOSES
            # print("Web:", website_values)
            # print("Cell:", cell_values)
            # print("Final:", final_values)

            '''
            replace any data changes for Favored Team (Column 'C'), the Spread (Column 'D'), and the Underdog Team 
            (Column 'E')
            - Ignores any Thurs games as this point as it is running on Fri updating for any Sat and Sun/Mon games.
            '''
            for r in range(difference_num_of_games_after_thur, num_games_after_thur):
                day = day_of_week[r]
                ht = home_team[r]
                if day != 4:
                    wk_sheet.cell(row=r + (2 + num_of_thur_games), column=3).value = favorite_teams_final[r]
                    wk_sheet.cell(row=r + (2 + num_of_thur_games), column=3).fill = clear_fill
                    if favorite_teams[r] == ht:
                        wk_sheet.cell(row=r + (2 + num_of_thur_games), column=3).fill = home_fill
                    wk_sheet.cell(row=r + (2 + num_of_thur_games), column=4).value = spreads_int[r]
                    wk_sheet.cell(row=r + (2 + num_of_thur_games), column=5).value = underdog_teams_final[r]
                    wk_sheet.cell(row=r + (2 + num_of_thur_games), column=5).fill = clear_fill
                    if underdog_teams[r] == ht:
                        wk_sheet.cell(row=r + (2 + num_of_thur_games), column=5).fill = home_fill
            wb.save(file)
        if current_day == 5:
            # Compare Website data to existing spreadsheet data
            website_values = []
            cell_values = []
            final_values = []

            for r in range(0, num_games_after_sat - num_pre_sun_games):
                game_values = [favorite_teams[r], spreads_int[r], underdog_teams[r], dotw_adjusted[r]]
                website_values.append(game_values)
            for c1, c2, c3 in cells:
                game_values = [c1.value, float(c2.value), c3.value]
                cell_values.append(game_values)

            for web, cell in zip(website_values, cell_values):
                if web != cell:
                    final_values.append(web)
                else:
                    final_values.append(cell)

            day_of_week = [day[3] for day in final_values]
            favorite_teams_final = [fav_teams[0] for fav_teams in final_values]
            spreads_string = [spreads[1] for spreads in final_values]
            spreads_int = []
            for i in range(0, len(spreads_string)):
                if spreads_string[i] == '':
                    spreads_string[i] = 0
                spreads_string[i] = spreads_int.append(float(spreads_string[i]))
            underdog_teams_final = [under_teams[2] for under_teams in final_values]

            # KEEP NEXT 3 LINES FOR TESTINGS PURPOSES
            # print("Web:", website_values)
            # print("Cell:", cell_values)
            # print("Final:", final_values)

            '''
            replace any data changes for Favored Team (Column 'C'), the Spread (Column 'D'), and the Underdog Team 
            (Column 'E')
            - Ignores any Thurs and Sat games as at this point is running on Sat for updates on Sun/Mon games.
            '''
            for r in range(difference_num_of_games_after_sat, num_games_after_sat - num_pre_sun_games):
                day = day_of_week[r]
                ht = home_team[r]
                if day == 6 or day == 0 or day == 1:
                    wk_sheet.cell(row=r + (2 + num_pre_sun_games), column=3).value = favorite_teams_final[r]
                    wk_sheet.cell(row=r + (2 + num_pre_sun_games), column=3).fill = clear_fill
                    if favorite_teams[r] == ht:
                        wk_sheet.cell(row=r + (2 + num_pre_sun_games), column=3).fill = home_fill
                    wk_sheet.cell(row=r + (2 + num_pre_sun_games), column=4).value = spreads_int[r]
                    wk_sheet.cell(row=r + (2 + num_pre_sun_games), column=5).value = underdog_teams_final[r]
                    wk_sheet.cell(row=r + (2 + num_pre_sun_games), column=5).fill = clear_fill
                    if underdog_teams[r] == ht:
                        wk_sheet.cell(row=r + (2 + num_pre_sun_games), column=5).fill = home_fill
            wb.save(file)
except Exception as e:
    logging.critical(e, exc_info=True)
