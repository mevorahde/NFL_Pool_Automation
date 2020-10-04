import requests
from bs4 import BeautifulSoup as bs
import re

# Load the webpage content
url = "https://www.scoresandodds.com/las-vegas-odds"
r = requests.get(url)

# Convert to a beautiful soup object
webpage = bs(r.content, "html.parser")

# Get the table for Column names
table = webpage.select("table.segmented")[0]
columns = table.find("td", attrs={"id": "betting-trends--nfl"}).find_all("th", text=re.compile("Line"))
column_names = [c.string for c in columns]

print(column_names)


l = []
spread_rows = columns = table.find("td", attrs={"id": "betting-trends--nfl"}).find_all("span", text=re.compile("-"))
spreads = [str(s.get_text()).strip() for s in spread_rows]
print(spreads)


# for tr in table_rows:
#     td = tr.find_all("td")
#     row = [str(tr.get_text()).strip() for tr in td]
#     l.append(row)

# print(table.prettify())

