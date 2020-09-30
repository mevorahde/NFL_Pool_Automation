import requests
from bs4 import BeautifulSoup as bs
import re


# Load the webpage content
url = "https://www.scoresandodds.com/las-vegas-odds"
r = requests.get(url)

# Convert to a beautiful soup object
webpage = bs(r.content, "html.parser")

# Get the table of
table = webpage.select("table.segmented tbody")[0]
columns = table.find("thread").find("th", string=re.compile("l|Line"))
column_names = [c.string for c in columns]

print(columns)