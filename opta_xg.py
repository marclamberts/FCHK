import pandas as pd
import requests
import json
import os
from pandas import json_normalize
from datetime import datetime, timedelta
from bs4 import BeautifulSoup

########################################
######### CHANGE THE INFO HERE #########
########################################

lg = 'A-League Women 2021-2022'
lg_id = 'd6tuztyp7kqkmikytugafu3v8'  # League ID from Opta (found in the html of a league's page)

# In Colab, use /content/ paths (Google Drive mount optional)
# Example with Drive: '/content/drive/MyDrive/xG'
lg_schedule_location = "/content/xG"
match_download_location = "/content/xG"

# Create output dirs if they don't exist
os.makedirs(lg_schedule_location, exist_ok=True)
os.makedirs(f"{match_download_location}/{lg}", exist_ok=True)

##################################################################################
# TO SAVE TIME, ONLY FETCH GAMES AFTER last_run DATE
##################################################################################
d = datetime.today() - timedelta(days=2)

if lg == 'FBL':
    last_run = "2022-09-01"
elif lg == 'Indian Super League':
    last_run = "2023-02-16"
else:
    last_run = '2019-07-19'  # set to a far-back date to get everything

##########################################################################
# STEP 1: Fetch match schedule from Opta API and save as CSV
##########################################################################

url = (
    "https://api.performfeeds.com/soccerdata/match/qxcx5jmswgto1qeqcjzghtddt/"
    f"?_rt=c&tmcl={lg_id}&live=yes&_pgSz=400&_lcl=en&_fmt=jsonp"
    "&_clbk=W360b23205d13a5cdd46c315ae9c2097b829cbc909"
)

page = requests.get(url, headers={'referer': 'scoresway.com'})
soup = BeautifulSoup(page.content, "html.parser")

raw = soup.getText()
start = raw.find('{"match"') + 9
json_str = raw[start:-2]

# Parse and flatten JSON directly — no browser needed
match_data = json.loads(json_str)
df = json_normalize(match_data)
df.to_csv(f'{lg_schedule_location}/{lg} Matches.csv', index=False, encoding='utf-8-sig')
print('DONE WITH MATCHES FILE')

##########################################################################
# STEP 2: Filter matches and download per-match event data
##########################################################################

df = pd.read_csv(f'{lg_schedule_location}/{lg} Matches.csv')
df.dropna(subset=['liveData/matchDetails/scores/ft/home'], inplace=True)
df.reset_index(drop=True, inplace=True)
df = df[::-1].reset_index(drop=True)

cols = [
    'matchInfo/localDate', 'matchInfo/id',
    'matchInfo/contestant/0/officialName', 'liveData/matchDetails/scores/ft/home',
    'liveData/matchDetails/scores/ft/away', 'matchInfo/contestant/1/officialName',
    'matchInfo/contestant/0/id', 'matchInfo/contestant/1/id'
]
matches = df[cols].copy()

matches['matchName'] = matches.apply(
    lambda r: f"{r['matchInfo/localDate']}_{r['matchInfo/contestant/0/officialName']} - {r['matchInfo/contestant/1/officialName']}",
    axis=1
)

if last_run:
    matches['matchInfo/localDate'] = pd.to_datetime(matches['matchInfo/localDate'])
    matches = matches[matches['matchInfo/localDate'] > pd.to_datetime(last_run)].reset_index(drop=True)

for i, row in matches.iterrows():
    match_url = (
        "https://api.performfeeds.com/soccerdata/matchevent/qxcx5jmswgto1qeqcjzghtddt/"
        f"{row['matchInfo/id']}?_rt=c&_lcl=en&_fmt=jsonp"
        "&_clbk=W3bd804271ec7f49caa576ff8367109721760953b4"
    )
    page = requests.get(match_url, headers={'referer': 'scoresway.com'})
    soup = BeautifulSoup(page.content, "html.parser")

    raw = soup.getText()
    start = raw.find("liveData") + 10
    output = raw[start:-2]

    event_data = json.loads(output)

    safe_name = row['matchName'].replace('/', '')
    out_path = f"{match_download_location}/{lg}/{safe_name}.json"
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(event_data, f, ensure_ascii=True, indent=4)
    print(f"downloaded - {row['matchName']}")

print('DONE WITH EVERYTHING')
