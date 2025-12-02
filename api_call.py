import os

import requests
from dotenv import load_dotenv

load_dotenv("api.env")

user = os.getenv("USER")
password = os.getenv("PASS")

session = requests.Session()

# Set User-Agent
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/140.0.0.0 Safari/537.36"
})

# Define headers (excluding "method", "path", "scheme" since requests handles that)
headers = {
    "authority": "my.mci4me.at",
    "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,"
              "image/webp,image/apng,*/*;q=0.8",
    "accept-encoding": "gzip, deflate, br, zstd",
    "accept-language": "de-DE,de;q=0.8",
    "cache-control": "max-age=0",
    "origin": "https://my.mci4me.at",
    "priority": "u=0, i",
    "referer": "https://my.mci4me.at/",
    "sec-ch-ua": '"Chromium";v="140", "Not=A?Brand";v="24", "Brave";v="140"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "document",
    "sec-fetch-mode": "navigate",
    "sec-fetch-site": "same-origin",
    "sec-fetch-user": "?1",
    "sec-gpc": "1",
    "upgrade-insecure-requests": "1"
}

# Form data
data = {
    "request": "login",
    "username": user,
    "password": password,
    "submit": "Login"
}

# Send POST request
login_url = "https://my.mci4me.at/"
login_response = session.post(login_url, headers=headers, data=data)

print(f"Login response status: {login_response.status_code}")
print(f"Login response text: {login_response.text}")


timetable_url = "https://callmyapi.mci4me.at/api/my/4/termine?lang=de"
json_headers = {
    "authority": "callmyapi.mci4me.at",
    "accept": "application/json, text/javascript, */*; q=0.01",
    "accept-encoding": "gzip, deflate, br, zstd",
    "accept-language": "de-DE,de;q=0.8",
    "priority": "u=1, i",
    "referer": "https://my.mci4me.at/",
    "sec-ch-ua": '"Chromium";v="140", "Not=A?Brand";v="24", "Brave";v="140"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    "sec-gpc": "1",
    "x-requested-with": "XMLHttpRequest"
}

# GET request for JSON data
timetable_response = session.get(timetable_url, headers=json_headers)

print(f"Timetable response status: {timetable_response.status_code}")

# If response is JSON, parse directly
try:
    timetable_json = timetable_response.json()
    timetable_list = timetable_json.get("aaData", []) # for some reason it's called aaData

    for entry in timetable_list:
        print(entry)

        # unix time stamp of start in my time zone
        # Start date with format "DAY<br/>DD.MM.YYYY"
        # Time start/end with format "HH:MM<br/>HH:MM"
        # Title
        # Groups
        # Lecturer(s)
        # currently empty --> Room?
        # duration in teaching units
        # Internal name code for lecture: e.g. MECH-B-3-SWD-SWD-ILV
        # currently empty --> ?
        # currently empty --> ?

except ValueError:
    print(f"Response is not JSON. Raw text: {timetable_response.text[:500]}")


session.close()