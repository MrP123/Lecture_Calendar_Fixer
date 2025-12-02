import os

import requests
import hashlib

from dotenv import load_dotenv

load_dotenv("api.env")

user = os.getenv("USER")
password = os.getenv("PASS")

session = requests.Session()

# --- Configuration ---
DEVICE_FINGERPRINT = hashlib.sha256(user.encode()).hexdigest()

BASE_URL = "https://callmyapi.mci4me.at"
LOGIN_URL = f"{BASE_URL}/api/my/4/auth/credentials?lang=de"
TERMINE_URL = f"{BASE_URL}/api/my/4/termine?lang=de"

# ---------------------------------------------------------
# Create a reusable session (same as PowerShell WebRequestSession)
# ---------------------------------------------------------
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/142.0.0.0 Safari/537.36",
})


# ---------------------------------------------------------
# 1) LOGIN REQUEST  (POST)
# ---------------------------------------------------------
login_headers = {
    "authority": "callmyapi.mci4me.at",
    "method": "POST",
    "path": "/api/my/4/auth/credentials?lang=de",
    "scheme": "https",
    "accept": "application/json",
    "accept-encoding": "gzip, deflate, br, zstd",
    "accept-language": "de-DE,de;q=0.8",
    "origin": "https://my.mci4me.at",
    "priority": "u=1, i",
    "referer": "https://my.mci4me.at/",
    "sec-ch-ua": '"Chromium";v="142", "Brave";v="142", "Not_A Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "sec-gpc": "1",
    "software-version": "2.0.0-beta.9",
    "x-device-fingerprint": DEVICE_FINGERPRINT,
    "x-platform": "desktop",
    "Content-Type": "application/json"
}

login_payload = {
    "username": user,
    "password": password,
    "deviceFingerprint": DEVICE_FINGERPRINT
}

login_response = session.post(
    LOGIN_URL,
    headers=login_headers,
    json=login_payload
)

login_response.raise_for_status()
data = login_response.json()
auth_token = data.get("token").get("auth_token")

print("Auth token:", auth_token)

# Headers for GET /termine
termine_headers = {
    "authority": "callmyapi.mci4me.at",
    "method": "GET",
    "path": "/api/my/4/termine?lang=de",
    "scheme": "https",
    "accept": "application/json",
    "accept-encoding": "gzip, deflate, br, zstd",
    "accept-language": "de-DE,de;q=0.8",
    "authorization": f"Bearer {auth_token}",
    "origin": "https://my.mci4me.at",
    "priority": "u=1, i",
    "referer": "https://my.mci4me.at/",
    "sec-ch-ua": '"Chromium";v="142", "Brave";v="142", "Not_A Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-site",
    "sec-gpc": "1",
    "software-version": "2.0.0-beta.9",
    "x-device-fingerprint": DEVICE_FINGERPRINT,
    "x-platform": "desktop",
}

session.headers.update(termine_headers)

appointment_response = session.get(
    TERMINE_URL,
    headers=termine_headers
)

appointment_response.raise_for_status()
appointments = appointment_response.json()

session.close()

for appointment in appointments:
    for key, value in appointment.items():
        print(f"{key}: {value}")

