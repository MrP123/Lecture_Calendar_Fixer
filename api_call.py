import logging

import requests
import hashlib

import keyring

def load_from_mymci_api(user: str, keyring_system: str = "lecture_calendar_fixer") -> list[dict]:
    password = keyring.get_password(keyring_system, user)

    if password is None:
        logging.error(F"No password found in system keyring for user: {user}. Please set it beforehand")
        logging.error("To do so, run: `python -m keyring set lecture_calendar_fixer <username>`")
        exit(1)

    session = requests.Session()
    DEVICE_FINGERPRINT = hashlib.sha256(user.encode()).hexdigest() #arbitrary design choice

    BASE_URL = "https://callmyapi.mci4me.at"
    LOGIN_URL = f"{BASE_URL}/api/my/4/auth/credentials?lang=de"
    TERMINE_URL = f"{BASE_URL}/api/my/4/termine?lang=de"

    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/142.0.0.0 Safari/537.36",
    })


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

    try:
        login_response.raise_for_status()
    except requests.exceptions.RequestException as e:
        logging.error(F"Could not authenticate with myMCI API: {e}")
        exit(1)

    data = login_response.json()
    auth_token = data.get("token").get("auth_token")

    if auth_token is None:
        logging.error("Could not retrieve auth token from login response")
        exit(1)

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

    try:
        appointment_response.raise_for_status()
    except requests.exceptions.RequestException as e:
        logging.error(F"Could not fetch calendar: {e}")
        exit(1)

    appointments = appointment_response.json()

    session.close()

    return appointments