# Lecture_Calendar_Fixer

## Setup

1. Install Python 3.11+
2. Create a `.env` file in the root directory that contains the key `WEBCAL_URL` that is a http link to your ical file
3. Create a virtual environment with `python -m venv .venv`
4. Activate the virtual environment
5. Run `pip install -r requirements.txt`
6. Modify the `run_script.bat` file to point to your virtual environment's python.exe and to the `lecture_calendar_fixer.py` file
7. Create a category called "Vorlesung" in your outlook calendar with your desired color
8. Setup a scheduled task to run the `run_script.bat` file for example every morning