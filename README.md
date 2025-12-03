# Lecture_Calendar_Fixer

## Setup

1. Install Python 3.13+
2. Rename the file `.env.template` in the root directory to `.env`.
3. Within the `.env ` file find the key `WEBCAL_URL` and insert the URL to your exported calendar from myMCI as a string. Follow the example in the `.env` file.
4. Create a virtual environment with `python -m venv .venv`
5. Activate the virtual environment
6. Run `pip install -r requirements.txt`
7. Modify the `run_script.bat` file to point to your virtual environment's python.exe and to the `lecture_calendar_fixer.py` file
8. Create the categories called "Vorlesung" and "Vorlesung-Anderer-Standort" in your outlook calendar with your desired colors
9. Setup a scheduled task to run the `run_script.bat` file, for example every morning

### Optional steps
1. Within `config.py` you can modify which MCI location is your default location. Those will be put in the category "Vorlesung", all others will be put in "Vorlesung-Anderer-Standort". The default location is "MCI IV".
2. Within `config.py` you can modify how much earlier your calendar appointment triggers a notification depending on the location of the lecture. Currently all six MCI locations are covered with a base notification time of 15 minutes before the lecture starts. The default is based around the "MCI IV" location.