# Lecture_Calendar_Fixer

## Setup

1. Install Python 3.13+
2. Rename the file `.env.template` in the root directory to `.env`.
4. Create a virtual environment with `python -m venv .venv`
5. Activate the virtual environment
6. Run `pip install -r requirements.txt`
7. Within the `.env ` file find various configuration keys:
    - The key `WEBCAL_URL` is used if the `ical` link shall be used as source for your calendar appointments. Insert the URL to your exported calendar from myMCI as a string. Follow the example in the `.env` file.
    - The key `USER` is used if the myMCI API shall be used as source for your calendar appointments. Insert your myMCI username (i.e. your email) as string. This method requires your password to be stored in the `Python` keyring of your operating system with the `system`-string of `lecture_calendar_fixer`. To set this initially run `python -m keyring set lecture_calendar_fixer <my_username>` in your terminal and enter your myMCI password when prompted.
8. Within `config.py` make one of the functions `use_ical_link()` or `use_api_call()` return `True` depending on which source you want to use for fetching your calendar appointments. Only one of them can return `True`.
7. Modify the `run_script.bat` file to point to your virtual environment's python.exe and to the `lecture_calendar_fixer.py` file
8. Create the categories called "Vorlesung" and "Vorlesung-Anderer-Standort" in your outlook calendar with your desired colors
9. Setup a scheduled task to run the `run_script.bat` file, for example every morning

### Optional steps
1. Within `config.py` you can modify which MCI location is your default location. Those will be put in the category "Vorlesung", all others will be put in "Vorlesung-Anderer-Standort". The default location is "MCI IV".
2. Within `config.py` you can modify how much earlier your calendar appointment triggers a notification depending on the location of the lecture. Currently all six MCI locations are covered with a base notification time of 15 minutes before the lecture starts. The default is based around the "MCI IV" location.