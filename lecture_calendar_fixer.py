import logging
import datetime
import os
from pathlib import Path

from dotenv import load_dotenv
import win32com.client

import icalendar
import requests

from event import EventWrapper
from api_call import load_from_mymci_api
import config

def delete_all_existing_lecture_events(outlook):
    outlook_calendar = outlook.GetNamespace("MAPI").GetDefaultFolder(9)  # calendar = 9
    appointments = outlook_calendar.Items

    found_appointments = [
        appointment
        for appointment in appointments
        if EventWrapper.get_default_organizer() in getattr(appointment, "Organizer", "Error")
    ]
    logging.info(f"Found {len(found_appointments)} appointments")

    should_retry = 1
    tries = 0
    while should_retry > 0 and tries < 5:
        tries += 1
        should_retry -= 1

        for appointment in found_appointments:
            logging.info(f"\nTrying to delete appointment:\n\t{EventWrapper.from_outlook_event(appointment)}")
            try:
                appointment.Delete()
            except Exception as e:
                logging.warning(f"Could not delete appointment (Exception {e}), adding retry")
                should_retry += 1

def add_lecture_events_to_outlook(webcalendar, outlook):
    all_events = [subcomp for subcomp in webcalendar.subcomponents if subcomp.name == "VEVENT"]
    lecture_events = [event for event in all_events if "Abgabetermin" not in event["summary"]]

    logging.info(f"Found {len(lecture_events)} lecture events")

    for event in lecture_events:
        wrapper = EventWrapper.from_ical_event(event)
        logging.info(f"\nAdding event:\n\t{wrapper}")
        wrapper.to_outlook_event(outlook)


def try_deleting_outlook_appointment(appointment) -> bool:
    attempts = 0
    while attempts < 5:
        try:
            appointment.Delete()  # failure to delete will raise an exception
            return True
        except Exception as e:
            logging.warning(f"Could not delete appointment (Exception {e}), retrying...")
        attempts += 1
    return False

def webcal_dict_to_wrapper(webcalendar_dict: list[dict]) -> list[EventWrapper]:
    lecture_events = []
    for event in webcalendar_dict:
        # skip submission dates
        if event["art"] not in ["Lehrveranstaltung", "Prüfung", "Sonstiges"]:
            continue

        lecture_events.append(EventWrapper.from_api_dict(event))

    return lecture_events

def webcal_to_wrapper(webcalendar) -> list[EventWrapper]:
    all_events = [subcomp for subcomp in webcalendar.subcomponents if subcomp.name == "VEVENT"]
    lecture_events = []
    for event in all_events:
        # skip submission dates
        if "Abgabetermin" in event["summary"]:
            continue

        # skip all events that are stemming from your own SAKAI calendar
        # UID of event is either "MCI-DESIGNER-TERMIN-xxxx" or "MCI-SAKAI-TERMIN-xxxx"
        if "MCI-SAKAI-TERMIN" in event["uid"]:
            continue

        lecture_events.append(EventWrapper.from_ical_event(event))

    return lecture_events

def update_changed_events(wrapped_events: list[EventWrapper], outlook):
    outlook_calendar = outlook.GetNamespace("MAPI").GetDefaultFolder(9) #calender = 9
    appointments = outlook_calendar.Items

    found_appointments = [appointment for appointment in appointments if EventWrapper.get_default_organizer() in getattr(appointment, "Organizer", "Error")]
    logging.info(f"Found {len(found_appointments)} outlook appointments")
    outlook_appointment_dict = {appointment.Organizer: appointment for appointment in found_appointments}

    # create a dict of all found/imported events
    lecture_event_dict: dict[str, EventWrapper] = {}
    for event in wrapped_events:
        lecture_event_dict[event.organizer] = event

    logging.info(f"Found {len(wrapped_events)} lecture ical events")

    for imported_event in lecture_event_dict.values():
        # get corresponding outlook appointment from dict
        corresponding_outlook_appointment = outlook_appointment_dict.get(imported_event.organizer)

        if corresponding_outlook_appointment:
            # if found wrap it for comparison
            outlook_appointment_wrapped = EventWrapper.from_outlook_event(corresponding_outlook_appointment)
            
            if imported_event != outlook_appointment_wrapped:
                # event has changed --> delete and add again
                logging.info(f"\nTrying to delete appointment:\n\t{outlook_appointment_wrapped}")
                if try_deleting_outlook_appointment(corresponding_outlook_appointment):
                    #also remove from dict
                    del outlook_appointment_dict[imported_event.organizer]

                logging.info(f"\nAdding event:\n\t{imported_event}")
                imported_event.to_outlook_event(outlook)
            else:
                logging.info(f"\nAppointment is up to date:\n\t{imported_event}")
        else:
            # if it is not available then add it
            logging.info(f"\nAdding event:\n\t{imported_event}")
            imported_event.to_outlook_event(outlook)

    # More outlook appointments than ical events --> something has been deleted in the ical events
    if len(outlook_appointment_dict) > len(lecture_event_dict):
        for outlook_appointment in outlook_appointment_dict.values():
            if not lecture_event_dict.get(outlook_appointment.Organizer):
                outlook_appointment_wrapped = EventWrapper.from_outlook_event(outlook_appointment)
                
                #if the outlook appointment is not in the ical events then delete it only if it is in the future
                if outlook_appointment_wrapped.start_dt > datetime.datetime.now(outlook_appointment_wrapped.start_dt.tzinfo):
                    logging.info(f"\nTrying to delete appointment:\n\t{outlook_appointment_wrapped}")
                    if try_deleting_outlook_appointment(outlook_appointment):
                        del outlook_appointment_dict[imported_event.organizer]

if __name__ == "__main__":
    #Dynamically load the type info for the underlying COM object. This data can be generated with the following command:
    #python .\.venv\Lib\site-packages\win32com\client\makepy.py -i "Microsoft Outlook 16.0 Object Library"
    #This allows for accuarte type information when debugging
    #try:
    #    win32com.client.gencache.EnsureModule('{00062FFF-0000-0000-C000-000000000046}', 0, 9, 6)
    #except Exception as e:
    #    logging.debug(F"Could not ensure module for type data on COM object: {e}")

    load_dotenv()
    
    logfile_path = Path(__file__).parent.resolve() / "full.log" #always logs into the same folder as the script, even when run from task scheduler
    logging.basicConfig(filename=logfile_path, encoding='utf-8', level=logging.DEBUG)
    logging.info(f"Running at {datetime.datetime.now()}")

    if config.use_ical_link() and config.use_api_call():
        logging.error("Both use_ical_link and use_api_call are set to True in config.py. Please choose only one method to fetch events.")
        exit(1)

    if not (config.use_ical_link() or config.use_api_call()):
        logging.error("Both use_ical_link and use_api_call are set to False in config.py. Please choose one method to fetch events.")
        exit(1)

    if config.use_ical_link():
        url = os.getenv("WEBCAL_URL")
        if url is None:
            logging.error("No webcal url found in .env file")
            exit(1)
        
        try:
            response = requests.get(url)
        except requests.exceptions.RequestException as e:
            logging.error(F"Could not fetch calendar: {e}")
            exit(1)   
        
        webcalendar = icalendar.Calendar.from_ical(response.text)
        wrapped_events = webcal_to_wrapper(webcalendar)

    elif config.use_api_call():
        user = os.getenv("USER")
        if user is None:
            logging.error("No user found in .env file")
            exit(1)

        webcalendar = load_from_mymci_api(user)
        wrapped_events = webcal_dict_to_wrapper(webcalendar)


    #delete_all_existing_lecture_events(outlook)
    #add_lecture_events_to_outlook(webcalendar, outlook)

    outlook = win32com.client.Dispatch("Outlook.Application")
    update_changed_events(wrapped_events, outlook)