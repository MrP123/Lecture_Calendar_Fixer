import logging
import datetime
import os
import hashlib

from dotenv import load_dotenv
import icalendar
import requests
import win32com.client

from event import EventWrapper

def delete_all_existing_lecture_events(outlook):
    outlook_calendar = outlook.GetNamespace("MAPI").GetDefaultFolder(9) #calender = 9
    appointments = outlook_calendar.Items

    found_appointments = [appointment for appointment in appointments if EventWrapper.get_default_organizer() in getattr(appointment, "Organizer", "Error")]
    logging.info(F"Found {len(found_appointments)} appointments")

    should_retry = 1
    tries = 0
    while should_retry > 0 and tries < 5:   
        tries += 1
        should_retry -= 1

        for appointment in found_appointments:
            logging.info(F"\nTrying to delete appointment:\n\t{EventWrapper.from_outlook_event(appointment)}")
            try:
                appointment.Delete()
            except:
                logging.warn("Could not delete appointment, adding retry")
                should_retry += 1

def add_lecture_events_to_outlook(webcalendar, outlook):
    all_events = [subcomp for subcomp in webcalendar.subcomponents if subcomp.name == "VEVENT"]
    lecture_events = [event for event in all_events if not "Abgabetermin" in event['summary']]

    logging.info(F"Found {len(lecture_events)} lecture events")

    for event in lecture_events:
        wrapper = EventWrapper.from_ical_event(event)
        logging.info(F"\nAdding event:\n\t{wrapper}")
        wrapper.to_outlook_event(outlook)

def try_deleting_outlook_appointment(appointment) -> bool:
    attempts = 0
    while attempts < 5:   
        try:
            appointment.Delete() #failure to delete will raise an exception
            return True
        except:
            logging.warn("Could not delete appointment")
        attempts += 1
    return False

def update_changed_events(webcalender, outlook):
    outlook_calendar = outlook.GetNamespace("MAPI").GetDefaultFolder(9) #calender = 9
    appointments = outlook_calendar.Items

    found_appointments = [appointment for appointment in appointments if EventWrapper.get_default_organizer() in getattr(appointment, "Organizer", "Error")]
    logging.info(F"Found {len(found_appointments)} outlook appointments")
    outlook_appointment_dict = {appointment.Organizer: appointment for appointment in found_appointments}

    all_events = [subcomp for subcomp in webcalendar.subcomponents if subcomp.name == "VEVENT"]
    lecture_events = [event for event in all_events if not "Abgabetermin" in event['summary']]

    lecture_event_dict = {}
    for event in lecture_events:
        ical_event_wrapped = EventWrapper.from_ical_event(event)
        lecture_event_dict[ical_event_wrapped.organizer] = ical_event_wrapped

    logging.info(F"Found {len(lecture_events)} ical events")

    for event in lecture_events:
        ical_event_wrapped = EventWrapper.from_ical_event(event)

        # get corresponding outlook appointment from dict
        corresponding_outlook_appointment = outlook_appointment_dict.get(ical_event_wrapped.organizer)
        if corresponding_outlook_appointment:

            # if found wrap it for comparison
            outlook_appointment_wrapped = EventWrapper.from_outlook_event(corresponding_outlook_appointment)
            if ical_event_wrapped != outlook_appointment_wrapped:
                # event has changed --> delete and add again
                logging.info(F"\nTrying to delete appointment:\n\t{outlook_appointment_wrapped}")
                try_deleting_outlook_appointment(corresponding_outlook_appointment)

                logging.info(F"\nAdding event:\n\t{ical_event_wrapped}")
                ical_event_wrapped.to_outlook_event(outlook)
            else:
                logging.info(F"\nAppointment is up to date:\n\t{ical_event_wrapped}")
        else:
            # if it is not available then add it
            logging.info(F"\nAdding event:\n\t{ical_event_wrapped}")
            ical_event_wrapped.to_outlook_event(outlook)

    if len(found_appointments) > len(lecture_events):
        for outlook_appointment in found_appointments:
            if not lecture_event_dict.get(outlook_appointment.Organizer):
                outlook_appointment_wrapped = EventWrapper.from_outlook_event(outlook_appointment)

                #if the outlook appointment is not in the ical events then delete it only if it is in the future
                if outlook_appointment_wrapped.start_dt > datetime.datetime.now(outlook_appointment_wrapped.start_dt.tzinfo):
                    logging.info(F"\nTrying to delete appointment:\n\t{outlook_appointment_wrapped}")
                    try_deleting_outlook_appointment(outlook_appointment)

def update_hash_file(webcalendar):
    # turn all events into a string and hash it
    all_events = [subcomp for subcomp in webcalendar.subcomponents if subcomp.name == "VEVENT"]
    full_string = "".join([str(EventWrapper.from_ical_event(event)) for event in all_events if not "Abgabetermin" in event['summary']])

    with open("./.hash", "a+") as f:
        f.seek(0)
        old_hash = f.read()
        new_hash = hashlib.sha256(str(full_string).encode('utf-8')).hexdigest()
        if old_hash == new_hash:
            logging.info("No changes in hash --> will continue anyway for now")
        else:
            logging.info("Hash changed --> will update")
            f.truncate(0)
            f.write(new_hash)

if __name__ == "__main__":
    #Dynamically load the type info for the underlying COM object. This data can be generated with the following command:
    #python .\.venv\Lib\site-packages\win32com\client\makepy.py -i "Microsoft Outlook 16.0 Object Library"
    #This allows for accuarte type information when debugging
    try:
        win32com.client.gencache.EnsureModule('{00062FFF-0000-0000-C000-000000000046}', 0, 9, 6)
    except Exception as e:
        logging.debug(F"Could not ensure module for type data on COM object: {e}")

    load_dotenv()
    logging.basicConfig(filename='full.log', encoding='utf-8', level=logging.DEBUG)

    # .env file with webcal link as http link must be available
    url = os.getenv("WEBCAL_URL")
    if url is None:
        logging.error("No webcal url found in .env file")
        exit(1)

    logging.info(F"Running at {datetime.datetime.now()}")
    try:
        response = requests.get(url)
    except requests.exceptions.RequestException as e:
        logging.error(F"Could not fetch calendar: {e}")
        exit(1)   

    webcalendar = icalendar.Calendar.from_ical(response.text)
    update_hash_file(webcalendar)
    
    outlook = win32com.client.Dispatch("Outlook.Application")

    #delete_all_existing_lecture_events(outlook)
    #add_lecture_events_to_outlook(webcalendar, outlook)

    update_changed_events(webcalendar, outlook)