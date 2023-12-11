import logging
import datetime
import os

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
            logging.info("\nTrying to delete appointment:\n\t" + EventWrapper.from_outlook_event(appointment))
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
        logging.info("\nAdding event:\n\t" + wrapper)
        wrapper.to_outlook_event(outlook)

def try_deleting_outlook_appointment(appointment):
    should_retry = 1
    tries = 0
    while should_retry > 0 and tries < 5:   
        tries += 1
        should_retry -= 1
        try:
            appointment.Delete()
        except:
            logging.warn("Could not delete appointment")
            should_retry += 1

def update_changed_events(webcalender, outlook):
    outlook_calendar = outlook.GetNamespace("MAPI").GetDefaultFolder(9) #calender = 9
    appointments = outlook_calendar.Items

    found_appointments = [appointment for appointment in appointments if EventWrapper.get_default_organizer() in getattr(appointment, "Organizer", "Error")]
    logging.info(F"Found {len(found_appointments)} outlook appointments")
    outlook_appointment_dict = {appointment.Organizer: appointment for appointment in found_appointments}

    all_events = [subcomp for subcomp in webcalendar.subcomponents if subcomp.name == "VEVENT"]
    lecture_events = [event for event in all_events if not "Abgabetermin" in event['summary']]
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
                logging.info("\nTrying to delete appointment:\n\t" + outlook_appointment_wrapped)
                try_deleting_outlook_appointment(corresponding_outlook_appointment)

                logging.info("\nAdding event:\n\t" + ical_event_wrapped)
                ical_event_wrapped.to_outlook_event(outlook)
            else:
                logging.info("\nAppointment is up to date:\n\t" + ical_event_wrapped)
        else:
            # if it is not available then add it
            logging.info("\nAdding event:\n\t" + ical_event_wrapped)
            ical_event_wrapped.to_outlook_event(outlook)

if __name__ == "__main__":
    load_dotenv()

    logging.basicConfig(filename='full.log', encoding='utf-8', level=logging.DEBUG)
    logging.basicConfig(filename='error.log', encoding='utf-8', level=logging.ERROR)

    # .env file with webcal link as http link must be available
    url = os.getenv("WEBCAL_URL")

    logging.info(F"Running at {datetime.datetime.now()}")
    try:
        response = requests.get(url)
    except:
        logging.error("Could not fetch calendar")
        exit(1)

    webcalendar = icalendar.Calendar.from_ical(response.text)
    outlook = win32com.client.Dispatch("Outlook.Application")

    #delete_all_existing_lecture_events(outlook)
    #add_lecture_events_to_outlook(webcalendar, outlook)

    update_changed_events(webcalendar, outlook)