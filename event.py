from icalendar.cal import Event as ICalEvent
from win32com.client.dynamic import CDispatch as Outlook
import pywintypes

from datetime import datetime

from config import get_travel_time

class EventWrapper():

    def __init__(self, subject: str, start: str, duration: int, location: str, organizer: str | None = None, start_dt: datetime | None = None) -> None:
        self.subject = subject
        self.start = start
        self.duration = duration
        self.location = location
        self.start_dt = start_dt
        self.organizer = organizer if organizer else EventWrapper.get_default_organizer()

    @staticmethod
    def from_ical_event(event: ICalEvent):
        subject = event['summary']
        start_dt = event['dtstart'].dt
        end_dt = event['dtend'].dt
        dur = int((end_dt - start_dt).total_seconds() / 60)
        start = start_dt.strftime("%Y-%m-%d %H:%M")
        location = event.get('location')
        location = location if location else "-"

        organizer = event.get('UID')
        organizer = str(organizer) if organizer else EventWrapper.get_default_organizer()

        return EventWrapper(subject, start, dur, location, organizer=organizer, start_dt=start_dt)
    
    @staticmethod
    def from_outlook_event(event):
        return EventWrapper(event.Subject, event.Start, event.Duration, event.Location, event.Organizer)

    def to_outlook_event(self, outlook: Outlook):
        appt = outlook.CreateItem(1) # AppointmentItem
        appt.Start = self.start # yyyy-MM-dd hh:mm
        appt.Subject = self.subject
        appt.Duration = self.duration # In minutes (60 Minutes)
        appt.Location = self.location

        #by default handle events as if they were past events
        is_past = True
        if self.start_dt:
            is_past = self.start_dt < datetime.now(self.start_dt.tzinfo)

        reminder_time = 15
        if self.location != "-":
            # check if location is in mci_travel_times
            mci_location = self.location.split("/ ")[1]
            reminder_time += get_travel_time(mci_location)

        appt.ReminderSet = not is_past
        appt.ReminderMinutesBeforeStart = reminder_time
        appt.BusyStatus = 2 # 2 = olBusy
        appt.Categories = "Vorlesung"

        appt.Organizer = self.organizer
        appt.Save()
        appt.Send()
        return appt

    @staticmethod
    def get_default_organizer() -> str:
        return "MCI-DESIGNER-TERMIN"    

    def __eq__(self, __value: object) -> bool:
        same_subject = self.subject == __value.subject

        # If the EventWrapper is coming from an icalendar event then the start is a string, but it also always has the start_dt attribute
        # If the EventWrapper is coming from an outlook event then the start is a pywintype.datetime object
        # ToDo: Remove this distinction between the two ways a EventWrapper can be created
        self_start = self.start
        if isinstance(self_start, str):
            self_start = self.start_dt

        other_start = __value.start
        if isinstance(other_start, str):
            other_start = __value.start_dt

        same_start = self_start.ctime() == other_start.ctime()
        same_duration = self.duration == __value.duration
        same_location = self.location == __value.location
        same_organizer = self.organizer == __value.organizer
        return same_subject and same_start and same_duration and same_location and same_organizer

    def __str__(self) -> str:
        return F"Subject: {self.subject}\n\tStart: {self.start}\n\tDuration: {self.duration}\n\tLocation: {self.location}\n\tOrganizer: {self.organizer}\n\tStart_dt: {self.start_dt}"
    
    # only define left and right addition for string representation of EventWrapper
    def __add__(self, other: str) -> str:
        return str(self) + other
    
    def __radd__(self, other: str) -> str:
        return other + str(self)