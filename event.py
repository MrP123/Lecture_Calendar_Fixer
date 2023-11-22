from icalendar.cal import Event as ICalEvent
from win32com.client.dynamic import CDispatch as Outlook
from datetime import datetime

class EventWrapper():

    def __init__(self, subject: str, start: str, duration: int, location: str, organizer: str | None = None, start_dt: datetime | None = None) -> None:
        self.subject = subject
        self.start = start
        self.duration = duration
        self.location = location
        self.start_dt = start_dt
        self.organizer = organizer if organizer else "MCI-DESIGNER-TERMIN"

    @staticmethod
    def from_ical_event(event: ICalEvent):
        subject = event['summary']
        start_dt = event['dtstart'].dt
        end_dt = event['dtend'].dt
        dur = int((end_dt - start_dt).total_seconds() / 60)
        start = start_dt.strftime("%Y-%m-%d %H:%M")
        location = event.get('location')
        location = location if location else "-"
        return EventWrapper(subject, start, dur, location, start_dt=start_dt)
    
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

        appt.ReminderSet = not is_past
        appt.ReminderMinutesBeforeStart = 15
        appt.BusyStatus = 2 # 2 = olBusy

        appt.Organizer = self.organizer
        appt.Save()
        appt.Send()
        return appt

    @staticmethod
    def get_default_organizer() -> str:
        return "MCI-DESIGNER-TERMIN"

    def __str__(self) -> str:
        return F"Subject: {self.subject}\n\tStart: {self.start}\n\tDuration: {self.duration}\n\tLocation: {self.location}\n\tOrganizer: {self.organizer}\n\tStart_dt: {self.start_dt}"
    
    # only define left and right addition for string representation of EventWrapper
    def __add__(self, other: str) -> str:
        return str(self) + other
    
    def __radd__(self, other: str) -> str:
        return other + str(self)