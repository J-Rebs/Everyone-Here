"""
Help Sources:

https://scriptsview.com/lng/pymailscp/ 
https://stackoverflow.com/questions/38899956/python-win32com-get-outlook-event-appointment-meeting-response-status
https://docs.microsoft.com/en-us/office/vba/outlook/how-to/search-and-filter/search-the-calendar-for-appointments-within-a-date-range-that-contain-a-specific
https://pythoninoffice.com/get-outlook-calendar-meeting-data-using-python/ 
"""
import win32com.client, win32timezone
import os

class Check():
    def __init__(self):
        self.outlook = win32com.client.Dispatch('outlook.application')
        self.mapi = self.outlook.GetNamespace("MAPI")
        self.calendar = self.mapi.GetDefaultFolder(9)


    # method substantially from: https://pythoninoffice.com/get-outlook-calendar-meeting-data-using-python/ 
    def get_meeting_info(self, begin, end):       
        appointments = self.calendar.Items
        appointments.IncludeRecurrences = True
        appointments.sort('[Start]')
        restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [END] <= '" + end.strftime('%m/%d/%Y') + "'"
        appointments = appointments.Restrict(restriction)
        
        for appointment in appointments:
            print('MEETING-->')
            print(appointment.subject)
            print(appointment.start)
            print('Recipients are:')
            recipients = appointment.Recipients
            for person in recipients:
                if person.MeetingResponseStatus == 0:
                    print(person, "has NOT accepted")
                else:
                    print(person, "has accepted")

