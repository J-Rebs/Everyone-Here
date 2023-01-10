"""
Help Sources:

https://scriptsview.com/lng/pymailscp/ 
https://stackoverflow.com/questions/38899956/python-win32com-get-outlook-event-appointment-meeting-response-status
https://docs.microsoft.com/en-us/office/vba/outlook/how-to/search-and-filter/search-the-calendar-for-appointments-within-a-date-range-that-contain-a-specific
https://pythoninoffice.com/get-outlook-calendar-meeting-data-using-python/ 
https://docs.microsoft.com/en-us/office/client-developer/outlook/pia/how-to-check-all-responses-to-a-meeting-request 
https://docs.microsoft.com/en-us/office/vba/api/outlook.olresponsestatus
https://learn.microsoft.com/en-us/office/vba/api/outlook.appointmentitem.getorganizer
https://tdalon.blogspot.com/2020/09/outlook-vba-get-from-email.html
"""
import win32com.client, win32timezone
import os


class Check():
    def __init__(self):
        self.outlook = win32com.client.Dispatch('outlook.application')
        self.mapi = self.outlook.GetNamespace("MAPI")
        self.userID = self.mapi.CurrentUser.EntryID
        self.calendar = self.mapi.GetDefaultFolder(9)
        self.responseDict = {3: "Meeting accepted",
                             4: "Meeting declined",
                             0: "Unclear response -- no response, may be a simple appointment, may be person is "
                                "registered with "
                                "Outlook under Accounts, or may be something else",
                             5: "Recipient has not responded",
                             1: "The AppointmentItem is on the Organizer's calendar or the recipient is the Organizer "
                                "of the meeting",
                             2: "Meeting tentatively accepted"}

    # method substantially from: https://pythoninoffice.com/get-outlook-calendar-meeting-data-using-python/
    def get_meeting_info(self, begin, end):
        appointments = self.calendar.Items
        appointments.IncludeRecurrences = True
        appointments.sort('[Start]')
        restriction = "[Start] >= '" + begin.strftime('%m/%d/%Y') + "' AND [Start] <= '" + end.strftime('%m/%d/%Y') + "'"
        appointments = appointments.Restrict(restriction)
        print('MEETING INFO -->')
        for appointment in appointments:
            organizer = appointment.GetOrganizer()
            if self.userID == organizer.ID:
                print('subject: ', appointment.subject, '|', 'organizer: ', appointment.organizer)
                print('start time: ', appointment.start)
                print('recipient status -->')
                recipients = appointment.Recipients
                for person in recipients:
                    print(person.Name, 'status is: ', self.responseDict[person.MeetingResponseStatus])
