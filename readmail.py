import win32com.client
import datetime as dt
import datefinder
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
outlook_event = win32com.client.Dispatch("Outlook.Application")


inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
recipient = outlook.createRecipient("Tanya.Shourya@dell.com")
resolved = recipient.Resolve()
sharedCalendar = outlook.GetSharedDefaultFolder(recipient, 9).Folders("SAP general dates")
messages = inbox.Items
# dates earliest: ....
#birthdays, holiday list
#check if already exists -> ignore
#check why date is not populated
#add to another calendar
message_today = messages.restrict("[SentOn]>'6/23/2020 12:00 AM'")
for message in message_today:
    body_content = message.body
    subject = message.Subject
    event = ('PROPEL Awareness Notice' in subject)
    if event == True:
        print(subject)
        matches = list(datefinder.find_dates(body_content))
        matches.sort( reverse = False)
        for match in matches:
            print(match)
        appointment = sharedCalendar.Items.Add(1)
        appointment.Start = matches[0]
        appointment.Subject = subject
        if matches[-1] == matches[0]:
            appointment.Duration = 30
        else:
            appointment.End = matches[-1]
        appointment.Location = subject
        appointment.ReminderSet = True
        appointment.ReminderMinutesBeforeStart = 15
        appointment.Save()