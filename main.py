
import imapy, re, datetime
from imapy.query_builder import Q
from pushbullet import pushbullet
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools

# Information for chcking email
server = 'imap-mail.outlook.com'
user = 'johnmarsh19@hotmail.com'
pw = 'PASSWORD'
sender = 'EMAIL'

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/calendar'


#Pushbullet Information
pb = pushbullet.Pushbullet("o.gEn2Z006PCOhAYb16VLFIRJ2oEngzPLV")


def create_google_calender_event(start_time,end_time,day,month, year, service):
    text = str(month) + '/' + str(day[:-1]) + '/' + str(year)
    lis = service.events().list(calendarId='CALENDER').execute()

    if text not in str(lis):

        event = {
            'summary': 'Work',
            'description': text,
            'start': {
                'dateTime': str(start_time),
                'timeZone': 'America/Chicago',
            },
            'end': {
                'dateTime': str(end_time),
                'timeZone': 'America/Chicago',
            },
        }

        event = service.events().insert(calendarId='CALENDER', body=event).execute()
        print('Added Work Event: ' + str(start_time) + " to " + str(end_time))






def convert_month_to_num(month):
    monthDict = {'January': 1,'February': 2,'March': 3,'April': 4,'May': 5,'June': 6,'July': 7,'August': 8,'September': 9,'October': 10,'November': 11,'December': 12}


    return monthDict.get(month)


def convert_to_google_time(time):
    string = time

    if 'PM' in time:
        new_time = int(time.split(':')[0]) + 12
        string = str(new_time) + ':' + (time.split(':')[1])


    return string[:-3]

def get_formatted_schedule (sender, server, user,pw):
    string = ""

    imap = imapy.connect(host=server, username=user, password=pw, ssl=True)


    # Object needed to find emails
    q = Q()
    # Gets all unseen emails sent by me from my work e-mail
    emails = imap.folder('Inbox').emails(
        q.sender(sender).unseen()
    )


    regex = r'\d{1,2}:\d{2} AM|\d{1,2}:\d{2} PM'

    for email in emails:
        text = str(email['text']).split('\\r\\n')
        start_time = ""
        end_time = ""
        y = 0
        x = 0
        for i in text:
            x += 1

            if 'text_normalized' in i:
                break
            elif 'Activity' in i:
                y+= 1
            elif i == "" and 'Phone Time' in text[x-2]:
                match = re.findall(regex, text[x-2])
                end_time = match[1]
                string = string + start_time + " - " + end_time + "\n"
            elif ('Phone Time' in i or 'Break' in i or 'Lunch' in i or 'Meeting' in i or 'Flash Training' in i or 'Huddle' in i) and (y == 2):
                matches = re.findall(regex, i)
                if start_time == "":
                    start_time = matches[0]
            elif ('Monday' in i or 'Tuesday' in i or 'Wednesday' in i or 'Thursday' in i or 'Friday' in i or 'Saturday' in i) and ('Sent' not in i):
                string = string + "\n"
                string = string + i + "\n"
                start_time = ""
                y = 0
            #print(string)

    return string

    #smtpObj.sendmail(user, user, 'Subject: Work Schedule\n' + string)
    #channel = pb.get_channel('workschedule')
    #channel.push_note("Work Schedule", string)

def get_google_formatted_schedule (sender, server, user,pw, service):

    string = ""
    day = ''
    month = ''
    year = ''

    imap = imapy.connect(host=server, username=user, password=pw, ssl=True)


    # Object needed to find emails
    q = Q()
    # Gets all unseen emails sent by me from my work e-mail
    emails = imap.folder('Inbox').emails(
        q.sender(sender).unseen()
    )


    regex = r'\d{1,2}:\d{2} AM|\d{1,2}:\d{2} PM'

    for email in emails:
        text = str(email['text']).split('\\r\\n')
        start_time = ""
        end_time = ""
        y = 0
        x = 0
        for i in text:
            x += 1
            if 'text_normalized' in i:
                break
            elif 'Activity' in i:
                y+= 1
            elif i == "" and 'Phone Time' or 'General' in text[x-2]:
                match = re.findall(regex, text[x-2])
                print(text)
                end_time = match[1]

                print(string + convert_to_google_time(start_time))
                create_google_calender_event(string + convert_to_google_time(start_time) + ':00' + '-' +'05:00',  string + convert_to_google_time(end_time) + ':00' + '-' + '05:00', day,month,year, service)
                string = ""
            elif ('Phone Time' in i or 'Break' in i or 'Lunch' in i or 'Meeting' in i or 'Flash Training' in i or 'Huddle' in i) and (y == 2):
                matches = re.findall(regex, i)
                if start_time == "":
                    start_time = matches[0]
            elif ('Monday' in i or 'Tuesday' in i or 'Wednesday' in i or 'Thursday' in i or 'Friday' in i or 'Saturday' in i) and ('Sent' not in i):
                if 'Friday' in i:
                    print(i)
                month = convert_month_to_num(i.split(' ')[1])
                day = i.split(' ')[2]
                year = i.split(' ')[3]

                string = string + ""
                #if('T' not in string):
                string = '{}{}-{:02}-{:02}T'.format(string,year, month, int(day[:-1]))

                start_time = ""
                y = 0

    return string

def main():
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('calendar', 'v3', http=creds.authorize(Http()))

    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time

    get_google_formatted_schedule(sender, server, user, pw, service)

    print(get_formatted_schedule(sender, server, user, pw))
    print("HELLO")


if __name__ == '__main__':
    main()
