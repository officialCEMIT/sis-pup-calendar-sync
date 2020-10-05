# PARSERS
from selenium                                       import webdriver
from selenium.webdriver.support.select              import Select
from selenium.webdriver.support                     import expected_conditions as EC
from selenium.webdriver.common.by                   import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui                  import WebDriverWait
from bs4                                            import BeautifulSoup
import configparser

# GOOGLE-API
from googleapiclient.discovery                      import build
from google_auth_oauthlib.flow                      import InstalledAppFlow
from google.auth.transport.requests                 import Request

# STANDARD LIBS
from time                                           import sleep
from pprint                                         import pprint

import datetime, os, pickle, json

config = configparser.ConfigParser()
config.read('settings.ini')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# SETTINGS
settings = {
    'GOOGLE-CalendarName': config.get('GOOGLE', 'CalendarName'), # CALENDAR TITLE
    'GOOGLE-TimeZone': 'Asia/Manila', # http://www.timezoneconverter.com/cgi-bin/zonehelp
    'GOOGLE-GMT': '+0800', # PDT/MST/GMT
    'SIS-StudentNumber': config.get('PUP', 'StudentNumber'), # 20XX-XXXXX-MN-0
    'SIS-PWD': config.get('PUP', 'Password'),
    'SIS-BirthMonth': config.get('PUP', 'BirthMonth'), # 01-12
    'SIS-BirthDay': config.get('PUP', 'BirthDay'),
    'SIS-BirthYear': config.get('PUP', 'BirthYear'),
    'SIS-SemStart': config.get('PUP', 'SemesterStart')+"T00:00:00+0800",
    'SIS-SemEnd': config.get('PUP', 'SemesterEnd').replace("-","")+'T000000Z'
}


DAYS_OF_THE_WEEK = {
    'M': [0],
    'T': [1],
    'W': [2],
    'TH': [3],
    'F': [4],
    'S': [5],
    'SUN': [6]
}



def clear(): # CLEAR SCREEN FUNCTION
    os.system('cls' if os.name == 'nt' else 'clear')

def create_calendar(data):
    # AUTHENTICATION
    SCOPES = ['https://www.googleapis.com/auth/calendar']

    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)


    # CREATE CALENDAR
    calendar = {
    'summary': settings['GOOGLE-CalendarName'],
    'timeZone': 'Asia/Manila'
    }

    created_calendar = service.calendars().insert(body=calendar).execute()
    print("CREATED CALENDAR:",created_calendar['id'])
    for i in data:
        print(i)
        start_time = (datetime.datetime.strptime(DAYS_OF_THE_WEEK[i['day']][1], '%Y-%m-%dT%H:%M:%S%z').replace(hour=int(i['start-time'][0]),
                                                          minute=int(i['start-time'][1]),
                                                          second=int(i['start-time'][2]))).isoformat()
        end_time = (datetime.datetime.strptime(DAYS_OF_THE_WEEK[i['day']][1], '%Y-%m-%dT%H:%M:%S%z').replace(hour=int(i['end-time'][0]),
                                                          minute=int(i['end-time'][1]),
                                                          second=int(i['end-time'][2]))).isoformat()
        try:
            event = {
                'summary': i['description'],
                'description': f"{i['subject_code']} ",#@{i['room']}
                # 'location': i['location'],
                'start': {
                    'dateTime': start_time ,
                    'timeZone': settings['GOOGLE-TimeZone'],
                  },
                  'end': {
                    'dateTime': end_time ,
                    'timeZone': settings['GOOGLE-TimeZone'],
                  },
                  'reminders': {
                    'useDefault': False,
                    'overrides': [
                      {'method': 'popup', 'minutes': 60},
                    ],
                  },
                'recurrence': [
                    'RRULE:FREQ=WEEKLY;UNTIL=%s' % (settings['SIS-SemEnd']),
                ]
            }
            event = service.events().insert(calendarId=created_calendar['id'], body=event).execute()
            print('Event created: %s' % (event.get('htmlLink')))
        except Exception as e:
            service.calendars().delete(calendarId=created_calendar['id']).execute()
            raise e

def dict_data(data):
    y = []
    for i in data:
        j = []
        for sched in i[3]:
            start_time = str(datetime.datetime.strptime(sched[1][0],'%I:%M%p').time()).split(':')
            end_time = str(datetime.datetime.strptime(sched[1][1],'%I:%M%p').time()).split(':')
            x = {
                "subject_code": i[0],
                "description": i[1],
                # "room": sched[2][0],
                "section": i[2],
                # "location": sched[2][1],
                "day": sched[0],
                "start-time": start_time,
                "end-time": end_time
            }
            y.append(x)

    return y

def pup_locator(data):
    location_map = {
        "CEA" : "Polytechnic University of the Philippines - College of Engineering and Architecture, Anonas, Sta. Mesa, Maynila, 1008 Kalakhang Maynila",
        "FIELD": "Polytechnic University of the Philippines (Main Campus), Anonas, Santa Mesa, Manila, Metro Manila",
    }
    loc = ""
    for key in location_map:
        if key in data:
            loc = [data,location_map[key]]
            break
        else:
            loc = [data,location_map['FIELD']]
            break

    return loc

def location_handler(data,sched):
    time_map = ['07:00AM','07:30AM','08:00AM','08:30AM','09:00AM','09:30AM',
          '10:00AM','10:30AM','11:00AM','11:30AM','12:00PM','12:30PM',
          '01:00PM','01:30PM','02:00PM','02:30PM','03:00PM','03:30PM',
          '04:00PM','04:30PM','05:00PM','05:30PM','06:00PM','06:30PM',
          '07:00PM','07:30PM','08:00PM','08:30PM','09:00PM','09:30PM',]
    day_map = ['M','T','W','TH','F','S','SUN']
    rows = data.find_all('tr')

    locations = []
    for t, time in enumerate(sched['time']):
        try:
            for index,value in enumerate(time_map):
                if time[0] == value:
                    for d,day in enumerate(day_map):
                        for td in rows[1+index].find_all('td'):
                            if sched['subject'] in td:
                                x = BeautifulSoup(str(td).replace('<br/>',' '),features='html.parser').getText().split(' ')[-1]
                                x = pup_locator(x)
                                locations.append(x)
                                raise StopIteration
        except StopIteration: pass

    return locations

def sis_connect():
    BASE_URL = ''
    if os.name == 'nt':
        d = 'tools/chromedriver.exe'
    else:
        d = os.path.join(BASE_DIR, 'chromedriver')
    driver = webdriver.Chrome(d)
    try:
        driver.get("http://sisstudents.pup.edu.ph/")
        while True:
            try:
                WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "introduction")))
            except:
                print("TIMEOUT")
            finally:
                BASE_URL = driver.current_url
                driver.get(BASE_URL + 'student/')
                break

        driver.find_element_by_name("studno").send_keys(settings['SIS-StudentNumber'])
        Select(driver.find_element_by_name("SelectMonth")).select_by_value(settings['SIS-BirthMonth']) #value 01-12
        Select(driver.find_element_by_name("SelectDay")).select_by_visible_text(settings['SIS-BirthDay']) #value 1-31
        Select(driver.find_element_by_name("SelectYear")).select_by_visible_text(settings['SIS-BirthYear']) #value 1940-2020
        driver.find_element_by_name("password").send_keys(settings['SIS-PWD'])
        driver.find_element_by_name('Login').click()

        while True:
            try:
                WebDriverWait(driver, 60).until(EC.title_contains("Home"))
            except:
                print("UNABLE TO LOGIN")
            finally:
                print("SUCCESSFUL LOGIN")
                break

    except Exception as e:
        print("ERROR ON SIGNIN IN:", e)
    
    # NAVIGATE TO SCHEDULE
    try:
        # WebDriverWait(driver, 30).until(EC.alert_is_present(),
        #                         'Timed out waiting for PA creation ' +
        #                         'confirmation popup to appear.')
        # driver.switch_to.alert.accept()
        # WebDriverWait(driver, 30).until(EC.title_contains("Message"))
        driver.get(BASE_URL + "student/schedule")
        # driver.find_element_by_xpath("/html/body/table/tbody/tr/td/table/tbody/tr[2]/td[2]/table/tbody/tr[2]/td/table/tbody/tr[2]/td[2]/a[4]").click()
    except Exception as e:
        print("NO SIS POPUP",e)

    try:
        WebDriverWait(driver, 30).until(EC.title_contains("Schedule"))
    except Exception as e:
        print("ERROR ON SCHEDULE PARSING:", e)
    finally:

        #######################################
        ## COLLECT SUBJ, SUBJ_CODE, DAY/TIME ##
        #######################################
        scheds = []
        table_id = driver.find_element_by_id('Subject')
        rows = table_id.find_elements(By.TAG_NAME, "tr")

        # loc = driver.find_element_by_xpath("/html/body/table/tbody/tr/td/table/tbody/tr[2]/td[2]/table/tbody/tr[3]/td/table/tbody/tr[3]/td/table/tbody/tr/td/form/table/tbody/tr[3]/td")
        # table2 = BeautifulSoup(loc.get_attribute('outerHTML'), features="html.parser")

        for row in rows[1:]:
            subject_code = row.find_elements(By.TAG_NAME, "td")[1].text
            subject = row.find_elements(By.TAG_NAME, "td")[2].text #
            schedule = row.find_elements(By.TAG_NAME, "td")[6].text.split() # SCHEDULE COLUMN
            print(subject,"\t",subject_code,"\t",schedule)

            if len(schedule) > 1:
                day = schedule[4].split('/')
                time = schedule[5].split('/')
                time = [i.split('-') for i in time]
                # locations = "LOC" #location_handler(table2,{"subject":subject_code,"time":time})
                days = list(zip(day,time))

                scheds.append([subject_code,subject,schedule[3],days])
        
        pprint(scheds)
        # print(scheds[0])
        return scheds
    # os.system("echo Press enter to continue; read dummy;")
    driver.close()
def main():
    clear()
    # FIRST SCHOOL YEAR DATES

    d = datetime.datetime.strptime(settings['SIS-SemStart'], '%Y-%m-%dT%H:%M:%S%z')

    # ASSIGNS FIRST DATES OF THE FIRST WEEK ASSUMING EVERY SEMESTER STARTS ON A MONDAY
    loop = 0
    while loop < 7:
        for abbv, value in DAYS_OF_THE_WEEK.items():
            if len(value) < 2:
                if value[0] == d.weekday():
                    value.append(d.isoformat("T"))
                    d += datetime.timedelta(days=1)
                    loop += 1

    scheds = sis_connect()
    schedules = dict_data(scheds) # CLEAN DATA

    create_calendar(schedules)



if __name__ == "__main__": #EXECUTION POINT
    main()
