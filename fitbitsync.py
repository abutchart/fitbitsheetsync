from __future__ import print_function
import fitbit
import gather_keys_oauth2 as Oauth2
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import datetime
import webbrowser

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of the spreadsheet.
SPREADSHEET_ID = 'ADD YOUR SPREADSHEET ID HERE'
RANGE_NAME = 'A1:P'

#Fitbit ids 
CLIENT_ID = 'ADD YOUR FITBIT API ID HERE'
CLIENT_SECRET = 'ADD YOUR FITBIT API SECRET HERE'

now = datetime.datetime.now()
#for days earlier than today
#now = now - datetime.timedelta(days=1)


def main():

    ####FITBIT API STUFF

    #Oauth
    server = Oauth2.OAuth2Server(CLIENT_ID, CLIENT_SECRET)
    server.browser_authorize()

    ACCESS_TOKEN = str(server.fitbit.client.session.token['access_token'])
    REFRESH_TOKEN = str(server.fitbit.client.session.token['refresh_token'])

    authd_client = fitbit.Fitbit(CLIENT_ID, CLIENT_SECRET, access_token=ACCESS_TOKEN, refresh_token=REFRESH_TOKEN)

    sleepData = authd_client.get_sleep(now)

    if len(sleepData['sleep']) == 0:
        print('sync your phone ya dummy')
    else:

        #parsing data from sleep data
        timeInBedMins = sleepData['sleep'][0]['timeInBed']
        timeInBedHrs = '{:d}:{:02d}'.format(*divmod(timeInBedMins, 60))

        deepMins = sleepData['summary']['stages']['deep']
        lightMins = sleepData['summary']['stages']['light']
        remMins = sleepData['summary']['stages']['rem']
        awakeMins = sleepData['summary']['stages']['wake']

        deepPercent = round((deepMins / timeInBedMins) * 100)
        lightPercent = round((lightMins / timeInBedMins) * 100)
        remPercent = round((remMins / timeInBedMins) * 100)
        awakePercent = round((awakeMins / timeInBedMins) * 100)

        sleepMins = timeInBedMins - awakeMins
        sleepHrs = '{:d}:{:02d}'.format(*divmod(sleepMins, 60))

        startTime = sleepData['sleep'][0]['startTime'][-12:-7]

        #ugly time formatting thing god I hate working with time
        #split by colon
        temp = startTime.split(':')
        #hours
        temp[0] = int(temp[0])
        #minutes
        temp[1] = int(temp[1])

        #convert to 24 hr time
        if temp[0] > 12:
            temp[0] -= 12
            startTime = '{:d}:{:02d} PM'.format(temp[0], temp[1])
        #if midnight
        elif temp[0] == 0:
            temp[0] = 12
            startTime = '{:d}:{:02d} AM'.format(temp[0], temp[1])
        else:
            startTime += " AM"

        endTime = ((temp[0]*60)+temp[1]) + timeInBedMins

        if endTime >= 780:
            endTime -= 720

        endTime = '{:d}:{:02d}'.format(*divmod(endTime, 60)) + " AM"

        sleepRange = startTime + " - " + endTime

        id = str(sleepData['sleep'][0]['logId'])
        urlDate = str(now.strftime('%Y-%m-%d'))
        website = 'https://www.fitbit.com/sleep/' + urlDate + '/' + id

        comments = input('Comments:')
        mood = input('Mood:')

        actualTime = input('Actual Bed Time:')
        actualTimeSplit = actualTime.split(':')


        if not actualTime.endswith("PM"):
            actualTime = actualTime + " PM"

        webbrowser.open_new(website)

        actualSleepTime = input('Actual Sleep Time: ')
        actualSleepTime = actualSleepTime.split(':')

        latency = ((int(actualSleepTime[0])*60)+int(actualSleepTime[1])) - ((int(actualTimeSplit[0])*60)+int(actualTimeSplit[1]))

        ####GOOGLE SHEETS STUFF

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


        #build sheet
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()

        #get row number of last row
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
        numRows = len(result.get('values'))

        #create request body
        requests = []

        # delete last row
        requests.append({
            'deleteDimension': {
                "range": {
                    "startIndex": numRows-1,
                    "endIndex": numRows,
                    "dimension": "ROWS"
                }
            }
        })

        # add fitbit data
        requests.append({
            'appendCells': {
                "rows": [
                    {
                        'values': [
                            {
                                "userEnteredValue": {
                                    "stringValue": str(now.strftime('%m/%d/%y'))
                                },
                                "userEnteredFormat": {
                                    "horizontalAlignment": 'RIGHT'
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": str(sleepHrs)
                                },
                                "userEnteredFormat": {
                                    "horizontalAlignment": 'RIGHT'
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": str(sleepMins)
                                },
                                "userEnteredFormat": {
                                    "horizontalAlignment": 'RIGHT'
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": sleepRange
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": str(timeInBedHrs)
                                },
                                "userEnteredFormat": {
                                    "horizontalAlignment": 'RIGHT'
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": str(timeInBedMins)
                                },
                                "userEnteredFormat": {
                                    "horizontalAlignment": 'RIGHT'
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": str(awakeMins)
                                },
                                "userEnteredFormat": {
                                    "horizontalAlignment": 'RIGHT'
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": actualTime
                                },
                                "userEnteredFormat": {
                                    "horizontalAlignment": 'RIGHT'
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": str(latency)
                                },
                                "userEnteredFormat": {
                                    "horizontalAlignment": 'RIGHT'
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": str(awakePercent)
                                },
                                "userEnteredFormat": {
                                    "horizontalAlignment": 'RIGHT'
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": str(remPercent)
                                },
                                "userEnteredFormat": {
                                    "horizontalAlignment": 'RIGHT'
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": str(lightPercent)
                                },
                                "userEnteredFormat": {
                                    "horizontalAlignment": 'RIGHT'
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": str(deepPercent)
                                },
                                "userEnteredFormat": {
                                    "horizontalAlignment": 'RIGHT'
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": comments
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "stringValue": mood
                                }
                            }
                        ]
                    },
                    {},
                    {
                        'values': [
                            {
                                "userEnteredValue": {
                                    "stringValue": "Averages:"
                                },
                                "userEnteredFormat": {
                                    "textFormat": {
                                        "bold": True
                                    }
                                }
                            },
                            {
                                #119 if on real sheet
                                "userEnteredValue": {
                                    "formulaValue": "=AVERAGE(B119:B{})".format(numRows-1)
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "formulaValue": "=AVERAGE(C119:C{})".format(numRows - 1)
                                }
                            },
                            {},
                            {
                                "userEnteredValue": {
                                    "formulaValue": "=AVERAGE(E119:E{})".format(numRows - 1)
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "formulaValue": "=AVERAGE(F119:F{})".format(numRows - 1)
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "formulaValue": "=AVERAGE(G119:G{})".format(numRows - 1)
                                }
                            },
                            {},
                            {
                                "userEnteredValue": {
                                    "formulaValue": "=AVERAGE(I119:I{})".format(numRows - 1)
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "formulaValue": "=AVERAGE(J119:J{})".format(numRows - 1)
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "formulaValue": "=AVERAGE(K119:K{})".format(numRows - 1)
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "formulaValue": "=AVERAGE(L119:L{})".format(numRows - 1)
                                }
                            },
                            {
                                "userEnteredValue": {
                                    "formulaValue": "=AVERAGE(M119:M{})".format(numRows - 1)
                                }
                            }
                        ]
                    }
                ],
                "fields": "*"
            }
        })

        body = {
            'requests': requests
        }

        response = sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
        print(response)
        webbrowser.open_new('https://docs.google.com/spreadsheets/d/1ACwR29JfHAJG3ro8asn53B1f88A8aj5AIBtH2DKhOO8/edit#gid=0&range=A' + str(numRows-1))


if __name__ == '__main__':
    main()

