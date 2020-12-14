from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
import datetime
from matplotlib import pyplot as plt

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SHEETS = [['1ipfuXZYv18ft9Wfp5RYSz6akmTWEnyoiaIjjrkNLTJs', 'A2:G729'],
          ['1LWxmfVmTqNvv1Gnooin1coNtZwgCAe5oNwSjFXDnGJM', 'A2:H341']]


def get_data(SAMPLE_SPREADSHEET_ID='1ipfuXZYv18ft9Wfp5RYSz6akmTWEnyoiaIjjrkNLTJs', SAMPLE_RANGE_NAME='A2:G729'):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
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
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()

    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME, valueRenderOption='UNFORMATTED_VALUE').execute()
    values = result.get('values', [])

    if not values:
        return 'No data found.'
    else:
        return values


data = list()
for k in SHEETS:
    data1 = get_data(k[0], k[1])
    for i in data1:
        if k[0] == '1LWxmfVmTqNvv1Gnooin1coNtZwgCAe5oNwSjFXDnGJM':
            del i[6]
        data.append(i)

student_names = list(sorted(set([i[2] for i in data])))
payments = []
summaries = {"Payments summary": 0, "Lessons summary": 0}
all_payments = []
all_lessons = []
all_dates = []

# (datetime.datetime(1899, 12, 30) + datetime.timedelta(days=i[0])).date()
month_list = [
    [datetime.date(2019, 10, 1), datetime.date(2019, 10, 31)],
    [datetime.date(2019, 11, 1), datetime.date(2019, 11, 30)],
    [datetime.date(2019, 12, 1), datetime.date(2019, 12, 31)],
    [datetime.date(2020, 1, 1), datetime.date(2020, 1, 31)],
    [datetime.date(2020, 2, 1), datetime.date(2020, 2, 29)],
    [datetime.date(2020, 3, 1), datetime.date(2020, 3, 31)],
    [datetime.date(2020, 4, 1), datetime.date(2020, 4, 30)],
    [datetime.date(2020, 5, 1), datetime.date(2020, 5, 31)],
    [datetime.date(2020, 6, 1), datetime.date(2020, 6, 30)],
    [datetime.date(2020, 7, 1), datetime.date(2020, 7, 31)],
    [datetime.date(2020, 8, 1), datetime.date(2020, 8, 31)],
    [datetime.date(2020, 9, 1), datetime.date(2020, 9, 30)],
    [datetime.date(2020, 10, 1), datetime.date(2020, 10, 31)],
    [datetime.date(2020, 11, 1), datetime.date(2020, 11, 30)],
    [datetime.date(2020, 12, 1), datetime.date(2020, 12, 31)], ]
sum_lessons = []
for dates in month_list:
    summary = []
    stud = []
    for i in data:
        if dates[0] <= (datetime.date(1899, 12, 30) + datetime.timedelta(days=i[0])) <= dates[1]:
            summary.append(i[6])
            stud.append(i[2])
    summary.sort()
    for i in range(3):
        del summary[0]
        del summary[len(summary) - 1]
    sum_lessons.append([dates[0], dates[1], sum(summary), len(set(stud))])

print(all_dates)
for name in student_names:
    summary = [0, 0]
    for dt in range(len(data)):
        if data[dt][2].lower().replace(' ', '') == name.lower().replace(' ', ''):
            summary[0] += data[dt][1]
            summary[1] += data[dt][6]
            data[dt][2] = 'none'
    # print(name, '--', summary[0], '--', summary[1])
    all_payments.append(summary[0])
    all_lessons.append(summary[1])
    payments.append([name, summary[0], summary[1]])

all_payments.sort()
all_lessons.sort()

payments.append(["Сумми : ", sum(all_payments), sum(all_lessons)])
payments.append(["Середній прихід від студента: ", sum(all_payments) / len(student_names)])
payments.append(["Середня кількість уроків студента: ", sum(all_lessons) / len(student_names)])

for i in range(4):
    del all_payments[i]
    del all_payments[len(all_payments) - i - 1]

    del all_lessons[i]
    del all_lessons[len(all_payments) - i - 1]

payments.append(["Середній прихід від студента (Артефакт): ", sum(all_payments) / len(student_names)])
payments.append(["Середня кількість уроків студента (Артефакт): ", sum(all_lessons) / len(student_names)])

payments.append([i[0].strftime("%m.%Y") for i in sum_lessons])
payments.append([i[2] for i in sum_lessons])
payments.append([i[3] for i in sum_lessons])

df = pd.DataFrame(payments)
writer = pd.ExcelWriter('Students_payments.xlsx')
df.to_excel(writer, 'Sheet1')
writer.save()
