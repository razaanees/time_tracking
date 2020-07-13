from __future__ import print_function
from apiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
from apiclient.http import MediaIoBaseDownload, MediaFileUpload
import io, re, csv
from datetime import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from pprint import pprint

def convert_td(tim):
    return tim.days*24+tim.seconds//3600+(tim.seconds//60)%60/60

# Setup the Drive v3 API and get credentials
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']
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

service = build('drive', 'v3', credentials=creds)
service_sheets = build('sheets', 'v4', credentials=creds)

# Ensure id of work data still the same
results = service.files().list(
    pageSize=20, fields="nextPageToken, files(id, name)").execute()
items = results.get('files', [])

work_id = task_id = "Not Found"

for item in items:
    if item["name"] == "timerec-workunits-pro.txt":
        work_id = item["id"]
    if item["name"] == "timerec-tasks-pro.txt":
        task_id = item["id"]

print("Work ID: {}\nTask ID: {}".format(work_id, task_id))

if work_id == 'Not Found' or task_id == "Not Found":
    print(items)

# Download the latest file of task list from Google Drive
task_down = service.files().get_media(fileId=task_id)
fh = io.BytesIO()
downloader = MediaIoBaseDownload(fh,task_down)
done = False
while done is False:
    status, done = downloader.next_chunk()
    print("Task list download progress: {}%".format(int(status.progress())*100))

fh.seek(0)
with open("task_list.txt", "wb") as f:
    f.write(fh.read())

# Get an updated task name-id relationship
with open("task_list.txt", "r+") as r:
    t = r.readlines()

    tasks = {}
    for line in t:
        if line[0].isdigit():
            new_line = line.split("|")
            tasks[new_line[0]] = new_line[1]

print("Task-Id relationship has been refreshed")
print(tasks)

# Download the task hours from Google Drive
request = service.files().get_media(fileId=work_id)
fh = io.BytesIO()
downloader = MediaIoBaseDownload(fh,request)
done = False
while done is False:
    status, done = downloader.next_chunk()
    print("Work list download progress: {}%".format(int(status.progress())*100))

fh.seek(0)
with open("work_list.csv", "wb") as f:
    f.write(fh.read())

# Parse the hours log as a list of dictionaries
with open("work_list.csv", "r+") as file:

    reader = csv.DictReader(file, delimiter=",")
    hour_tracking = []

    for row in reader:
        # Not calculating by week number anymore
        #iso_cal = datetime.date(datetime.strptime(row["# DATE"],"%Y-%m-%d")).isocalendar()
        cal_date = datetime.strptime(row["# DATE"], "%Y-%m-%d")
        day = {}
        day["Year"], day["Month"], day["Week"], day["Day"] = cal_date.strftime("%Y"), cal_date.strftime("%m"), cal_date.strftime("%W"), cal_date.strftime("%d")
        day["Task"] = row["TASK-ID"]
        day["Duration"] = datetime.strptime(row["CHECKOUT"],"%Y-%m-%d %H:%M:%S")-datetime.strptime(row["CHECKIN"],"%Y-%m-%d %H:%M:%S")
        hour_tracking.append(day)

days = set()
all_days = []
exclude = ["Year", "Month", "Day"]

for line in hour_tracking:
    days.add(line["Year"] + line["Month"] + line["Day"])

for d in sorted(days):
    output = {}
    for item in hour_tracking:
        if item["Year"] + item["Month"] + item["Day"] == d:
            if tasks[item["Task"]] in output.keys():
                output[tasks[item["Task"]]] += item["Duration"]
            else:
                output[tasks[item["Task"]]] = item["Duration"]
        output["Year"] = d[:4]
        output["Month"] = d[4:6]
        output["Day"] = d[6:]
    for k in output.keys():
        if k not in exclude:
            output[k] = convert_td(output[k])
    all_days.append(output)

# get the categories from the header of the Week Optimizer file
spreadsheet_id = '1SeCU_7GEZ_1MMh4i5jKXC7i2-RnlUxj-PaZBmsxXKXU'
range_name = 'Data'

result = service_sheets.spreadsheets().values().get(
    spreadsheetId=spreadsheet_id, range=range_name).execute()

categories = (result["values"][0])

all_values = []
for i in range(len(all_days)):
    values = []
    for c in categories:
        if c in all_days[i].keys():
            values.append(all_days[i][c])
        else:
            values.append(0)
    all_values.append(values)

# update the Week Optimizer sheet with current values

for i in range(len(all_values)):

    body = {
        'values': [all_values[i]]
    }
    value_input_option = "USER_ENTERED"

    result = service_sheets.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id, range=range_name,
        valueInputOption=value_input_option, body=body).execute()
print('{0} cells appended.'.format(result \
                                       .get('updates') \
                                       .get('updatedCells')));
