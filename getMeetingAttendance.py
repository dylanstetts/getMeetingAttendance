import requests
import msal
import datetime
import csv
import logging
import re

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Azure AD App Credentials
TENANT_ID = 'xxxxxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxxx'
CLIENT_ID = 'xxxxxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxxxx'
CLIENT_SECRET = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxx'

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"

# Prompt for user input
UPN = input("Enter the target user's UPN (e.g., user@example.com): ").strip()

# Prompt for time range with validation loop
while True:
    time_range_input = input("Enter the time range (e.g., 30 days, 6 weeks, 3 months): ").strip().lower()
    match = re.fullmatch(r'(\d+)\s*(day|week|month)s?', time_range_input)
    if match:
        value, unit = int(match.group(1)), match.group(2)
        if unit == 'day':
            delta = datetime.timedelta(days=value)
        elif unit == 'week':
            delta = datetime.timedelta(weeks=value)
        elif unit == 'month':
            delta = datetime.timedelta(days=30 * value)
        break
    else:
        print("Invalid time range format. Please use formats like '30 days', '6 weeks', or '3 months'.")

# Authenticate
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)
token_response = app.acquire_token_for_client(scopes=SCOPE)
access_token = token_response.get("access_token")

if not access_token:
    logging.error("Failed to acquire access token.")
    exit()

headers = {
    'Authorization': f'Bearer {access_token}',
    'Content-Type': 'application/json'
}

# Resolve UPN to user ID
user_lookup_url = f"{GRAPH_API_ENDPOINT}/users/{UPN}"
user_response = requests.get(user_lookup_url, headers=headers)
if user_response.status_code != 200:
    logging.error(f"Failed to resolve user ID: {user_response.status_code} - {user_response.text}")
    exit()

user_id = user_response.json().get('id')
logging.info(f"👤 Resolved user ID: {user_id}")

# Date range
end_date = datetime.datetime.utcnow()
start_date = end_date - delta

# Get calendar events
events_url = f"{GRAPH_API_ENDPOINT}/users/{user_id}/calendar/calendarView?startDateTime={start_date.isoformat()}Z&endDateTime={end_date.isoformat()}Z"
events = []

while events_url:
    response = requests.get(events_url, headers=headers)
    if response.status_code != 200:
        logging.error(f"Failed to retrieve events: {response.status_code} - {response.text}")
        break
    data = response.json()
    events.extend(data.get('value', []))
    events_url = data.get('@odata.nextLink', None)

logging.info(f"Total events retrieved: {len(events)}")

# Filter online meetings and get attendance
attendance_data = []

for event in events:
    if event.get('isOnlineMeeting') and event.get('onlineMeeting', {}).get('joinUrl'):
        join_url = event['onlineMeeting']['joinUrl']
        subject = event.get('subject')
        start_time = event.get('start', {}).get('dateTime')
        end_time = event.get('end', {}).get('dateTime')

        logging.info(f"Looking up online meeting by join URL: {join_url}")

        # Lookup online meeting by join URL using user ID
        lookup_url = f"{GRAPH_API_ENDPOINT}/users/{user_id}/onlineMeetings?$filter=joinWebUrl eq '{join_url}'"
        lookup_response = requests.get(lookup_url, headers=headers)

        if lookup_response.status_code != 200:
            logging.warning(f"Failed to lookup meeting by join URL: {lookup_response.status_code} - {lookup_response.text}")
            continue

        meetings = lookup_response.json().get('value', [])
        if not meetings:
            logging.warning("No matching online meeting found for join URL.")
            continue

        meeting_id = meetings[0].get('id')
        logging.info(f"Found meeting ID: {meeting_id}")

        # Get attendance reports
        attendance_reports_url = f"{GRAPH_API_ENDPOINT}/users/{user_id}/onlineMeetings/{meeting_id}/attendanceReports"
        reports_response = requests.get(attendance_reports_url, headers=headers)
        logging.info(f"Attendance reports status: {reports_response.status_code}")
        if reports_response.status_code != 200:
            logging.warning(f"Failed to retrieve attendance reports: {reports_response.text}")
            continue

        reports = reports_response.json().get('value', [])
        for report in reports:
            report_id = report.get('id')
            logging.info(f"Found attendance report ID: {report_id}")

            # Get attendance records
            records_url = f"{attendance_reports_url}/{report_id}/attendanceRecords"
            records_response = requests.get(records_url, headers=headers)
            logging.info(f"Attendance records status: {records_response.status_code}")
            if records_response.status_code != 200:
                logging.warning(f"Failed to retrieve attendance records: {records_response.text}")
                continue

            records = records_response.json().get('value', [])
            for record in records:
                attendance_data.append({
                    'Meeting Subject': subject,
                    'Start Time': start_time,
                    'End Time': end_time,
                    'Attendee': record.get('identity', {}).get('displayName'),
                    'Email': record.get('emailAddress'),
                    'Join Time': record.get('joinDateTime'),
                    'Leave Time': record.get('leaveDateTime'),
                    'Duration (minutes)': record.get('totalAttendanceDurationInSeconds') // 60 if record.get('totalAttendanceDurationInSeconds') else ''
                })


# Save to CSV
csv_file = 'meeting_attendance.csv'
with open(csv_file, 'w', newline='', encoding='utf-8') as file:
    writer = csv.DictWriter(file, fieldnames=[
        'Meeting Subject', 'Start Time', 'End Time', 'Attendee', 'Email',
        'Join Time', 'Leave Time', 'Duration (minutes)'
    ])
    writer.writeheader()
    writer.writerows(attendance_data)

logging.info(f"Attendance data saved to '{csv_file}'")
