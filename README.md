# Microsoft Graph API Meeting Attendance Script

This Python script retrieves Microsoft Teams meeting attendance data for a specified user over a configurable time range using Microsoft Graph API with application permissions.

## Script Overview

The script will:

1. Authenticate using MSAL.
2. Query the user's calendar for events in the last 6 months.
3. Filter for online meetings using the joinWebURL.
4. Retrieve attendance reports (if available).

## Features

- Authenticate using MSAL with application credentials
- Prompt for target user's UPN (email address)
- Prompt for time range (e.g., `30 days`, `6 weeks`, `3 months`)
- Resolve UPN to Azure AD object ID
- Retrieve calendar events and filter for online meetings
- Match join URLs to online meeting IDs
- Retrieve attendance reports and attendance records
- Export results to a CSV file

## Prerequisites

- Python 3.7+
- Azure AD App Registration with the following:
  - Application permissions:
    - `OnlineMeetings.Read.All`
    - `Calendars.Read`
  - Admin consent granted
- Application Access Policy configured in Microsoft Teams PowerShell:
    - https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy

```powershell
# Example PowerShell commands
Connect-MicrosoftTeams
New-CsApplicationAccessPolicy -Identity "MyAppPolicy" -AppIds "YOUR-APP-ID" -Description "Allow app to access online meetings"
Grant-CsApplicationAccessPolicy -PolicyName "MyAppPolicy" -Identity "user@domain.com"
```

- Python libraries:
    - requests
    - MSAL (Microsoft Authentication Library)

## Installation

1. Clone the repository or download the script.
2. Install MSAL:

```python
pip install msal requests
```

## Usage

Run the script:

```python
python getMeetingAttendance.py
```

You will be prompted to enter:

- The target user's UPN (e.g., user@example.com)
- A time range (e.g., 30 days, 6 weeks, 3 months)

The script will generate a meeting_attendance.csv file with the attendance data.

## Output

The CSV file includes:

- Meeting Subject
- Start Time
- End Time
- Attendee Name
- Email
- Join Time
- Leave Time
- Duration (minutes)

## Notes 

- Attendance reports are only available for meetings created by the user and joined by participants.
- The script assumes meetings were created in Microsoft Teams and are accessible via Graph API.
- To find meetings that a user didn't organize, you would need to use an application with delegated permisisons, or use Chat/Transcript data
