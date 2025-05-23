from email.message import EmailMessage
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.service_account import Credentials as G_Credentials
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from gspread_formatting import CellFormat, Color, format_cell_range
import base64
import gspread
import os.path

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://mail.google.com/"]

def main():
    """Credentials and login"""
    creds = None

    # Load saved credentials
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # If no valid creds, do OAuth flow
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "../creds/gmail_creds.json", SCOPES
            )
            creds = flow.run_local_server(port=0)

        # Save the credentials for next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    """Connect to the minutes google sheet"""
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    gspread_creds = G_Credentials.from_service_account_file("../creds/credentials.json", scopes=scopes)
    client = gspread.authorize(gspread_creds)

    workbook_id = "17Y8J7oHgFL8q6ZZaEJIXMuWzIuGZOTuBf6EFSR1Ipb4"
    workbook = client.open_by_key(workbook_id)
    
    # Collect name/id of latest meeting sheet
    len_sheets, latest_sheet_name, latest_sheet_id = latest_sheet_details(workbook)
    sheet = workbook.worksheet(latest_sheet_name)
    
    """Generate Message Content"""
    # Get names of recipients
    start_row = 11
    end_row = 299
    data_range = f"H{start_row}:H{end_row}"
    data = sheet.get(data_range) # Get all data in the range
    
    recipients = set()

    for i in data:
       if i and i[0] != 'Empty':  # make sure it's not an empty list or blank string
        recipients.add(i[0].strip())

    # Get e-mails of recipients
    recipient_emails = []
    sheet = workbook.worksheet("People")
    start_row = 4
    end_row = 53
    data_range = f"B{start_row}:D{end_row}"
    data = sheet.get(data_range) # Get all data in the range

    for row in data:
        name = row[0].strip()  # Column B, the name/ID
        email = row[2].strip()  # Column D, the email

        if name in recipients:
            recipient_emails.append((name,email))

    # Create a dictionary with list of outstanding items for each recipient
    sheet = workbook.worksheet(latest_sheet_name)

    start_row = 11
    end_row = 299
    data_range = f"E{start_row}:H{end_row}"
    data = sheet.get(data_range) # Get all data in the range

    tasks = {}

    for name, email in recipient_emails:
        tasks[email] = [
            [i[1],i[2]] for i in data 
            if len(i) > 3 and i[3] == name and i[1]  # Check length, name match, and non-empty task
        ]

    # Get projefct name and date of next meeting
    sheet = workbook.worksheet(latest_sheet_name)
    date = sheet.acell('C6').value
    project_name = sheet.acell('C3').value

    """Insert content and send message"""
    for name, email in recipient_emails:
        gmail_send_message(creds,email,name,tasks,project_name,date)

def latest_sheet_details(workbook):
    # Calculate number of existing sheets
    list_of_sheets = workbook.worksheets()
    len_sheets = len(list_of_sheets)

    # Collect name/id of latest meeting sheet
    latest_sheet_name = list_of_sheets[-1].title
    latest_sheet_id = list_of_sheets[-1].id

    return len_sheets, latest_sheet_name, latest_sheet_id

def gmail_send_message(creds,email,name,tasks,project_name,date):
  """Create and send an email message
  Returns: Message object, including message id
  """

  try:
    service = build("gmail", "v1", credentials=creds)
    message = EmailMessage()

    if tasks[email]:
        # Create a markdown-style table with fixed spacing
        task_lines = "\n".join(f"{status}" for task, status in tasks[email])
        task_list = f"{task_lines}"
    else:
        task_list = "None"

    message.set_content(
f"""Hi {name},

Ahead of the next project meeting for {project_name} on {date}, please see a list of your outstanding actions:

{task_list}

See you at the next meeting.

Kind regards,

Michael Fortune 
""")

    message["To"] = email
    message["From"] = "michaelfortune91@googlemail.com"
    message["Subject"] = f"{project_name} - Actions for next meeting"

    # encoded message
    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    create_message = {"raw": encoded_message}
    # pylint: disable=E1101
    send_message = (
        service.users()
        .messages()
        .send(userId="me", body=create_message)
        .execute()
    )
  except HttpError as error:
    print(f"An error occurred: {error}")
    send_message = None
  return send_message

if __name__ == "__main__":
    main()
