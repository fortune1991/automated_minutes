import gspread
from google.oauth2.service_account import Credentials

def main():
    # Connect to the Google Sheet
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file("creds/credentials.json", scopes=scopes)
    client = gspread.authorize(creds)

    workbook_id = "17Y8J7oHgFL8q6ZZaEJIXMuWzIuGZOTuBf6EFSR1Ipb4"
    workbook = client.open_by_key(workbook_id)

    # Calculate number of existing sheets
    list_of_sheets = workbook.worksheets()
    len_sheets = len(list_of_sheets)

    # Create new sheet for "Meeting 1"
    workbook.duplicate_sheet(source_sheet_id=0,insert_sheet_index=(len_sheets+1),new_sheet_name="Meeting Minutes 1")

    # Collect variables from Project information sheet
    sheet = workbook.worksheet("Project Information and setup")
    project_name = sheet.acell("C3").value
    meeting_no = 1

    # Auto complete new sheet with above variables
    sheet = workbook.worksheet("Meeting Minutes 1")
    sheet.update('C3', [[project_name]])
    sheet.update('C4', [[meeting_no]])

if __name__ == "__main__":
    main()

