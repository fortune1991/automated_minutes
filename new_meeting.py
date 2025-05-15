import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import CellFormat, Color, format_cell_range

def main():
    # Connect to the Google Sheet
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file("creds/credentials.json", scopes=scopes)
    client = gspread.authorize(creds)

    workbook_id = "17Y8J7oHgFL8q6ZZaEJIXMuWzIuGZOTuBf6EFSR1Ipb4"
    workbook = client.open_by_key(workbook_id)
    
    # Collect name/id of latest meeting sheet
    len_sheets, latest_sheet_name, latest_sheet_id = latest_sheet_details(workbook)

    # Collect variables from existing sheet
    sheet = workbook.worksheet(f"{latest_sheet_name}")
    project_name = sheet.acell('C3').value
    meeting_no = sheet.acell('C4').value
 
    # Identify sheet_no for new meeting sheet
    sheet_no = int(meeting_no) + 1
    meeting_no = sheet_no

    # Create new sheet for meeting
    workbook.duplicate_sheet(source_sheet_id=0,insert_sheet_index=(len_sheets+1),new_sheet_name=f"Meeting Minutes {sheet_no}")

    # Refresh sheet data input where required
    sheet = workbook.worksheet(f"Meeting Minutes {sheet_no}")
    sheet.update('C3', [[project_name]])
    sheet.update('C4', [[meeting_no]])
        
def latest_sheet_details(workbook):
    # Calculate number of existing sheets
    list_of_sheets = workbook.worksheets()
    len_sheets = len(list_of_sheets)

    # Collect name/id of latest meeting sheet
    latest_sheet_name = list_of_sheets[-1].title
    latest_sheet_id = list_of_sheets[-1].id

    return len_sheets, latest_sheet_name, latest_sheet_id
    

if __name__ == "__main__":
    main()