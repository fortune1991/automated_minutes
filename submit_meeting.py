import gspread
from google.oauth2.service_account import Credentials
from gspread_formatting import CellFormat, Color, format_cell_range

def main():
    # Connect to the Google Sheet
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_file("../creds/credentials.json", scopes=scopes)
    client = gspread.authorize(creds)

    workbook_id = "17Y8J7oHgFL8q6ZZaEJIXMuWzIuGZOTuBf6EFSR1Ipb4"
    workbook = client.open_by_key(workbook_id)
    
    # Collect name/id of latest meeting sheet
    len_sheets, latest_sheet_name, latest_sheet_id = latest_sheet_details(workbook)

    # Collect data from existing sheet
    sheet = workbook.worksheet(f"{latest_sheet_name}")
    meeting_no = sheet.acell('C4').value
    meeting_purpose = sheet.acell('C5').value
    meeting_date = sheet.acell('C6').value
    next_meeting = sheet.acell('C7').value
    
    # Add existing sheet to Record of Meetings
    sheet = workbook.worksheet("Record of Meetings")
    x = len_sheets - 4
    sheet.update_acell(f'B{x}',meeting_no)
    sheet.update_acell(f'C{x}',meeting_purpose)
    sheet.update_acell(f'D{x}',meeting_date)
    fmt = CellFormat(backgroundColor=Color(0.7176, 0.8824, 0.8039))
    format_cell_range(sheet, f'B{x}:D{x}', fmt)
 
    # Identify sheet_no for new meeting sheet
    sheet_no = len_sheets - 5

    # Create new sheet for next meeting
    if next_meeting:
        workbook.duplicate_sheet(source_sheet_id=latest_sheet_id,insert_sheet_index=(len_sheets+1),new_sheet_name=f"Meeting Minutes {sheet_no}")

        # Collect variables from existing sheet
        meeting_no = sheet_no

        # Refresh sheet data input where required
        sheet = workbook.worksheet(f"Meeting Minutes {sheet_no}")
        sheet.update(range_name='C4', values=[[meeting_no]])
        sheet.update(range_name='C5', values=[[""]])
        sheet.update(range_name='C6', values=[[next_meeting]])
        sheet.update(range_name='C7', values=[[""]])
        values = [['Empty']] * 50
        sheet.update(range_name='B11:B60', values=values)

        # Delete closed items
        delete_closed_items(sheet)

def latest_sheet_details(workbook):
    # Calculate number of existing sheets
    list_of_sheets = workbook.worksheets()
    len_sheets = len(list_of_sheets)

    # Collect name/id of latest meeting sheet
    latest_sheet_name = list_of_sheets[-1].title
    latest_sheet_id = list_of_sheets[-1].id

    return len_sheets, latest_sheet_name, latest_sheet_id
    
def delete_closed_items(sheet):
    # Define the range of the table
    start_row = 10
    end_row = 290
    data_range = f"E{start_row}:J{end_row}"
    
    # Get all data in the range
    data = sheet.get(data_range)
    
    # Identify rows to keep (where status is not "Closed")
    rows_to_keep = []
    for row in data:
        if len(row) >= 6 and row[5].strip().lower() != "closed":  # Column J is index 5 (0-based)
            rows_to_keep.append(row)
    
    # Calculate how many rows we've deleted
    deleted_rows = len(data) - len(rows_to_keep)
    
    if deleted_rows == 0:
        return
    
    # Update item numbers sequentially only for rows with data 
    item_number = 1
    for row in rows_to_keep[1:]:
        if row[1] != "":
            row[0] = item_number 
            item_number += 1
        else:
            row[0] = ""
    
    # Clear the entire range
    sheet.batch_clear([data_range])
    
    # Write back only the rows to keep
    if rows_to_keep:
        sheet.update(range_name=data_range, values=rows_to_keep)

if __name__ == "__main__":
    main()