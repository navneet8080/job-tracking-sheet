import pandas as pd
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
import gspread
from google.oauth2.service_account import Credentials
import os

def create_excel_job_tracker(output_file='Job_Tracker.xlsx'):
    """Create an Excel job tracker with all specified formatting"""
    
    # Define column headers
    headers = [
        '#', 'Company Name', 'Job Title', 'Job Location', 'Date Applied',
        'Job Posting Link', 'Resume Link', 'Cover Letter Link', 
        'Application Status', 'Follow-Up Date', 'Response Received?', 'Notes'
    ]
    
    # Create a DataFrame with headers
    df = pd.DataFrame(columns=headers)
    
    # Create Excel writer
    writer = pd.ExcelWriter(output_file, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Job Applications')
    
    # Get workbook and worksheet
    workbook = writer.book
    worksheet = writer.sheets['Job Applications']
    
    # Format headers
    header_font = Font(bold=True)
    for col_num, header in enumerate(headers, 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.font = header_font
    
    # Set column widths
    column_widths = {
        'A': 5,    # #
        'B': 25,   # Company Name
        'C': 25,   # Job Title
        'D': 20,   # Job Location
        'E': 15,   # Date Applied
        'F': 40,   # Job Posting Link
        'G': 40,   # Resume Link
        'H': 40,   # Cover Letter Link
        'I': 20,   # Application Status
        'J': 15,   # Follow-Up Date
        'K': 20,   # Response Received?
        'L': 40    # Notes
    }
    
    for col, width in column_widths.items():
        worksheet.column_dimensions[col].width = width
    
    # Add data validation for Application Status
    dv = DataValidation(type="list", formula1='"Applied,Interview Scheduled,Offer,Rejected,Followed Up,No Response,On Hold"', 
                       allow_blank=True)
    worksheet.add_data_validation(dv)
    dv.add(f'I2:I1048576')  # Apply to entire column except header
    
    # Define conditional formatting colors
    status_colors = {
        'Applied': 'ADD8E6',          # Light Blue
        'Interview Scheduled': '90EE90',  # Light Green
        'Offer': 'FFFF00',            # Gold/Yellow
        'Rejected': 'FF9999',          # Light Red
        'Followed Up': 'FFA500',       # Orange
        'No Response': 'D3D3D3',       # Light Gray
        'On Hold': 'E6E6FA'           # Lavender
    }
    
    # Apply conditional formatting
    for status, color in status_colors.items():
        fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        formula = f'$I2="{status}"'
        
        # Create a Rule object
        rule = FormulaRule(formula=[formula], fill=fill)
        
        # Apply to entire row (columns A-L)
        for col in range(1, 13):
            column_letter = get_column_letter(col)
            worksheet.conditional_formatting.add(
                f'{column_letter}2:{column_letter}1048576',
                rule
            )
    
    # Freeze top row
    worksheet.freeze_panes = 'A2'
    
    # Save the workbook
    workbook.save(output_file)
    print(f"Excel job tracker created successfully: {output_file}")

def create_google_sheets_job_tracker(creds_file=None, sheet_name='Job Tracker'):
    """Create a Google Sheets job tracker with all specified formatting"""
    
    if not creds_file or not os.path.exists(creds_file):
        print("Google Sheets creation requires a credentials file. Please provide the path to your service account JSON file.")
        return
    
    try:
        # Authenticate with Google Sheets API
        scopes = ['https://www.googleapis.com/auth/spreadsheets',
                 'https://www.googleapis.com/auth/drive']
        
        creds = Credentials.from_service_account_file(creds_file, scopes=scopes)
        client = gspread.authorize(creds)
        
        # Create a new spreadsheet
        spreadsheet = client.create(sheet_name)
        
        # Get the first worksheet
        worksheet = spreadsheet.get_worksheet(0)
        
        # Define column headers
        headers = [
            '#', 'Company Name', 'Job Title', 'Job Location', 'Date Applied',
            'Job Posting Link', 'Resume Link', 'Cover Letter Link', 
            'Application Status', 'Follow-Up Date', 'Response Received?', 'Notes'
        ]
        
        # Update the worksheet with headers
        worksheet.append_row(headers)
        
        # Format headers (bold)
        header_format = {
            "textFormat": {"bold": True}
        }
        worksheet.format('A1:L1', header_format)
        
        # Set column widths
        column_widths = {
            1: 50,    # #
            2: 150,   # Company Name
            3: 150,   # Job Title
            4: 120,   # Job Location
            5: 100,   # Date Applied
            6: 250,   # Job Posting Link
            7: 250,   # Resume Link
            8: 250,   # Cover Letter Link
            9: 120,   # Application Status
            10: 100,   # Follow-Up Date
            11: 150,   # Response Received?
            12: 250    # Notes
        }
        
        # Google Sheets uses pixel widths
        worksheet.resize_columns(column_widths)
        
        # Add data validation for Application Status
        validation_rule = {
            "condition": {
                "type": "ONE_OF_LIST",
                "values": ["Applied", "Interview Scheduled", "Offer", "Rejected", 
                          "Followed Up", "No Response", "On Hold"]
            },
            "strict": True,
            "showCustomUi": True
        }
        
        # Apply to column I (9th column)
        start_cell = gspread.utils.rowcol_to_a1(2, 9)
        end_cell = gspread.utils.rowcol_to_a1(1000, 9)
        range_notation = f"{start_cell}:{end_cell}"
        
        worksheet.set_data_validation(range_notation, validation_rule)
        
        # Define conditional formatting rules
        status_colors = {
            'Applied': {'red': 0.678, 'green': 0.847, 'blue': 0.902},          # Light Blue
            'Interview Scheduled': {'red': 0.565, 'green': 0.933, 'blue': 0.565},  # Light Green
            'Offer': {'red': 1.0, 'green': 1.0, 'blue': 0.0},            # Gold/Yellow
            'Rejected': {'red': 1.0, 'green': 0.6, 'blue': 0.6},          # Light Red
            'Followed Up': {'red': 1.0, 'green': 0.647, 'blue': 0.0},       # Orange
            'No Response': {'red': 0.827, 'green': 0.827, 'blue': 0.827},       # Light Gray
            'On Hold': {'red': 0.902, 'green': 0.902, 'blue': 0.980}           # Lavender
        }
        
        # Apply conditional formatting
        requests = []
        for status, color in status_colors.items():
            # Create a format rule for the entire row
            rule = {
                "ranges": [{
                    "sheetId": worksheet.id,
                    "startRowIndex": 1,  # Skip header row
                    "startColumnIndex": 0,
                    "endColumnIndex": 12  # Columns A-L
                }],
                "booleanRule": {
                    "condition": {
                        "type": "CUSTOM_FORMULA",
                        "values": [{"userEnteredValue": f'=$I2="{status}"'}]
                    },
                    "format": {
                        "backgroundColor": color
                    }
                }
            }
            requests.append({"addConditionalFormatRule": {
                "index": 0,
                "rule": rule
            }})
        
        # Freeze the top row
        freeze_request = {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": worksheet.id,
                    "gridProperties": {
                        "frozenRowCount": 1
                    }
                },
                "fields": "gridProperties.frozenRowCount"
            }
        }
        requests.append(freeze_request)
        
        # Batch update the spreadsheet
        spreadsheet.batch_update(requests)
        
        print(f"Google Sheets job tracker created successfully: {spreadsheet.url}")
        return spreadsheet.url
        
    except Exception as e:
        print(f"Error creating Google Sheets job tracker: {str(e)}")
        return None

def main():
    print("Job Tracker Spreadsheet Creator")
    print("1. Create Excel file (.xlsx)")
    print("2. Create Google Sheets (requires service account credentials)")
    
    choice = input("Enter your choice (1 or 2): ")
    
    if choice == '1':
        output_file = input("Enter output filename (default: Job_Tracker.xlsx): ") or 'Job_Tracker.xlsx'
        create_excel_job_tracker(output_file)
    elif choice == '2':
        creds_file = input("Enter path to your Google service account credentials JSON file: ")
        sheet_name = input("Enter name for your Google Sheet (default: Job Tracker): ") or 'Job Tracker'
        create_google_sheets_job_tracker(creds_file, sheet_name)
    else:
        print("Invalid choice. Please enter 1 or 2.")

if __name__ == "__main__":
    main()