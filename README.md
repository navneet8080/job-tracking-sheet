# job-tracking-sheet
Job tracking sheet creater using python with advanced options.

For Excel Output:
Simply run the script and choose option 1

It will create an Excel file with all the formatting you requested

For Google Sheets Output:
You'll need to set up Google API credentials:

Go to the Google Cloud Console

Create a new project

Enable the Google Sheets API and Google Drive API

Create a service account and download the JSON credentials file

Run the script and choose option 2

Provide the path to your credentials file

The script will create a new Google Sheet with all the formatting

Features Included:
All column headers as specified

Bold headers

Appropriate column widths

Dropdown for Application Status

Conditional formatting for different statuses

Frozen header row

Support for both Excel and Google Sheets

Requirements:
Python 3.x

Required packages: pandas, openpyxl, gspread, google-auth

Install with: pip install pandas openpyxl gspread google-auth

Let me know if you'd like me to modify any part of this script or if you need help with the Google API setup process!