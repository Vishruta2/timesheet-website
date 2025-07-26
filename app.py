from flask import Flask, jsonify
from google.oauth2 import service_account
from googleapiclient.discovery import build

app = Flask(__name__)

SHEET_ID = 'your_google_sheet_id_here'  # Replace with your actual Google Sheet ID
SHEET_RANGE = 'Form Responses 1'        # Or 'Sheet1!A1:H'
SERVICE_ACCOUNT_FILE = 'credentials.json'

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

@app.route('/timesheet', methods=['GET'])
def get_timesheet():
    try:
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(
            spreadsheetId=SHEET_ID,
            range=SHEET_RANGE
        ).execute()
        rows = result.get('values', [])
        if not rows:
            return 'No data found.', 404
        return jsonify({'data': rows})
    except Exception as e:
        print('Error:', e)
        return 'Failed to fetch data from Google Sheets', 500

if __name__ == '__main__':
    app.run(port=3000, debug=True)
