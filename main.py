from flask import Flask, request, redirect, url_for
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from werkzeug.utils import secure_filename
import os
import pdfplumber
from docx import Document
from dateutil import parser

app = Flask(__name__)
SCOPES = ['https://www.googleapis.com/auth/calendar']

@app.route('/')
def index():
    return '''
    <h1>イベント登録アプリへようこそ！</h1>
    <form method="post" action="/upload" enctype="multipart/form-data">
        <input type="file" name="file">
        <input type="submit" value="アップロード">
    </form>
    '''

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'ファイルがないっぽいよ！'

    file = request.files['file']
    if file.filename == '':
        return 'ファイル名がないよ！'

    filename = secure_filename(file.filename)
    file_path = os.path.join('uploads', filename)
    file.save(file_path)

    text = extract_text_from_file(file_path)
    event_info = extract_event_info(text)
    create_google_calendar_event(event_info['summary'], event_info['start'], event_info['end'])

    return 'Googleカレンダーにイベント登録完了！'

def extract_text_from_file(file_path):
    if file_path.endswith('.pdf'):
        with pdfplumber.open(file_path) as pdf:
            text = ''.join([page.extract_text() for page in pdf.pages])
    elif file_path.endswith('.docx'):
        doc = Document(file_path)
        text = ''.join([para.text for para in doc.paragraphs])
    else:
        text = '対応してないファイル形式だよ！'
    return text

def extract_event_info(text):
    try:
        date = parser.parse(text, fuzzy=True)
    except:
        date = None
    return {
        'summary': '抽出イベント',
        'start': date,
        'end': date
    }

def create_google_calendar_event(summary, start, end):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    service = build('calendar', 'v3', credentials=creds)

    event = {
        'summary': summary,
        'start': {
            'dateTime': start.isoformat(),
            'timeZone': 'Asia/Tokyo',
        },
        'end': {
            'dateTime': end.isoformat(),
            'timeZone': 'Asia/Tokyo',
        },
    }

    service.events().insert(calendarId='primary', body=event).execute()

if __name__ == '__main__':
    app.run(debug=True)
