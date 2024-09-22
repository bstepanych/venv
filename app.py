from flask import Flask, request, render_template, jsonify
from googleapiclient.discovery import build
from google.oauth2 import service_account
import os

app = Flask(__name__)

# Google Docs API setuping
SCOPES = ['https://www.googleapis.com/auth/documents','https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'D:/DocTemplate/venv/forward-ace-436213-m6-de170b2f07f7.json'

TEMPLATE_DOCUMENT_ID = '1-0j_tIQR1cJoT4aMczqhROnxhWv3XUZ-uyno9HaHt_c'

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('docs', 'v1', credentials=credentials)

@app.route('/')
def form():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    data = request.form.to_dict()
    document_id = create_document_from_template(data)
    print(document_id)
    save_document_locally(document_id)
    return jsonify({'status': 'success'})

def create_document_from_template(data):
    # Copy the template document
    copy_title = 'Generated Document'
    body = {
        'name': copy_title
    }
    drive_service = build('drive', 'v3', credentials=credentials)
    copied_file = drive_service.files().copy(
        fileId=TEMPLATE_DOCUMENT_ID, body=body).execute()
    document_id = copied_file.get('id')
    
    # Grant edit permissions to the specified email
    permission = {
        'type': 'user',
        'role': 'writer',
        'emailAddress': 'bogdan.sak.st@gmail.com'
    }
    drive_service.permissions().create(
        fileId=document_id,
        body=permission,
        fields='id'
    ).execute()

    # Replace placeholders in the copied document
    requests = [
        {'replaceAllText': {
            'containsText': {'text': '{{name}}', 'matchCase': True},
            'replaceText': data['name']
        }},
        {'replaceAllText': {
            'containsText': {'text': '{{lastname}}', 'matchCase': True},
            'replaceText': data['lastname']
        }},
        {'replaceAllText': {
            'containsText': {'text': '{{taxcode}}', 'matchCase': True},
            'replaceText': data['taxcode']
        }},
        {'replaceAllText': {
            'containsText': {'text': '{{date}}', 'matchCase': True},
            'replaceText': data['date']
        }},
        {'replaceAllText': {
            'containsText': {'text': '{{address}}', 'matchCase': True},
            'replaceText': data['address']
        }},
        {'replaceAllText': {
            'containsText': {'text': '{{phone}}', 'matchCase': True},
            'replaceText': data['phone']
        }},
        {'replaceAllText': {
            'containsText': {'text': '{{position}}', 'matchCase': True},
            'replaceText': data['position']
        }},
        {'replaceAllText': {
            'containsText': {'text': '{{email}}', 'matchCase': True},
            'replaceText': data['email']
        }}
    ]

    result = service.documents().batchUpdate(
        documentId=document_id, body={'requests': requests}).execute()
    
    return document_id

def save_document_locally(document_id):
    document = service.documents().get(documentId=document_id).execute()
    content = document.get('body').get('content')
    with open('generated_document.txt', 'w', encoding='utf-8') as f:
        for element in content:
            if 'paragraph' in element:
                for text_run in element['paragraph']['elements']:
                    f.write(text_run['textRun']['content'])

if __name__ == '__main__':
    app.run(debug=True)