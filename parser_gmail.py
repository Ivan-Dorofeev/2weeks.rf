import base64
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


def get_gmail_service():
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credes.json', SCOPES)
                creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
        try:
            service = build('gmail', 'v1', credentials=creds)
            return service
        except HttpError as error:
            print(f'An error occurred: {error}')


def get_list_messages():
    service = get_gmail_service()
    messages = service.users().messages().list(userId='me', q='from:vladyxa@mail.ru has:attachment').execute()
    return messages.get('messages', [])[0]['id']


def get_email_content(message_id):
    print('message_id = ', message_id)
    service = get_gmail_service()
    attach = service.users().messages().get(userId='me', id=message_id).execute()
    for part in attach["payload"]["parts"]:
        file_name = part["filename"]
        if len(part["filename"]) > 3 and part["filename"][-3:] == 'pdf':
            print('attachmentId = ', part['body']['attachmentId'])
            attachment = service.users().messages().attachments().get(userId='me', messageId=message_id,
                                                                      id=part['body']['attachmentId']).execute()
            return attachment, file_name


def decoder(attachment: str):
    data = attachment['data']
    file_data = base64.urlsafe_b64decode(data.encode('UTF8'))
    return file_data


def write_file(file_data, file_name: str):
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    file_path = f'uploads/{file_name}'
    if not os.path.exists(file_path):
        with open(file_path, 'wb') as f:
            f.write(file_data)


if __name__ == '__main__':
    message_id = get_list_messages()
    attachment, file_name = get_email_content(message_id)
    file_data = decoder(attachment)
    write_file(file_data, file_name)
