from __future__ import print_function
import ConfigParser
import httplib2
import io
import os
import re
import time
import unicodedata

from apiclient import discovery
from apiclient import errors
from apiclient.http import MediaFileUpload, MediaIoBaseDownload

from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import base64
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes



try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = ['https://www.googleapis.com/auth/gmail.compose',
          'https://www.googleapis.com/auth/drive']
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'EFGB'

def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    credential_dir = '.credentials'
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'efgb.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def main():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    gmail_service = discovery.build('gmail', 'v1', http=http)
    drive_service = discovery.build('drive', 'v3', http=http)
    sheets_service = discovery.build('sheets', 'v4', http=http)

    # Config
    Config = ConfigParser.ConfigParser()
    Config.read('send_reports_config.txt')
    reports_folder = Config.get('Reports', 'ReportsFolder')
    reports_pdf_folder = Config.get('Reports', 'ReportsPDFFolder')
    students_filename = Config.get('Reports', 'StudentsFilename')
    email_from = Config.get('Reports', 'EmailFrom')
    email_body_filename = Config.get('Reports', 'EmailBodyFilename')
    reports_name_string_to_strip = Config.get('Reports',
        'ReportsNameStringToStrip')

    try:
        print('Retrieving students list from %s...' % (students_filename))
        spreadsheet_id = get_file_id(drive_service, students_filename)
        if spreadsheet_id is None:
            print('Cannot find students file')
            return
        students = get_students(sheets_service, spreadsheet_id)
        print('Retrieving reports list...')
        reports_folder_id = get_file_id(drive_service, reports_folder)
        if reports_folder_id is None:
            print('Cannot find reports folder')
            return
        reports = get_reports(drive_service, reports_folder_id)
        if len(reports) != len(students.keys()):
            print('Length reports != students')
            return
        students = merge_students_with_reports(students, reports,
            reports_name_string_to_strip)
        reports_pdf_folder_id = get_file_id(drive_service, reports_pdf_folder)
        if reports_pdf_folder_id is None:
            print('Cannot find reports PDF folder')
            return
        return
        print('Exporting PDFs...')
        export_pdfs(drive_service, reports, reports_pdf_folder_id)
        print('Sending emails...')
        f = open(email_body_filename)
        email_body = f.read()
        f.close()
        send_emails(gmail_service, students, email_from, email_body)
    except Exception as e:
        print(str(e))


def send_emails(service, students, email_from, email_body):
    emails_sent = []
    if os.path.exists('emails_sent.txt'):
        f = open('emails_sent.txt', 'r')
        emails_sent = [x.strip('\n') for x in f.readlines()]
        f.close()
    f = open('emails_sent.txt', 'w')
    for id, info in students.items():
        if info['name'] in emails_sent:
            continue
        email_subject = 'Livret de competence / Progress report: %s' % (
            info['name'])
        email_to = info['email1']
        if info['email2']:
            email_to += ',' + info['email2']
        msg = create_message_with_attachment(email_from,
            email_to, email_subject, email_body,
            info['report_name'] + '.pdf')
        print('Sending ' + info['name'] + ' file to: ' + email_to)
        for i in range(3):
            try:
                send_message(service, 'me', msg)
                break
            except Exception as e:
                if i == 2:
                    raise Exception('Sending failed: ' + str(e))
                    f.close()
                time.sleep(5)
        os.remove(info['report_name'] + '.pdf')
        f.write(info['name'] + '\n')
        time.sleep(5)

    f.close()

def get_file_id(service, name):
    page_token = None
    while True:
        response = service.files().list(q="name = '" + name + "'",
                                        spaces='drive',
                                        fields='nextPageToken, files(id, name)',
                                        pageToken=page_token).execute()
        for file in response.get('files', []):
            return file.get('id')
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break


def get_reports(service, folder_id):
    reports = []
    page_token = None
    while True:
        response = service.files().list(q="'" + folder_id + "' in parents",
                                        spaces='drive',
                                        fields='nextPageToken, files(id, name)',
                                        pageToken=page_token).execute()
        reports.extend([(file.get('id'),
            file.get('name')) for file in response.get('files', [])])
        page_token = response.get('nextPageToken', None)
        if page_token is None:
            break

    return reports


def get_students(service, spreadsheet_id):
    students = {}
    range_name = 'Sheet1!A6:C100'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=range_name).execute()
    for i, value in enumerate(result['values']):
        name = value[0]
        for s in ('PS', 'MS', 'GS', 'CP'):
            pos = name.find(s)
            if pos != -1:
                break
        if pos != -1:
            name = value[0][:pos].strip()
        students[i] = {
            'name': name,
            'email1': value[1],
            'report_id': None,
            'report_name': None}
        emails = [value[1],]
        if len(value) > 2:
            students[i]['email2'] = value[2]
            emails.append(value[2])
        else:
            students[i]['email2'] = None

        for email in emails:
            if not re.match(r"[^@]+@[^@]+\.[^@]+",email):
                raise Exception("Invalid email: " + email)

    return students


def merge_students_with_reports(students, reports, string_to_strip):
    for id, info in students.items():
        name = students[id]['name']
        name_parts = get_name_parts(name)
        for report in reports:
            pos = report[1].find('2016')
            report_name = report[1][:pos]
            report_name = report_name.replace('Copy of','')
            report_name = report_name.replace('Livret scolaire','')
            report_name_parts = get_name_parts(report_name)
            print(name_parts)
            print(report_name_parts)
            if name_parts == report_name_parts:
                students[id]['report_id'] = report[0]
                students[id]['report_name'] = report[1]
                break

    incomplete = False
    for id, info in students.items():
        if info['report_id'] is None:
            print('Cannot find report for %s' % info['name'])
            incomplete = True

    if incomplete:
        raise Exception('Merge incomplete.')

    return students


def get_name_parts(name):
    n = name.replace('-', ' ').replace('_', ' ')
    parts = [strip_accents(e.strip()).lower() for e in n.split()]
    return parts


def export_pdfs(service, reports, reports_pdf_folder_id):
    for report in reports:
        request = service.files().export_media(
            fileId=report[0], mimeType='application/pdf')
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()

        fh.seek(0)
        f = open(report[1] + '.pdf', 'w')
        f.write(fh.read())
        f.close()
        fh.close()

        file_metadata = {
            'name' : report[1],
            'parents': [reports_pdf_folder_id],
            'mimeType' : 'application/pdf'}
        media = MediaFileUpload(report[1] + '.pdf',
                                mimetype='application/pdf',
                                resumable=True)
        # file = service.files().create(
        #    body=file_metadata, media_body=media, fields='id').execute()


def send_message(service, user_id, message):
  try:
    # message = (service.users().messages().send(userId=user_id, body=message)
    #           .execute())
    return message
  except errors.HttpError, error:
    print('An error occurred: %s' % error)


def create_message(sender, to, subject, message_text):
  message = MIMEText(message_text)
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  return {'raw': base64.urlsafe_b64encode(message.as_string())}


def create_message_with_attachment(
    sender, to, subject, message_text, file):
    message = MIMEMultipart()
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject

    msg = MIMEText(message_text)
    message.attach(msg)

    content_type, encoding = mimetypes.guess_type(file)

    main_type, sub_type = content_type.split('/', 1)
    fp = open(file, 'rb')
    msg = MIMEBase(main_type, sub_type)
    msg.set_payload(fp.read())
    fp.close()

    filename = os.path.basename(file)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

    return {'raw': base64.urlsafe_b64encode(message.as_string())}


def strip_accents(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s)
        if unicodedata.category(c) != 'Mn')


if __name__ == '__main__':
    main()
