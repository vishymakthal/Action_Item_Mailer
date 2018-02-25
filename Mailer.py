from __future__ import print_function
import httplib2
import os
from email import *
from email import errors
import base64
import requests

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEImage import MIMEImage


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/gmail-python-quickstart.json
SCOPES = 'https://mail.google.com/'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Action Item Mailer'

SHEETS_SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
SHEETS_CLIENT_SECRET_FILE = 'client_secret_sheets.json'


def get_sheets_credentials():
    """Gets valid user credentials from storage. Same stuff as function below,
        but for Google Sheets.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """

    credential_dir = '.\.credentials'
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'mailer-sheets-credentials.json')


    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(SHEETS_CLIENT_SECRET_FILE, SHEETS_SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def get_credentials():
    """

    Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """

    credential_dir = '.\.credentials'
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'mailer-credentials.json')


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


def create_html_message(sender, to, subject, message_text,images=[]):
    """

    Create an html message for an email.

    Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    images: List of image src paths to open and attach.

    Returns:
    An object containing a base64url encoded email object.
    """

    msgRoot = MIMEMultipart('related')
    msgRoot['subject'] = subject
    msgRoot['from'] = sender
    msgRoot['to'] = to

    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)

    msgText = MIMEText(message_text, 'html')
    msgAlternative.attach(msgText)


    for imagePath in images:
        fp = open(imagePath, 'rb')
        msgImage = MIMEImage(fp.read())
        fp.close()

        # Define the image's ID as referenced above
        msgImage.add_header('Content-ID', '<'+imagePath.split('.')[0]+'>')
        msgRoot.attach(msgImage)

    return {'raw': base64.urlsafe_b64encode(msgRoot.as_string())}

def send_message(service, user_id, message):
    """Send an email message.

    Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

    Returns:
    Sent Message.
    """
    try:
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        print(message['id'])
        return message
    except Exception, e:
        print('An error occurred: ' + str(e))


def main():
    """Shows basic usage of the Gmail API.

    Creates a Gmail API service object and outputs a list of label names
    of the user's Gmail account.
    """
    credentials = get_credentials()
    sheets_credentials = get_sheets_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('gmail', 'v1', http=http)
    sheets_service = discovery.build('sheets','v4',http=sheets_credentials.authorize(httplib2.Http()))

    ieee_email_address = "ncsu.ieee@gmail.com"

    #Change this line if using a new spreadsheet
    spreadsheet_id = '1z0hHNCdUFggiGwBe_qYFMw63cE1-UYCGDTwXVfzeZw0'

    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range='A1:E').execute()

    action_items = result['values']

    for line in action_items:

        officer_and_tasks = line

        msg = '''

        <html>
            <b> Your Action Items </b>
              <hr/>
                <ul>
        '''
        #The start of the email to the current officer

        task_counter = 1 #initialize a counter for the number of tasks, serves as the index

        print(officer_and_tasks)

        officer_email = officer_and_tasks[0]

        while task_counter < len(officer_and_tasks): #go through the officer's tasks
            msg += '<li>'
            msg += officer_and_tasks[task_counter]
            msg += '</li>'
            task_counter += 1

        msg += '''

                </ul>
            <hr/>
            <p>Thanks for your commitment to NC State IEEE!</p>

        </html>

        '''

        send_message(service,ieee_email_address,create_html_message(ieee_email_address,officer_email,"IEEE Action Items (Testing)",msg))

if __name__ == '__main__':
    main()
