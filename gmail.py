from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient import errors
from pprint import pprint
import os,sys
import base64

from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://mail.google.com/']

def main():
    
    create_message_with_attachment('me','krupod@gmail.com', 'In', 'body', 'files/test-attachment.xls')
    
    load_unread('files')

def create_message_with_attachment(sender, to, subject, message_text, file):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file: The path to the file to be attached.

  """
  message = MIMEMultipart()
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject

  msg = MIMEText(message_text)
  message.attach(msg)

  part = MIMEBase('application', "vnd.ms-excel")
  part.set_payload(open(file, "rb").read())
  encoders.encode_base64(part)
  part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
  message.attach(part)

  b64_bytes = base64.urlsafe_b64encode(message.as_bytes())
  b64_string = b64_bytes.decode()
  body = {'raw': b64_string}

  build_service().users().messages().send(userId='me', body=body).execute()

def load_unread(folder, whitelist, label_name, markasread = False):
    """ Loads excel attachments from unread messages in associated gmail account.

    Args:
      folder: Location to store files.
      whitelist: Array of emails of interest.
    """
    service = build_service()

    labels = [ label for label in service.users().labels().list(userId = 'me').execute()['labels'] if label['name'] == label_name]

    if len(labels) == 0:
      print('Creating label %s' % label_name)
      label = service.users().labels().create(userId='me', body={'name' : label_name}).execute()
    else:
      label = labels[0]
  
    print('Processing for label %s, id %s' % (label_name, label['id']))

    os.makedirs(folder, exist_ok = True)
    messages = list_query_messages(service, 'me', query='label:unread label:inbox has:attachment NOT label:' + label_name)

    for m in messages:
        message = service.users().messages().get(userId='me', id=m.get('id')).execute()
        headers = message['payload']['headers']
        sender = [i['value'] for i in headers if i['name'].lower() == 'from'][0]
        if sender in whitelist:
          print('Saving attachment from %s: [%s]' % (sender, message['snippet']))
          save_message_attachments(service, message, folder)
          print('Applying label %s' % label_name)
          service.users().messages().modify(userId='me', id=m.get('id'), body={'addLabelIds': [label['id']]}).execute()
          if markasread:
            print('Marking as read')
            service.users().messages().modify(userId='me', id=m.get('id'), body={'removeLabelIds': ['UNREAD']}).execute()


def list_labels():
    """Lists labels.
    """

    service = build_service()

    # Call the Gmail API
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])

    if not labels:
        print('No labels found.')
    else:
        print('Labels:')
        for label in labels:
            print(label['name'])

def build_service():
    """Build Gmail API service.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    return service

def save_message_attachments(service, message, store_dir):
    """Get and store attachment from Message with given id.

    Args:
      message: message object
      store_dir: The directory used to store attachments.
    """
    for part in message['payload']['parts']:
        if part['mimeType'] == 'application/vnd.ms-excel':
            print('Saving ' + part['filename'] + ' from ' + message['snippet'])

            attachment = service.users().messages().attachments().get(userId='me', messageId=message['id'], id=part['body']['attachmentId']).execute()
            file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
            path = '/'.join([store_dir, part['filename']])

            f = open(path, 'wb')
            f.write(file_data)
            f.close()

def list_query_messages(service, user_id, query=''):
  """List all Messages of the user's mailbox matching the query.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    query: String used to filter messages returned.
    Eg.- 'from:user@some_domain.com' for Messages from a particular sender.

  Returns:
    List of Messages that match the criteria of the query. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate ID to get the details of a Message.
  """
  try:
    response = service.users().messages().list(userId=user_id,
                                               q=query).execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id, q=query,
                                         pageToken=page_token).execute()
      messages.extend(response['messages'])

    return messages
  except error:
    print ('An error occurred: %s' % error)


def ListMessagesWithLabels(service, user_id, label_ids=[]):
  """List all Messages of the user's mailbox with label_ids applied.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    label_ids: Only return Messages with these labelIds applied.

  Returns:
    List of Messages that have all required Labels applied. Note that the
    returned list contains Message IDs, you must use get with the
    appropriate id to get the details of a Message.
  """
  try:
    response = service.users().messages().list(userId=user_id,
                                               labelIds=label_ids).execute()
    messages = []
    if 'messages' in response:
      messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=user_id,
                                                 labelIds=label_ids,
                                                 pageToken=page_token).execute()
      messages.extend(response['messages'])

    return messages
  except error:
    print ('An error occurred: %s' % error)


if __name__ == '__main__':
    main()