import os
import pickle
from bs4 import BeautifulSoup
# Gmail API utils
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
# for encoding/decoding messages in base64
from base64 import urlsafe_b64decode, urlsafe_b64encode
# for dealing with attachement MIME types
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from mimetypes import guess_type as guess_mime_type
from time import sleep

SCOPES = ['https://mail.google.com/']

# A variable that the program need to not note if there's a file twice
noted = False

def gmail_authenticate():
    creds = None
    # the file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)
    # if there are no (valid) credentials availablle, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # save the credentials for the next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)


# utility functions
def get_size_format(b, factor=1024, suffix="B"):
    """
    Scale bytes to its proper byte format
    e.g:
        1253656 => '1.20MB'
        1253656678 => '1.17GB'
    """
    for unit in ["", "K", "M", "G", "T", "P", "E", "Z"]:
        if b < factor:
            return f"{b:.2f}{unit}{suffix}"
        b /= factor
    return f"{b:.2f}Y{suffix}"


def clean(text):
    # clean text for creating a folder
    return "".join(c if c.isalnum() else "_" for c in text)


def search_messages(service, query):
    result = service.users().messages().list(userId='me',q=query).execute()
    messages = [ ]
    if 'messages' in result:
        messages.extend(result['messages'])
    while 'nextPageToken' in result:
        page_token = result['nextPageToken']
        result = service.users().messages().list(userId='me',q=query, pageToken=page_token).execute()
        if 'messages' in result:
            messages.extend(result['messages'])
    return messages


def parse_parts(service, parts, file_name, message):
    """
    Utility function that parses the content of an email partition and save it on a file
    """
    if parts:
        for part in parts:
            filename = part.get("filename")
            mimeType = part.get("mimeType")
            body = part.get("body")
            data = body.get("data")
            file_size = body.get("size")
            part_headers = part.get("headers")
            if part.get("parts"):
                # recursively call this function when we see that a part
                # has parts inside
                parse_parts(service, part.get("parts"), file_name, message)
            if mimeType == "text/plain":
                # it had to print the text, but i just need to save it
                pass
            elif mimeType == "text/html":
                # if the email part is an HTML content
                # convert HTML to txt and save it in the file
                old_path = os.path.join("Emails", "raw.txt")
                filepath = os.path.join("Emails", f"{file_name}.txt")
                os.rename(old_path, filepath)

                text = urlsafe_b64decode(data).decode()
                soup = BeautifulSoup(text, 'html.parser')
                with open(filepath, "a") as f:
                    f.write(soup.get_text())
            else:
                # write at the end of the file a note if it's present one or more attachment
                filepath = os.path.join("Emails", f"{file_name}.txt")
                with open(filepath, "a") as f:
                    global noted
                    if not noted:
                        f.write("Nota: Sono presenti uno o piu' allegati")
                        noted = True

def read_message(service, message):
    """
    This function takes Gmail API `service` and the given `message_id` and does the following:
        - Downloads the content of the email
        - Marks the email as read
        - Downloads text/html content (if available) and saves in the "Emails" folder in txt format
        - Sign on the file if there's one or more attachments
    """
    msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()

    # it mark the email that has been downloaded as read
    service.users().messages().batchModify(
        userId='me',
        body={
            'ids': msg['id'],
            'removeLabelIds': ['UNREAD']
        }
    ).execute()

    global noted
    noted = False

    # parts can be the message body, or attachments
    payload = msg['payload']
    headers = payload.get("headers")
    parts = payload.get("parts")

    if headers:
        # save the email's content in a file
        for header in headers:
            name = header.get("name")
            value = header.get("value")

            if name.lower() == 'from':
                # create the txt file and save from whom the email is
                filepath = os.path.join("Emails", f"raw.txt")
                with open(filepath, "x") as f:
                    f.write(f"Da: {value} \n")

            if name.lower() == "to":
                # The emails are usually sent to me so idc about it
                pass

            if name.lower() == "subject":
                # save the new name of the file based on the subject
                # if there's no subject, name it Subject_missing
                if clean(value) == '':
                    file_name = "Subject_missing.txt"
                else:
                    file_name = clean(value)
                # we will also handle emails with the same subject name
                file_counter = 0
                filepath = os.path.join("Emails", f"{file_name}.txt")
                while os.path.isfile(filepath):
                    file_counter += 1
                    # we have the same file name, add a number next to it
                    if file_name[-1].isdigit() and file_name[-2] == "_":
                        file_name = f"{file_name[:-2]}_{file_counter}"
                    elif file_name[-2:].isdigit() and file_name[-3] == "_":
                        file_name = f"{file_name[:-3]}_{file_counter}"
                    else:
                        file_name = f"{file_name}_{file_counter}"
                    filepath = os.path.join("Emails", f"{file_name}.txt")

            if name.lower() == "date":
                # I don't care about the date
                pass
    parse_parts(service, parts, file_name, message)


def main():
    # get the Gmail API service
    service = gmail_authenticate()

    # get emails that match the query you specify
    results = search_messages(service, "from:federico@geostefani.net is:unread")
    print(f"Found {len(results)} results.")

    # for each email matched, read it (output plain/text to console & save HTML and attachments)
    for msg in results:
        read_message(service, msg)


if __name__ == '__main__':
    while(True):
      main()
      sleep(3600)
