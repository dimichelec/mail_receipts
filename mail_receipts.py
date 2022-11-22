
# Abstract:
#
# * look at the emails in my gmail inbox for any that are labeled "Receipts"
# * Add each as a checkbox line item in my Google doc "Receipts", and trash the
#   email
# + ToDo: Ask me which should be forwarded to my coupa account instead
# + ToDo: Read Receipts doc and delete lines that I've checked-off

import os
import re
from datetime import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.exceptions import RefreshError
from googleapiclient.errors import HttpError

from googleapiclient.discovery import build

from bs4 import BeautifulSoup
import base64
import json


DEBUG = False


def getCredsFromAuthFlow(SCOPES):
    # your creds file here. Please create json file as here:
    # https://cloud.google.com/docs/authentication/getting-started
    flow = InstalledAppFlow.from_client_secrets_file('credentials.json',
                                                     SCOPES)
    print()
    creds = flow.run_local_server(port=0)
    print()
    return creds


SCOPES = ['https://www.googleapis.com/auth/gmail.readonly',
          'https://www.googleapis.com/auth/gmail.modify',
          'https://www.googleapis.com/auth/documents',
          'https://www.googleapis.com/auth/drive']


def authenticate():
    """Authenticate this app for Google APIs use. Return credentials."""

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except RefreshError as e:
                print('\n!ERROR:', e.args[0])
                ok = input("Type YES if you want to delete the expired "
                           "token.json and try getting a new one: ")
                if ok == "YES":
                    os.remove('token.json')
                    creds = getCredsFromAuthFlow(SCOPES)
                else:
                    print()
                    exit()
        else:
            creds = getCredsFromAuthFlow(SCOPES)

        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return creds


def getEmailReceipts(creds):
    """Return the set of 'Receipts' emails from my Gmail inbox using the
       supplied credentials."""

    service = build('gmail', 'v1', credentials=creds)
    messages = service.users().messages().list(userId='me', labelIds=['INBOX'],
                                               q="in:Receipts"
                                               ).execute().get('messages', [])
    receipts = []
    for message in messages:
        msg = service.users().messages().get(userId='me',
                                             id=message['id']).execute()
        receipts.append(formatReceipt(msg))
    return receipts


def getEmailLines(data):
    out = ''
    text = base64.urlsafe_b64decode(data.encode('utf-8')).decode('utf-8')
    soup = BeautifulSoup(text, 'lxml')
    for s in soup.text.splitlines():
        s = s.strip()
        if s:
            out += s + '|'
    return out


def formatReceipt(msg):
    """Accept a receipt email and return it as a receipt data row."""

    # get what we can from the email header
    email_data = msg['payload']['headers']
    msgid = msg['id']
    for values in email_data:
        name = values['name']
        if name == 'From':
            from_name = values['value']
            from_domain = re.findall(r"(\@[^>\"]+)", from_name)[0]
            if DEBUG:
                print('\nfrom:', from_domain)
        if name == 'Date':
            date = datetime.strptime(values['value'][5:16],
                                     '%d %b %Y').strftime('%b %d %Y')
        if name == 'Subject':
            subject = values['value']

    # process the snippet and payload
    out = ''
    snippet = msg['snippet']
    payload = msg['payload']
    if payload:
        if payload['body']['size'] > 0:
            out += getEmailLines(payload['body']['data'])

        if 'parts' in payload.keys():
            for part in payload['parts']:
                if (not part['filename']):
                    out += getEmailLines(part['body']['data'])

    if DEBUG:
        print("[5:", out, "]")

    relist = [r"Total\s+(\$\S+)",
              r"Amount\:\s+(\S+)",
              r"Receipt for (\S+)",
              r"payment of (\S+)"]
    type = merchant = note = ''
    total = []
    for r in relist:
        total += re.findall(r, snippet)
    if total:
        total = max(total)
    if from_domain == '@uber.com':
        tmp = re.findall(r"Total\|[$0-9.]+(\w\w\w)\w* (\d+)\, (\d+).*"
                         r"Payments.(\w+ ••••\d\d\d\d)", out)[0]
        date = tmp[0] + ' ' + tmp[1] + ' ' + tmp[2]
        type = tmp[3].replace('••••', 'x')
        merchant = 'Uber'
    elif from_domain == '@paypal.com':
        tmp = re.findall(r'.*Transaction date(\w+ \d+, \d+).*'
                         r'your credit card statement as "([^\"]+).*'
                         r'Sources Used[^\|]+\|([^\|]+)\|([$0-9.]+)', out)[0]
        date = tmp[0].replace(',', '')
        merchant = tmp[1]
        type = tmp[2].replace('-', '')
        total = tmp[3]
    elif from_domain == '@well-net.org':
        date = re.findall(r"Date: (\S+ \S+)", snippet)[0] + ' ' \
               + str(datetime.now().year)
        type = re.findall(r"Payment method: (\S+ \S+)", snippet)[0]
        merchant = 'Tufts'
        if total[-1] == '.':
            total = total[:-1]
    elif from_domain == '@messaging.squareup.com':
        tmp = re.findall(r"You paid ([$0-9.]+) with your (.*) ending in (\d+)"
                         r" to ([\w ]+) on (\w+ \d+ \d+) at .*", out)[0]
        total = tmp[0]
        type = tmp[1] + ' x' + tmp[2]
        merchant = 'Square|' + tmp[3]
        date = tmp[4]
    elif from_domain == '@steampowered.com':
        tmp = re.findall(r"Date issued:\|(\w+ \d+, \d+).*"
                         r"Payment method:\|(\w+).*"
                         r"Total:\|([$0-9.]+)", out)[0]
        date = tmp[0].replace(',', '')
        type = tmp[1]
        total = tmp[2]
        merchant = 'Steam'
    elif from_domain == "@toasttab.com":
        tmp = re.findall(r".* visit to ([^\|]+)\|.*"
                         r"Ordered:\|([\d\/]+).*"
                         r"Total\|([$0-9.]+)\|.*\|([^\|]+)\|x+(x\d+)\|", out)
        if not tmp:
            tmp = re.findall(r"([^\|]+)\|Check.*"
                             r"Ordered:\|([\d\/]+).*"
                             r"Total\|([$0-9.]+)\|([^\|]+)\|x+(x\d+)\|", out)
        if tmp:
            tmp = tmp[0]
            merchant = 'Toast|' + tmp[0]
            date = datetime.strptime(tmp[1], '%m/%d/%y').strftime('%b %d %Y')
            total = tmp[2]
            type = tmp[3] + ' ' + tmp[4]
    elif from_domain == "@bluebikes.com":
        tmp = re.findall(r"Amount: ([$0-9.]+)", out)[0]
        type = 'x'
        total = tmp
        merchant = 'BlueBikes'
    elif from_domain == "@lyftmail.com":
        tmp = re.findall(r"Lyft(?:\|[^\|]+)+(?:\|([^\|]+)"
                         r" \*([0-9]+))"
                         r"\|([^\|]+)(?:\|[^\|]+)+", out)[0]
        type = tmp[0] + ' x' + tmp[1]
        total = tmp[2]
        merchant = 'Lyft'
    elif from_domain == "@parkmobileglobal.com":
        tmp = re.findall(r"Parkmobile(?:\|[^\|]+)+\|Payment Method\|(.*)"
                         r" ending in ([0-9]+)"
                         r".*Amount Paid\|([^\|]+)", out)[0]
        type = tmp[0] + ' x' + tmp[1]
        total = tmp[2]
        merchant = 'Parkmobile'
    else:
        merchant = from_domain
        note = from_name + ' | ' + subject
        if not total:
            note += ' | ' + snippet

    type = type.replace('American Express', 'Amex')
    if DEBUG:
        print(f'date: {date}, type: {type}, total: {total}, '
              f'merchant: {merchant}, note: {note}')
    return {'date': date, 'type': type, 'total': total, 'merchant': merchant,
            'note': note, 'id': msgid}


def createNewReceiptsDoc(creds, filename):
    """Create a new Receipts doc on the Google Drive using supplied
       credentials. Return new documentId."""

    # Remove the old docId json if it exists
    if os.path.exists('docId.json'):
        os.remove('docId.json')

    service = build('docs', 'v1', credentials=creds)
    docId = service.documents().create(body={'title': filename})\
                               .execute()['documentId']

    # Save the docId to uniquely identify our doc next time
    js = json.dumps({'docId': docId})
    with open('docId.json', 'w') as fout:
        fout.write(js)

    return docId


def findReceiptsDoc(creds):
    """Find and return the documentId of the Receipts document in my Google
       Drive using the supplied credentials. Create a new one if one is not
       found, or we didn't have its documentId already recorded locally in
       docId.json.
    """

    # find all the files named 'Receipts', in the Google Drive root
    filename = 'Receipts' if not DEBUG else 'Receipts-test'

    try:
        creds.refresh(Request())
    except RefreshError as e:
        print('\n!ERROR:', e.args[0])
        ok = input("Type YES if you want to delete the expired "
                   "token.json and try getting a new one: ")
        if ok == "YES":
            os.remove('token.json')
            creds = getCredsFromAuthFlow(SCOPES)
            print('Re-run program')
        print()
        exit()

    service = build('drive', 'v3', credentials=creds)
    files = service.files().list(q=f"name='{filename}' and trashed=false "
                                   f"and 'root' in parents", spaces='drive',
                                   fields='files(id)').execute()['files']

    # if we didn't find a Receipts file, make one
    if files:
        if os.path.exists('docId.json'):
            with open('docId.json', 'r') as fin:
                docId = json.load(fin)['docId']

            for f in files:
                if f['id'] == docId:
                    return docId

    # if we got here, we didn't find our file, create one
    docId = createNewReceiptsDoc(creds, filename)
    print(f"New {filename} doc created")

    return docId


def addReceipts(creds, docId, receipts):
    """Add the list of receipts to the Receipts doc found earlier using the
       supplied credentials.
    """

    # get Receipts doc and end index (nextIndex) where we'll start our new
    # receipt list
    service = build('docs', 'v1', credentials=creds)
    doc = service.documents().get(documentId=docId).execute()
    content = doc['body']['content']
    nextIndex = max([v['endIndex'] for v in content]) + 1

    # add receipt lines
    index = nextIndex
    requests = []
    for r in receipts:
        line = f"{r['date']}{', ' + r['total'] if r['total'] else ''}" \
               f"{', ' + r['type'] if r['type'] else ''}" \
               f", {r['merchant']}{', ' + r['note'] if r['note'] else ''}, " \
               "link\n"
        index += len(line)
        requests.append([
            {'insertText': {
                'text': line,
                'endOfSegmentLocation': {}
            }},
            {'updateTextStyle': {
                'textStyle': {
                    'link': {
                        'url':
                        'https://mail.google.com/mail/u/0/'
                        '?tab=rm&ogbl#inbox/' + r['id']
                    }
                },
                'range': {
                    'startIndex': index-7,
                    'endIndex': index-3
                },
                'fields': 'link'
            }}
        ])

    service.documents().batchUpdate(
        documentId=docId, body={'requests': requests}).execute()

    # get updated doc and new end of doc index
    doc = service.documents().get(documentId=docId).execute()
    content = doc['body']['content']
    # -1 in the next line to exclude a new bullet on last blank line
    lastIndex = max([v['endIndex'] for v in content]) - 1

    # turn our receipt list into a checklist
    requests = [
        {
            'createParagraphBullets': {
                'range': {
                    'startIndex': nextIndex,
                    'endIndex': lastIndex
                },
                'bulletPreset': 'BULLET_CHECKBOX'
            },
        },
        {
            'updateParagraphStyle': {
                'range': {
                    'startIndex': nextIndex,
                    'endIndex': lastIndex
                },
                'paragraphStyle': {
                    'indentFirstLine': {'unit': 'PT'},
                    'indentStart': {'magnitude': 18, 'unit': 'PT'}
                },
                'fields': 'indentFirstLine,indentStart'
            }
        }
    ]
    service.documents().batchUpdate(documentId=docId,
                                    body={'requests': requests}).execute()


def trashReceipts(creds, receipts):

    service = build('gmail', 'v1', credentials=creds)
    for r in receipts:
        service.users().messages().trash(userId='me', id=r['id']).execute()


NAMESPACE = os.path.splitext(os.path.basename(__file__))[0].strip()
print(f'\n\n{NAMESPACE}\n')

if DEBUG:
    print('*** DEBUG mode ***\n')

print('Getting creds')
creds = authenticate()

print('Finding Receipts doc', end='')
docId = findReceiptsDoc(creds)
print(f' [Id: {docId}]')

print('Reading emails', end='')
receipts = getEmailReceipts(creds)
print(f' [{len(receipts)} receipt(s) found]')

if len(receipts):
    print('Adding receipts')
    addReceipts(creds, docId, receipts)

    if not DEBUG:
        print('Trashing receipt emails')
        trashReceipts(creds, receipts)

print('')
