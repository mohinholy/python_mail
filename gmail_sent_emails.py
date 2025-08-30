from __future__ import print_function
import re
import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os.path
import pickle

# Gmail API scope (read-only access)
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def main():
    print("🔹 Starting Gmail Sent Mail Extractor...")

    creds = None
    if os.path.exists('token.pickle'):
        print("✅ Found existing login token...")
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    else:
        print("❌ No login token found, will ask for Google login...")

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("🔄 Refreshing expired token...")
            creds.refresh(Request())
        else:
            print("🌐 Opening browser for Google login...")
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    print("✅ Connected to Gmail API... Now fetching mails from Sent...")

    service = build('gmail', 'v1', credentials=creds)

    contacts = {}   # dictionary: {email: name}
    page_token = None
    total_msgs = 0
    page_count = 0

    while True:
        try:
            results = service.users().messages().list(
                userId='me', labelIds=['SENT'], maxResults=100, pageToken=page_token
            ).execute()
        except Exception as e:
            print("❌ Error while fetching message list:", e)
            break

        messages = results.get('messages', [])
        page_count += 1
        print(f"📨 Page {page_count}: Found {len(messages)} messages")

        for msg in messages:
            total_msgs += 1
            try:
                msg_data = service.users().messages().get(
                    userId='me', id=msg['id'], format="metadata",
                    metadataHeaders=['To', 'Cc', 'Bcc']
                ).execute()

                headers = msg_data['payload']['headers']
                for header in headers:
                    if header['name'] in ['To', 'Cc', 'Bcc']:
                        to_value = header['value']

                        # Match "Name <email>" OR just "email"
                        matches = re.findall(r'(?:"?([^"<]*)"?\s*)?<([^<>]+)>', to_value)
                        if matches:
                            for name, email in matches:
                                name = name.strip() if name else ""
                                contacts[email] = name
                        else:
                            found = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-z]{2,}", to_value)
                            for email in found:
                                contacts[email] = ""

            except Exception as e:
                print(f"⚠️ Skipping message {msg['id']} due to error:", e)

        page_token = results.get('nextPageToken')
        if not page_token:
            print("✅ No more pages left.")
            break

    print(f"✅ Processed {total_msgs} mails, found {len(contacts)} unique addresses")

    # Save results
    if contacts:
        df = pd.DataFrame(
            [{"Name": contacts[email], "Email": email} for email in sorted(contacts.keys())]
        )
        df.to_excel("unique_sent_emails.xlsx", index=False)
        print("📂 Saved results to unique_sent_emails.xlsx")
    else:
        print("⚠️ No emails found in Sent folder!")

if __name__ == '__main__':
    main()
