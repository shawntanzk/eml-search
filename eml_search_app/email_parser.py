import os
import email
from email import policy
from email.parser import BytesParser
from eml_parser import eml_parser
import pandas as pd
from datetime import datetime

class EMLParser:
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.emails = []

    def parse_emails(self):
        """Parse all EML files in the folder and extract metadata."""
        for file in os.listdir(self.folder_path):
            if file.endswith('.eml'):
                file_path = os.path.join(self.folder_path, file)
                try:
                    with open(file_path, 'rb') as f:
                        msg = BytesParser(policy=policy.default).parse(f)
                    
                    # Extract basic info
                    subject = msg.get('subject', '')
                    sender = msg.get('from', '')
                    to = msg.get('to', '')
                    date_str = msg.get('date', '')
                    date = self._parse_date(date_str)
                    
                    # Extract body
                    body = self._extract_body(msg)
                    
                    # Use eml_parser for more details if needed
                    ep = eml_parser(file_path)
                    parsed_eml = ep.decode_email_bytes()
                    
                    email_data = {
                        'file': file,
                        'subject': subject,
                        'sender': sender,
                        'to': to,
                        'date': date,
                        'body': body,
                        'attachments': parsed_eml.get('attachments', []),
                        'headers': parsed_eml.get('header', {})
                    }
                    self.emails.append(email_data)
                except Exception as e:
                    print(f"Error parsing {file}: {e}")
        
        return pd.DataFrame(self.emails)

    def _parse_date(self, date_str):
        """Parse email date string to datetime."""
        try:
            return email.utils.parsedate_to_datetime(date_str)
        except:
            return None

    def _extract_body(self, msg):
        """Extract plain text body from email."""
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body += part.get_payload(decode=True).decode('utf-8', errors='ignore')
        else:
            body = msg.get_payload(decode=True).decode('utf-8', errors='ignore')
        return body