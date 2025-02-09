import imaplib
import email
from email.header import decode_header
import logging
import datetime
import csv

# Gmail IMAP settings
EMAIL_USER = "kyawzin.ccna@gmail.com"
EMAIL_PASS = "fqtf huup vcai zwhe"  # Use App Password
IMAP_SERVER = "imap.gmail.com"

# Keywords to search for
SPECIAL_WORDS = [
    "urgent", "meeting", "invoice", "refugees", "humanitarian",
    "asylum", "resettlement", "UNHCR", "Australia embassy",
    "visa application", "protection visa", "subclass 200", "subclass 201",
    "subclass 202", "subclass 203", "subclass 204","Amazon Web Services","Bualuang mBanking","Google","Global Special Humanitarian visa"
]

# CSV file path
csv_file_path = '/Users/mr.kyawzin/Desktop/output.csv'  # Replace with your desired file path

def search_gmail():
    try:
        # Connect to Gmail
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")
        
        # Calculate the date 7 days ago
        #date_since = (datetime.date.today() - datetime.timedelta(7)).strftime("%d-%b-%Y")


        # Search for today's emails
        today = datetime.date.today().strftime("%d-%b-%Y")
        status, email_ids = mail.search(None, f'(ON "{today}")')
        
        # Search for emails since 'date_since'
        #status, email_ids = mail.search(None, f'(SINCE "{date_since}")')
        
        email_ids = email_ids[0].split()

        for email_id in email_ids:
            # Fetch email data
            status, data = mail.fetch(email_id, "(RFC822)")
            raw_email = data[0][1]

            # Parse email content
            msg = email.message_from_bytes(raw_email)
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes) and encoding:
                subject = subject.decode(encoding)

            # Decode email body
            body = ""
       
            if msg.is_multipart():
               for part in msg.walk():
                   content_type = part.get_content_type()
                   content_disposition = str(part.get("Content-Disposition"))
                   if content_type == "text/plain" and "attachment" not in content_disposition:
                      charset = part.get_content_charset() or 'utf-8'
                      body = part.get_payload(decode=True).decode(charset, errors="ignore")
                      break
            else:
                charset = msg.get_content_charset() or 'utf-8'
                body = msg.get_payload(decode=True).decode(charset, errors="ignore")


            # Check for special words in subject or body
            if any(word.lower() in subject.lower() or word.lower() in body.lower() for word in SPECIAL_WORDS):
                print(f"\n🔍 Matched Email: {subject}")
                print(f"📩 From: {msg['From']}")
                print(f"📄 Body Preview: {body[:200]}...\n")  # Print first 200 characters
                
                # Write to CSV
                with open(csv_file_path, mode='a', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow([msg["Date"], msg["From"], subject, body[:200]])

        mail.logout()
    except Exception as e:
        print("Error:", e)
        

# Run the function
search_gmail()