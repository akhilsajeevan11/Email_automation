from imapclient import IMAPClient
import email
import os
from datetime import datetime
from email.header import decode_header
from dotenv import load_dotenv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import html2text
from dataclasses import dataclass
from typing import List, Optional
import time

# Load environment variables
load_dotenv()

# Outlook Email configuration
EMAIL = os.getenv('Outlook_EMAIL')  # your_outlook_email@outlook.com
PASSWORD = os.getenv('Outlook_PASSWORD')
SMTP_SERVER = "smtp-mail.outlook.com"  # Outlook SMTP server
SMTP_PORT = 587
IMAP_SERVER = "outlook.office365.com"  # Outlook IMAP server

@dataclass
class EmailContent:
    subject: str
    sender: str
    date: str
    text_content: str
    html_content: str
    attachments: List[str]

    def get_plain_text(self) -> str:
        if self.text_content:
            return self.text_content
        if self.html_content:
            h = html2text.HTML2Text()
            h.ignore_links = False
            return h.handle(self.html_content)
        return ""

# def send_email(to_email: str, subject: str, body: str, attachments: List[str] = None):
#     """
#     Send an email with optional attachments
#     """
#     try:
#         # Create message
#         msg = MIMEMultipart()
#         msg['From'] = EMAIL
#         msg['To'] = to_email
#         msg['Subject'] = subject

#         # Add body
#         msg.attach(MIMEText(body, 'plain'))

#         # Add attachments
#         if attachments:
#             for file_path in attachments:
#                 with open(file_path, 'rb') as f:
#                     part = MIMEApplication(f.read(), Name=os.path.basename(file_path))
#                     part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
#                     msg.attach(part)

#         # Send email
#         with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
#             server.starttls()
#             server.login(EMAIL, PASSWORD)
#             server.send_message(msg)
        
#         print(f"Email sent successfully to {to_email}")
#         return True

#     except Exception as e:
#         print(f"Error sending email: {str(e)}")
#         return False

def monitor_email():
    try:
        with IMAPClient(IMAP_SERVER) as client:
            client.login(EMAIL, PASSWORD)
            print("Successfully logged in")
            
            client.select_folder('INBOX')
            print("Selected INBOX")
            
            initial_messages = set(client.search(['ALL']))
            print("Monitoring for new emails...")
            
            idle_timeout = 60
            while True:
                try:
                    client.idle()
                    responses = client.idle_check(timeout=idle_timeout)
                    
                    if responses:
                        client.idle_done()
                        current_messages = set(client.search(['ALL']))
                        new_messages = current_messages - initial_messages
                        
                        if new_messages:
                            print(f"\nNew email(s) detected! Count: {len(new_messages)}")
                            for msg_id in new_messages:
                                data = client.fetch([msg_id], 'RFC822')[msg_id]
                                email_message = email.message_from_bytes(data[b'RFC822'])
                                email_content = extract_email_content(email_message)
                                
                                print(f"\nNew Email:")
                                print(f"From: {email_content.sender}")
                                print(f"Subject: {email_content.subject}")
                                print(f"Content: {email_content.get_plain_text()[:200]}...")
                                
                                # Save attachments if present
                                if email_content.attachments:
                                    save_attachments(data[b'RFC822'])
                            
                            initial_messages = current_messages
                    else:
                        client.idle_done()
                        
                except KeyboardInterrupt:
                    print("\nGracefully shutting down...")
                    client.idle_done()
                    break
                    
    except Exception as e:
        print(f"Error monitoring emails: {str(e)}")

def process_new_email(client):
    """
    Process new unread emails with improved logging
    """
    try:
        messages = client.search(['UNSEEN'])
        print(f"Processing {len(messages)} new messages")
        
        for msg_id, data in client.fetch(messages, 'RFC822').items():
            try:
                email_message = email.message_from_bytes(data[b'RFC822'])
                email_content = extract_email_content(email_message)
                
                # Process the content
                print(f"\nNew Email Received:")
                print(f"From: {email_content.sender}")
                print(f"Subject: {email_content.subject}")
                print(f"Date: {email_content.date}")
                print("\nContent:")
                print(email_content.get_plain_text())
                print("\nAttachments:", email_content.attachments)
                
                # Save attachments
                if email_content.attachments:
                    save_attachments(data[b'RFC822'])
                    
            except Exception as e:
                print(f"Error processing message {msg_id}: {str(e)}")
                continue
                
    except Exception as e:
        print(f"Error in process_new_email: {str(e)}")

def extract_email_content(email_message) -> EmailContent:
    """
    Extract content from email message
    """
    subject = decode_header(email_message["Subject"])[0][0]
    if isinstance(subject, bytes):
        subject = subject.decode()
        
    from_ = decode_header(email_message.get("From", ""))[0][0]
    if isinstance(from_, bytes):
        from_ = from_.decode()
        
    date_ = email_message.get("Date", "")
    
    body_text = ""
    body_html = ""
    attachments = []
    
    if email_message.is_multipart():
        for part in email_message.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            
            if "attachment" in content_disposition:
                filename = part.get_filename()
                if filename:
                    attachments.append(filename)
                continue
                
            try:
                body = part.get_payload(decode=True).decode()
            except:
                continue
                
            if content_type == "text/plain":
                body_text += body
            elif content_type == "text/html":
                body_html += body
    else:
        content_type = email_message.get_content_type()
        try:
            body = email_message.get_payload(decode=True).decode()
            if content_type == "text/plain":
                body_text = body
            elif content_type == "text/html":
                body_html = body
        except:
            pass
            
    return EmailContent(
        subject=subject,
        sender=from_,
        date=date_,
        text_content=body_text,
        html_content=body_html,
        attachments=attachments
    )

def save_attachments(email_data, save_dir="attachments"):
    """
    Save email attachments to directory with support for various file types
    """
    # Supported file types
    ALLOWED_EXTENSIONS = {
        # Images
        '.jpg', '.jpeg', '.png', '.gif', '.bmp',
        # Documents
        '.pdf', '.doc', '.docx', '.xls', '.xlsx',
        '.ppt', '.pptx', '.txt',
        # Other
        '.zip', '.rar'
    }

    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    msg = email.message_from_bytes(email_data)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        
        if part.get('Content-Disposition') is None:
            continue
            
        filename = part.get_filename()
        if filename:
            # Decode filename if needed
            filename_parts = decode_header(filename)[0]
            if isinstance(filename_parts[0], bytes):
                filename = filename_parts[0].decode(filename_parts[1] or 'utf-8')
            
            # Check file extension
            file_ext = os.path.splitext(filename)[1].lower()
            if file_ext not in ALLOWED_EXTENSIONS:
                print(f"Skipping unsupported file type: {filename}")
                continue
            
            # Create unique filename with timestamp
            safe_filename = f"{timestamp}_{filename}"
            filepath = os.path.join(save_dir, safe_filename)
            
            try:
                # Save the attachment
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                print(f"Saved attachment: {safe_filename}")
                
                # Get and print file size
                file_size = os.path.getsize(filepath)
                print(f"File size: {file_size/1024:.2f} KB")
                
            except Exception as e:
                print(f"Error saving attachment {filename}: {str(e)}")

if __name__ == "__main__":
    monitor_email()


    # Example usage
    # choice = input("Choose action (1: Monitor emails, 2: Send email): ")
    
    # if choice == "1":
        # monitor_email()
    # elif choice == "2":
    #     to_email = input("Enter recipient email: ")
    #     subject = input("Enter subject: ")
    #     body = input("Enter message body: ")
    #     attach = input("Enter attachment path (or press enter to skip): ")
        
    #     attachments = [attach] if attach else None
    #     send_email(to_email, subject, body, attachments)
    # else:
    #     print("Invalid choice")