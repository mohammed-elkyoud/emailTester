import imaplib
import email
from email.header import decode_header
import smtplib
import time
import os
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import google.generativeai as genai

# --- CONFIGURATION (REPLACE THESE VALUES) ---
EMAIL_ACCOUNT = ""  # Your Gmail address
PASSWORD = ""  # Your Gmail App Password (if 2FA is enabled)
GEMINI_API_KEY = ""  # Your Gemini API Key
# -------------------------------------------

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel("gemini-1.5-flash")  # or another model

REQUIRED_INFO = {
    "attestation de scolarite": ["Nom", "Prenom", "Annee d'inscription", "Numero de telephone", "Filliere"],
    "attestation de stage": ["Nom", "Prenom", "Annee de stage", "Numero de telephone", "Filliere"]
}

PDF_DIR = "./pdfs"
os.makedirs(PDF_DIR, exist_ok=True)

# Decode MIME header for email
def decode_mime_words(s):
    try:
        hdrs = decode_header(s)
        parts = []
        for text, encoding in hdrs:
            if isinstance(text, bytes):
                parts.append(text.decode(encoding or 'utf-8', errors='replace'))
            else:
                parts.append(text)
        return "".join(parts)
    except Exception as e:
        print(f"Error decoding MIME words: {e}")
        return str(s)

# Extract fields with Gemini API
def extract_fields_with_gemini(body, fields):
    extracted_data = {}
    for field in fields:
        prompt = f"Extract the {field} from the following email body:\n\n{body}\n\n{field}:"
        try:
            response = model.generate_content(prompt)
            extracted_value = response.text.strip() if response.text else None
            extracted_data[field] = extracted_value
        except Exception as e:
            print(f"Gemini API Error for {field}: {e}")
            extracted_data[field] = None
    return extracted_data

# Validate that required fields were extracted
def validate_extracted_data(extracted_data, required_fields):
    return [field for field in required_fields if not extracted_data.get(field)]

# Generate a PDF from the extracted details
def generate_pdf(subject, details):
    pdf_path = os.path.join(PDF_DIR, f"{subject.replace(' ', '_')}.pdf")
    c = canvas.Canvas(pdf_path, pagesize=letter)
    c.drawString(100, 750, f"Attestation: {subject}")
    c.drawString(100, 720, "Details Provided:")
    y_position = 700
    for key, value in details.items():
        c.drawString(100, y_position, f"{key}: {value}")
        y_position -= 20
    c.drawString(100, y_position - 20, "FSTG MARRAKECH")
    c.save()
    return pdf_path

# Send reply email
def send_reply(to_address, subject, body, attachment_path=None):
    try:
        # Check if this user has already received a reply
        if user_already_processed(to_address):
            print(f"Skipping email to {to_address} - already processed.")
            return  # Skip sending if already processed

        msg = MIMEMultipart()
        msg['From'] = EMAIL_ACCOUNT
        msg['To'] = to_address
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        if attachment_path:
            with open(attachment_path, 'rb') as attachment:
                part = MIMEBase('application', "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(attachment_path)}")
                msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            print("Connecting to SMTP server...")
            server.login(EMAIL_ACCOUNT, PASSWORD)
            server.sendmail(EMAIL_ACCOUNT, to_address, msg.as_string())

        # After sending the email, mark this user as processed
        mark_user_as_processed(to_address)
    except Exception as e:
        print(f"Error sending email: {e}")

# Check if the user has already received a reply
def user_already_processed(email_address):
    try:
        # Open the file containing processed users
        with open('processed_users.txt', 'r') as file:
            processed_users = file.readlines()
        # Strip any extra whitespace and check if the email is in the file
        return email_address.strip() + '\n' in processed_users
    except FileNotFoundError:
        # If the file doesn't exist, create it
        return False

# Mark the user as processed by saving to a file
def mark_user_as_processed(email_address):
    try:
        # Open the file in append mode and write the email address
        with open('processed_users.txt', 'a') as file:
            file.write(email_address.strip() + '\n')
    except Exception as e:
        print(f"Error writing to processed users file: {e}")

# Process the email
def process_email(msg):
    raw_subject = msg["Subject"]
    subject = decode_mime_words(raw_subject) if raw_subject else "(No Subject)"
    sender = decode_mime_words(msg["From"])
    body = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    body = part.get_payload(decode=True).decode('utf-8', errors='replace')
                except:
                    body = str(part.get_payload())  # if decoding fails use the raw payload as string
                break
    else:
        try:
            body = msg.get_payload(decode=True).decode('utf-8', errors='replace')
        except:
            body = str(msg.get_payload())

    print(f"Processing email from: {sender}\nSubject: {subject}\nBody preview: {body[:100]}...")

    matched_subject = next((key for key in REQUIRED_INFO if key in subject.lower()), None)
    if matched_subject:
        required_fields = REQUIRED_INFO[matched_subject]
        extracted_data = extract_fields_with_gemini(body, required_fields)
        missing_info = validate_extracted_data(extracted_data, required_fields)
        if missing_info:
            # Send request for missing info
            reply_subject = f"Re: Your Request - {matched_subject}"
            reply_body = f"""Dear {sender.split()[0] if ' ' in sender else sender},\n\nThank you for your request for an {matched_subject}.\n\nTo proceed, please provide the following information:\n{', '.join(missing_info)}\n\nOnce we have this information, we will process your request.\n\nSincerely,\nThe Yam Team"""
            send_reply(sender, reply_subject, reply_body)
        else:
            # Generate attestation PDF and send it
            pdf_path = generate_pdf(matched_subject, extracted_data)
            reply_subject = f"Re: {matched_subject}"
            reply_body = f"""Dear {sender.split()[0] if ' ' in sender else sender},\n\nThank you for your request for an {matched_subject}. Please find your attestation attached.\n\nSincerely,\nThe Yam Team"""
            send_reply(sender, reply_subject, reply_body, attachment_path=pdf_path)
    else:
        # Unrecognized subject
        reply_subject = "Re: Your Request"
        reply_body = f"""Dear {sender.split()[0] if ' ' in sender else sender},\n\nThank you for contacting us.\n\nPlease note that we can only process requests for:\n- attestation de scolarite\n- attestation de stage\n\nSubmit a new email with one of these subjects and include the required information.\n\nSincerely,\nThe Yam Team"""
        send_reply(sender, reply_subject, reply_body)

# Monitor the inbox and continuously fetch emails
def monitor_inbox():
    try:
        with imaplib.IMAP4_SSL("imap.gmail.com") as mail:
            mail.login(EMAIL_ACCOUNT, PASSWORD)
            mail.select("inbox")

            # Start continuously checking emails
            while True:
                print("Checking for new emails...")
                _, data = mail.search(None, 'ALL')  # Fetch all emails
                mail_ids = data[0].split() if data[0] else []

                if mail_ids:
                    for i in mail_ids:
                        _, data = mail.fetch(i, '(RFC822)')
                        for response_part in data:
                            if isinstance(response_part, tuple):
                                msg = email.message_from_bytes(response_part[1])
                                process_email(msg)

                # Sleep for a while before checking again
                time.sleep(60)

    except imaplib.IMAP4.error as e:
        print(f"IMAP error: {e}")
    except Exception as e:
        print(f"General Monitoring Error: {e}")


if __name__ == "__main__":
    monitor_inbox()
