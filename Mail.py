import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

def send_email(to_email, cc_email, subject, body, attachment_path):
    from_email = "harshwardhan22csu392@ncuindia.edu"  # Replace with your email
    from_password = "umrp qgqg rdyh qscq"  # Replace with your app-specific password

    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    if cc_email:
        msg['Cc'] = cc_email  # Add CC if provided
        to_email = [to_email] + [cc_email]  # Include CC in the recipient list

    # Attach the body with the msg instance
    msg.attach(MIMEText(body, 'plain'))

    # Attach the certificate
    if attachment_path:
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(attachment_path)}")
            msg.attach(part)

    # Create the SMTP session
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()  # Enable security
        server.login(from_email, from_password)  # Log in to your email account
        server.send_message(msg)  # Send the email

    print(f"Email sent to {to_email}")

def find_certificate(base_output_dir, participant):
    for root, dirs, files in os.walk(base_output_dir):
        for file in files:
            if file == f"{participant}.pdf":
                return os.path.join(root, file)
    return None

def read_excel_and_send_certificates(file_path, base_output_dir):
    df = pd.read_excel(file_path)
    print(f"Columns in {file_path}: {df.columns.tolist()}")  # Print column names

    name_column = input("Enter the column name for the participant's name: ")
    roll_no_column = input("Enter the column name for the participant's roll number: ")
    email_column = input("Enter the column name for the participant's email: ")
    cc_email = input("Enter the CC email address (or leave blank if none): ")

    participants = df.apply(lambda row: f"{row[name_column]} {row[roll_no_column]}", axis=1).tolist()
    emails = df[email_column].tolist()

    # Thank you note
    thank_you_note = (
        "Dear [Participant's Name],\n\n"
        "We would like to extend our heartfelt gratitude for your participation in the Quality Management Workshop organized by the ASQ Student Chapter. "
        "Your engagement and enthusiasm contributed significantly to the success of the event.\n\n"
        "We hope that the insights shared during the workshop, particularly on Lean Manufacturing and Quality Management Systems, will empower you to enhance quality and efficiency in your respective fields.\n\n"
        "Thank you once again for being a part of this enriching experience. We look forward to seeing you at future events!\n\n"
        "Warm regards,\n"
        "ASQ Student Chapter Team\n"
    )

    for index, email in enumerate(emails):
        participant_name = df.iloc[index][name_column]
        # Replace placeholder with actual participant name
        personalized_thank_you_note = thank_you_note.replace("[Participant's Name]", participant_name)
        
        certificate_path = find_certificate(base_output_dir, participants[index])
        if certificate_path:
            send_email(email, cc_email, "Your Certificate", personalized_thank_you_note, certificate_path)
        else:
            print(f"Certificate for {participants[index]} not found.")

# Example usage
excel_files = [r"Quality Management Workshop\Test - Mail.xlsx"]

base_output_dir = r"Quality Management Workshop\Processed Certificates"

for file in excel_files:
    read_excel_and_send_certificates(file, base_output_dir)