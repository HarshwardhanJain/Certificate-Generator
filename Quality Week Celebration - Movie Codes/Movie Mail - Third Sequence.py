import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from concurrent.futures import ThreadPoolExecutor

def send_email(to_email, cc_email, subject, body, attachment_paths):
    from_email = "harshwardhan22csu392@ncuindia.edu"  # Replace with your email
    from_password = "dhqz hgpc taqo ijps"  # Replace with your app-specific password

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

    # Attach the certificates
    for attachment_path in attachment_paths:
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(attachment_path)}")
            msg.attach(part)

    # Create the SMTP session
    try:
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()  # Enable security
            server.login(from_email, from_password)  # Log in to your email account
            server.send_message(msg)  # Send the email
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")

def find_certificates(base_output_dir, participant):
    certificates = []
    for root, dirs, files in os.walk(base_output_dir):
        for file in files:
            if participant in file and file.endswith(".pdf"):
                certificates.append(os.path.join(root, file))
    return certificates

def read_excel_and_send_certificates(file_path, base_output_dir):
    df = pd.read_excel(file_path)
    print(f"Columns in {file_path}: {df.columns.tolist()}")  # Print column names

    name_column = input("Enter the column name for the participant's name: ")
    roll_no_column = input("Enter the column name for the participant's roll number: ")
    email_column = input("Enter the column name for the participant's email: ")
    cc_email = input("Enter the CC email address (or leave blank if none): ")
    subject = input("Enter the subject line for the email: ")

    participants = df.apply(lambda row: f"{row[name_column]} {row[roll_no_column]}", axis=1).tolist()
    emails = df[email_column].tolist()

    # Thank you note (original)
    thank_you_note = (
        "Dear [Participant's Name],\n\n"
        "We would like to extend our heartfelt gratitude for your participation in the World Quality Week Celebration held on 11th and 12th November.\n\n"
        "Your enthusiasm and engagement in the Poster Making and Movie Screening events added immense value to the celebration. It is through contributions like yours that events like these become truly memorable and impactful.\n\n"
        "Please find your certificate of participation attached as a token of our appreciation.\n\n"
        "Thank you once again for being a part of this enriching experience. We look forward to seeing you at our future events!\n\n"
        "Warm regards,\n"
        "ASQ Student Chapter Team\n"
    )

    # Use ThreadPoolExecutor to send emails concurrently
    with ThreadPoolExecutor(max_workers=5) as executor:  # Adjust max_workers as needed
        for index, email in enumerate(emails):
            participant_name = df.iloc[index][name_column]
            # Replace placeholder with actual participant name
            personalized_thank_you_note = thank_you_note.replace("[Participant's Name]", participant_name)
            
            certificate_paths = find_certificates(base_output_dir, participants[index])
            if certificate_paths:
                # Each email is sent to its own recipient
                executor.submit(send_email, email, cc_email, subject, personalized_thank_you_note, certificate_paths)
            else:
                print(f"Certificates for {participants[index]} not found.")

# Example usage
excel_files = [r"Quality Week Celebration\ASQ_ Movie Coordinator List - World Quality Week Celebration.xlsx",
               r"Quality Week Celebration\ASQ_ Movie Participant List - World Quality Week Celebration.xlsx"]

# Prompt user to select which Excel sheet to start from
print("Available Excel files:")
for i, file in enumerate(excel_files):
    print(f"{i + 1}: {file}")

while True:
    try:
        start_index = int(input("Enter the number of the Excel sheet you would like to start from: ")) - 1
        if 0 <= start_index < len(excel_files):
            break
        else:
            print("Please enter a valid number corresponding to the available Excel files.")
    except ValueError:
        print("Invalid input. Please enter a number.")

base_output_dir = r"Quality Week Celebration\Processed Certificates"

# Process selected Excel sheet and any subsequent sheets
for file in excel_files[start_index:]:
    read_excel_and_send_certificates(file, base_output_dir)