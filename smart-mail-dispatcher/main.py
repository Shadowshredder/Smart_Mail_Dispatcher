import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import imaplib
import email
import re
from openpyxl import Workbook

st.set_page_config(page_title="Smart Mail Dispatcher", layout="centered")

st.title("Smart Mail Dispatcher")
st.write("This app lets you send emails in bulk and track undelivered ones.")

# Upload email list file
st.subheader("Upload Email List File")
email_file = st.file_uploader("Upload 'email_list.xlsx' or CSV", type=["xlsx", "csv"], key="emails")

# Upload message template file
st.subheader("Upload Message Template File")
template_file = st.file_uploader("Upload 'message_template.xlsx' or CSV", type=["xlsx", "csv"], key="template")

# Function to read email addresses
def read_email_list(file):
    df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
    email_list = df["Email Address"].dropna().str.strip().tolist()
    return email_list

# Function to read subject and body
def read_message_template(file):
    df = pd.read_csv(file) if file.name.endswith('.csv') else pd.read_excel(file)
    subject = df["Subject"].dropna().tolist()[0]
    body = df["Body"].dropna().tolist()[0]
    return subject, body

# Function to send bulk emails
def send_bulk_emails(sender_email, sender_password, subject, body, recipients):
    smtp_server = "smtp.gmail.com"
    smtp_port = 587

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)

        success = 0
        fail = 0
        total = len(recipients)
        progress_bar = st.progress(0)

        for i, recipient in enumerate(recipients):
            msg = MIMEMultipart()
            msg["From"] = sender_email
            msg["To"] = recipient
            msg["Subject"] = subject
            msg.attach(MIMEText(body, 'plain'))

            try:
                server.sendmail(sender_email, recipient, msg.as_string())
                success += 1
            except Exception:
                fail += 1

            progress = (i + 1) / total
            progress_bar.progress(progress)

        server.quit()
        st.success(f"Emails sent: {success}")
        if fail > 0:
            st.warning(f"Failed to send to {fail} recipients.")

    except Exception as e:
        st.error(f"Something went wrong: {str(e)}")

# Function to detect bounced emails
def fetch_bounced_emails(email_address, app_password, recipients):
    imap_server = 'imap.gmail.com'
    imap_port = 993
    mailbox = 'INBOX'

    bounced = []

    try:
        mail = imaplib.IMAP4_SSL(imap_server, imap_port)
        mail.login(email_address, app_password)
        mail.select(mailbox)

        status, response = mail.search(None, 'FROM', '"Mail Delivery Subsystem"')
        email_ids = response[0].split()

        for eid in email_ids:
            status, msg_data = mail.fetch(eid, '(RFC822)')
            for part in msg_data:
                if isinstance(part, tuple):
                    msg = email.message_from_bytes(part[1])
                    if msg.is_multipart():
                        for subpart in msg.walk():
                            if subpart.get_content_type() == 'text/plain':
                                content = subpart.get_payload(decode=True).decode('utf-8')
                                match = re.search(r"Your message wasn't delivered to ([\w\.-]+@[\w\.-]+)", content)
                                if match and match.group(1) in recipients:
                                    bounced.append(match.group(1))
                    else:
                        content = msg.get_payload(decode=True).decode('utf-8')
                        match = re.search(r"Your message wasn't delivered to ([\w\.-]+@[\w\.-]+)", content)
                        if match and match.group(1) in recipients:
                            bounced.append(match.group(1))

        mail.logout()

    except Exception as e:
        st.error(f"Error checking inbox: {str(e)}")

    return list(set(bounced))  # remove duplicates

# Function to save bounced emails to Excel
def save_bounced_to_excel(bounced_emails):
    wb = Workbook()
    ws = wb.active
    ws.title = "Bounced Emails"
    ws.append(["Email Address"])
    for email_id in bounced_emails:
        ws.append([email_id])

    filename = "bounced_emails.xlsx"
    wb.save(filename)
    return filename

# Main logic
if email_file and template_file:
    try:
        email_list = read_email_list(email_file)
        subject, body = read_message_template(template_file)

        st.success(f"Loaded {len(email_list)} email addresses successfully.")
        st.subheader("Subject Preview")
        st.write(subject)

        st.subheader("Message Body Preview")
        st.text(body)

        # Input sender credentials
        st.subheader("Enter Your Gmail Credentials")
        sender_email = st.text_input("Gmail address (sender)")
        sender_password = st.text_input("App Password (not Gmail password)", type="password")

        # Send Emails button
        if st.button("Send Emails"):
            if sender_email and sender_password:
                send_bulk_emails(sender_email, sender_password, subject, body, email_list)
            else:
                st.warning("Please enter both email and password.")

        # Check for Bounced Emails button
        if st.button("Check for Bounced Emails"):
            if sender_email and sender_password:
                bounced_emails = fetch_bounced_emails(sender_email, sender_password, email_list)

                if bounced_emails:
                    st.warning("Bounced Emails Found:")
                    for email_addr in bounced_emails:
                        st.write("- " + email_addr)

                    # Save bounced emails and provide download
                    saved_file = save_bounced_to_excel(bounced_emails)
                    with open(saved_file, "rb") as f:
                        st.download_button(
                            label="Download Bounced Emails",
                            data=f,
                            file_name=saved_file,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                    # Step 7: Resend to Bounced Emails
                    if st.button("Resend to Bounced Emails"):
                        st.info("Resending emails to bounced addresses...")
                        send_bulk_emails(sender_email, sender_password, subject, body, bounced_emails)

                else:
                    st.success("No bounced emails found.")
            else:
                st.warning("Please enter your Gmail and App Password to check inbox.")

    except Exception as e:
        st.error(f"Error reading files: {str(e)}")
