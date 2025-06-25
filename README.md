# Smart Mail Dispatcher with Real-time Bounce Tracking

This is a Python-based web app built using Streamlit to automate bulk email communication and detect undelivered (bounced) emails in real-time.

## 🔹 What it does

- Send bulk emails using an Excel or CSV file with email addresses
- Upload subject and body content from another Excel/CSV file
- Automatically check your Gmail inbox for bounce reports
- Show which emails bounced
- Download the list of bounced emails
- Resend emails only to bounced recipients

## 🔹 Features

- Simple web interface using Streamlit
- Supports both `.xlsx` and `.csv` file uploads
- Uses Gmail's SMTP and IMAP services
- No code required to use — just upload files and click buttons

## 🔹 Tech Stack

- Python
- Streamlit
- smtplib, imaplib, email
- pandas, openpyxl, re

## 🔹 How to Run (Locally)

1. Clone this repository  
2. Install dependencies:

```bash
pip install -r requirements.txt
