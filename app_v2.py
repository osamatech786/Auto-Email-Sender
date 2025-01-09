import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from io import BytesIO
import requests
import msal
from datetime import datetime
import os
from dotenv import load_dotenv

# Set page configuration with a favicon
st.set_page_config(
    page_title="Prevista Auto Email Sender",
    page_icon="https://lirp.cdn-website.com/d8120025/dms3rep/multi/opt/social-image-88w.png", 
    layout="centered"  # "centered" or "wide"
)

# Load environment variables from .env file
load_dotenv()

# Fetch credentials from environment variables
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID")
DRIVE_ID = os.getenv("DRIVE_ID")
EMAIL = os.getenv("EMAIL")
PASSWORD = os.getenv("PASSWORD")
SECRET = os.getenv("SECRET")

# Authenticate and acquire an access token
def acquire_access_token():
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    else:
        raise Exception(f"Failed to acquire token: {result.get('error_description')}")

# Function to determine the current academic year
def current_academic_year():
    current_date = datetime.now()
    if current_date.month >= 8:  # August to December
        return f"{current_date.year}-{str(current_date.year + 1)[-2:]}"
    else:  # January to July
        return f"{current_date.year - 1}-{str(current_date.year)[-2:]}"

# Function to locate the master sheet path in SharePoint
def find_master_sheet_path(access_token, drive_id, folder_path):
    list_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_path}:/children"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(list_url, headers=headers)

    if response.status_code == 200:
        for item in response.json().get("value", []):
            if item["name"].endswith(".xlsx") and "Invoices" in item["name"]:
                return f"{folder_path}/{item['name']}"
    else:
        raise Exception(f"Error fetching folder contents: {response.status_code} - {response.text}")
    raise FileNotFoundError("No master sheet found with '.xlsx' and 'Invoices' in the name.")

# Function to fetch and read recipients from the "Email" sheet of an Excel file
def fetch_recipients_from_excel(access_token, drive_id):
    try:
        academic_year = current_academic_year()
        folder_path = f"AEB Financial/{academic_year}"
        file_path = find_master_sheet_path(access_token, drive_id, folder_path)

        # Download the Excel file
        download_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{file_path}:/content"
        headers = {"Authorization": f"Bearer {access_token}"}
        response = requests.get(download_url, headers=headers)

        if response.status_code != 200:
            raise Exception(f"Error downloading file: {response.status_code} - {response.text}")

        # Load the Excel content
        excel_data = pd.read_excel(BytesIO(response.content), sheet_name="Email")

        if "Name" not in excel_data.columns or "Email" not in excel_data.columns:
            raise ValueError("The 'Email' sheet must contain 'Name' and 'Email' columns.")

        return list(zip(excel_data["Name"], excel_data["Email"]))
    except Exception as e:
        st.error(f"Error fetching recipients: {e}")
        return []

# Function to send an email via Microsoft SMTP
def send_email(sender_email, sender_password, recipient_name, recipient_email, subject, body, attachment=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        # Replace {name} in the body with the actual recipient's name
        personalized_body = body.replace("{name}", recipient_name)
        msg.attach(MIMEText(personalized_body, 'plain'))

        # Attach file if provided
        if attachment:
            attachment_name = attachment.name
            part = MIMEApplication(attachment.getvalue(), Name=attachment_name)
            part['Content-Disposition'] = f'attachment; filename="{attachment_name}"'
            msg.attach(part)

        # Connect to Outlook SMTP
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()

        st.success(f"Email sent successfully to {recipient_name} ({recipient_email})!")
    except Exception as e:
        st.error(f"Error sending email to {recipient_name}: {e}")


def main():

    # Logo space
    st.markdown(
        """
        <style>
            .logo {
                display: flex;
                justify-content: center;
                align-items: center;
                padding: 20px;
                border-bottom: 2px solid #f0f0f0;
                margin-bottom: 20px;
            }
        </style>
        """, unsafe_allow_html=True
    )

    st.image("resources/logo_removed_bg - enlarged.png", use_column_width=True)

    # Page title
    st.markdown(
        """
        <h2 style="text-align:center; color:#4CAF50;">Prevista Auto Email Sender</h2>
        """, unsafe_allow_html=True
    )

    st.divider()    

    # Security
    password = st.text_input("Enter password", type="password")
    if password != SECRET:
        st.stop()

    # Authentication
    access_token = acquire_access_token()

    # Fetch Recipients
    if st.button("Fetch Recipients"):
        recipients = fetch_recipients_from_excel(access_token, DRIVE_ID)
        if recipients:
            st.session_state['recipients'] = recipients
            st.success("Recipients loaded successfully!")
        else:
            st.warning("No recipients found in the Excel file.")

    # Recipient Selection
    recipients = st.session_state.get('recipients', [])
    if recipients:
        all_recipients = [{"label": name, "value": (name, email)} for name, email in recipients]
        selected_recipients = st.multiselect(
            "Select Recipients (All are selected by default)",
            options=all_recipients,
            default=all_recipients,
            format_func=lambda x: x['label']
        )
        selected_emails = [(item['value'][0], item['value'][1]) for item in selected_recipients]
    else:
        selected_emails = []

    # Email Details
    
    st.subheader("Email Details")
    subject = st.text_input("Email Subject", f"Pay time - {datetime.now().strftime('%B')} [Invoice Submission Reminder]")  # Current month
    body = st.text_area(
    "Email Body",
    """Dear {name},

Please submit your invoice, expenses and timesheet (when applicable) via our new Invoice Submission Portal. It will be automatically submitted and verified. If we need to clarify or amend your invoice, we will get in touch. Should you have any questions about the new portal, please get in touch with me and I will help.

To log in to the portal, you will need your Prevista email address and your UTR number (Unique Tax Reference number).
If you don't have a UTR number or it's not working, please get in touch.

Portal Link: https://invoice-v2.streamlit.app/ 

Many thanks :)

Kasia Kwiatkowska 
M: (+34)679161506
    """
    )


    # Send Emails Button
    if st.button("Send Emails"):
        if not all([selected_emails, subject, body]):
            st.error("Please fill in all required fields.")
        else:
            for name, email in selected_emails:
                send_email(EMAIL, PASSWORD, name, email, subject, body)

if __name__ == "__main__":
    main()
