# Prevista Auto Email Sender

Prevista Auto Email Sender is a Streamlit application designed to automate the process of sending emails to recipients listed in an Excel file stored on SharePoint. The application fetches the recipient details, personalizes the email content, and sends the emails via Microsoft SMTP.

## Features

- Fetch recipient details from an Excel file on SharePoint.
- Personalize email content with recipient names.
- Send emails via Microsoft SMTP.
- Secure access with password protection.

## Setup

### Prerequisites

- Python 3.7 or higher
- Streamlit
- Pandas
- Requests
- MSAL
- Python Dotenv

### Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/prevista-auto-email-sender.git
    cd prevista-auto-email-sender
    ```

2. Install the required packages:
    ```sh
    pip install -r requirements.txt
    ```

3. Create a `.env` file in the root directory and add the following environment variables:
    ```env
    CLIENT_ID=your_client_id
    CLIENT_SECRET=your_client_secret
    TENANT_ID=your_tenant_id
    DRIVE_ID=your_drive_id
    EMAIL=your_email
    PASSWORD=your_email_password
    SECRET=your_secret_password
    ```

## Usage

1. Run the Streamlit application:
    ```sh
    streamlit run app_v2.py
    ```

2. Open the application in your web browser. You will be prompted to enter a password. Use the password specified in the `SECRET` environment variable.

3. Click the "Fetch Recipients" button to load the recipient details from the Excel file on SharePoint.

4. Select the recipients you want to send emails to.

5. Fill in the email subject and body. Use `{name}` in the body to personalize the email with the recipient's name.

6. Click the "Send Emails" button to send the emails.

## License

This project is opensource and created for prevista.co.uk

Dev : https://linkedin.com/in/osamatech786