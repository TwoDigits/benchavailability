import pandas as pd
import glob
import os
from pathlib import Path
from appscript import app, k
import numpy as np
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

logging.basicConfig(level=logging.INFO)

# Constants
LIST_FILE_EXTENSION_PATTERN = "*.xlsx"
SPLIT_CHAR = "_"
ORG_LEVEL_COLUMN_NAME = 'Org Level 8'
ORG_LEVEL_COLUMN_FILTER_VALUE = 'Full-Stack Development'
AVAILABILITY_STATUS_COLUMN_NAME = 'Availability Status'
AVAILABILITY_STATUS_COLUMN_FILTER_VALUES = ['Now Available', 'Coming Available']
EID_COLUMN_NAME = 'Enterprise ID'
ENTERPRISE_EMAIL_DOMAIN = '@accenture.com'
EMAIL_SUBJECT = 'Actions for Bench Candidates'
EMAIL_TEXT_CONTENT = """
<html>
<head>
    <title></title>
    <link href="https://svc.webspellchecker.net/spellcheck31/lf/scayt3/ckscayt/css/wsc.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif">Hi,<br />
    &nbsp;<br />
    <span style="background-color:rgb(255, 255, 255); color:rgb(13, 13, 13)">You&#39;ve been identified as a candidate currently on the bench or soon to be available. Please take the following actions:</span></span></span>
    <ul>
        <li><span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif">Update your CV on:&nbsp;<a href="https://ts.accenture.com/:f:/r/sites/TwoDigits/Shared%20Documents/General/CVs/Fullstack?csf=1&amp;web=1&amp;e=U6X4Cl">Fullstack CVs</a><span style="background-color:rgb(255, 255, 255); color:rgb(13, 13, 13)">, ensuring that it reflects your recent project experience and the skills you utilized.</span></span></span></li>
        <li><span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif">Open the&nbsp;<a href="https://ts.accenture.com/:u:/r/sites/TwoDigits/SitePages/Fullstack-Bench-Availability.aspx?csf=1&amp;web=1&amp;e=uOqMdO">Fullstack-Bench Availability List</a>:</span></span>
        <ul>
            <li><span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif"><span style="background-color:rgb(255, 255, 255); color:rgb(13, 13, 13)">Click <em>&quot;Edit&quot;</em> at the top right of the page.</span></span></span></li>
            <li><span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif"><span style="background-color:rgb(255, 255, 255); color:rgb(13, 13, 13)">Navigate to the Level Collection that matches your Level.</span></span></span></li>
            <li><span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif"><span style="background-color:rgb(255, 255, 255); color:rgb(13, 13, 13)">Hover at the bottom of the Web-Part containing the Levels you belong to.</span></span></span></li>
            <li><span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif"><span style="background-color:rgb(255, 255, 255); color:rgb(13, 13, 13)">Click the <em>&quot;+&quot;</em> button to add your CV to the list, then select <em>&quot;File and Media&quot;</em> from the Context Menu to open the file dialog.</span></span></span></li>
            <li><span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif"><span style="background-color:rgb(255, 255, 255); color:rgb(13, 13, 13)">To find your updated CV, click on <em>&quot;OneDrive&quot;</em> in the left side panel of the file dialog and then select <em>&quot;Fullstack&quot;</em> to locate your CV and click on <em>&quot;Add File&quot;</em></span></span></span></li>
            <li><span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif"><span style="background-color:rgb(255, 255, 255); color:rgb(13, 13, 13)">Verify the placement of your CV according to the levels and click <em>&quot;Save&quot;</em></span></span></span></li>
        </ul>
        </li>
    </ul>
    <span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif"> <span style="background-color:rgb(255, 255, 255); color:rgb(13, 13, 13)"><strong>Note:</strong> </span></span></span>
    <ul>
        <li><span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif"><span style="background-color:rgb(255, 255, 255); color:rgb(13, 13, 13)">If your staffing is being extended or if you received this email mistakenly, no further action is required.</span></span></span></li>
        <li><span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif"><span style="background-color:rgb(255, 255, 255); color:rgb(13, 13, 13)">If you have recently left the bench or have received a hard booking, please take down your CV from the&nbsp;<a href="https://ts.accenture.com/:u:/r/sites/TwoDigits/SitePages/Fullstack-Bench-Availability.aspx?csf=1&amp;web=1&amp;e=uOqMdO">Fullstack-Bench Availability List</a>.</span></span></span></li>
        <li><span style="font-size:14px"><span style="font-family:arial,helvetica,sans-serif"><span style="background-color:rgb(255, 255, 255); color:rgb(13, 13, 13)">This email is routinely dispatched weekly to all bench candidates. If you have already completed the actions mentioned above, no additional steps are necessary.</span></span></span></li>
        <li>&nbsp;</li>
    </ul>
</body>
</html>
"""

# Function to check availabilities and send reminder


def check_availabilities_and_send_reminder(availability_lists_source_folder_url):
    all_lists = _fetch_source_folder(availability_lists_source_folder_url)
    if all_lists:
        latest_list_item = _pickup_latest_availability_list(availability_lists_source_folder_url, all_lists)
        filtered_bench_candidates_list = _read_availability_list(latest_list_item)
        eid_list_of_bench_candidates = _get_eid_list_of_bench_candidates(filtered_bench_candidates_list)
        email_list_of_bench_candidates = _build_enterprise_emails_from_eid_list(eid_list_of_bench_candidates)
        _send_reminder_email_to_bench_candidates(email_list_of_bench_candidates)
    else:
        logging.warning("No availability list found! Exit process.")

# Function to access the source folder where availability lists are stored


def _fetch_source_folder(availability_lists_source_folder_url):
    logging.info("Accessing the source folder where availability lists are stored ...")
    all_lists = glob.glob(os.path.join(availability_lists_source_folder_url, LIST_FILE_EXTENSION_PATTERN))
    logging.debug(f"Found {len(all_lists)} lists")
    return all_lists

# Function to pickup the latest availability list


def _pickup_latest_availability_list(availability_lists_source_folder_url, all_lists):
    logging.info("Accessing latest availability list ...")
    dates_of_list_items = [Path(list_item).stem.split(SPLIT_CHAR, 1)[0] for list_item in all_lists]
    dates_of_list_items.sort(reverse=True)
    latest_list_date = dates_of_list_items[0]
    latest_list_item = glob.glob(os.path.join(availability_lists_source_folder_url, latest_list_date + LIST_FILE_EXTENSION_PATTERN))
    logging.info("Returning the latest availability list")
    return latest_list_item[0]

# Function to read the availability list and retrieve only bench candidates


def _read_availability_list(latest_list_item):
    logging.info("Reading availability list and retrieving only bench candidates ...")
    candidates_df = pd.read_excel(latest_list_item)
    filtered_df = candidates_df.loc[
        (candidates_df[ORG_LEVEL_COLUMN_NAME] == ORG_LEVEL_COLUMN_FILTER_VALUE) &
        (candidates_df[AVAILABILITY_STATUS_COLUMN_NAME].isin(AVAILABILITY_STATUS_COLUMN_FILTER_VALUES))
    ]
    logging.info("Returning bench candidates")
    return filtered_df

# Function to get the EIDs of the bench candidates


def _get_eid_list_of_bench_candidates(candidates_df):
    logging.info("Getting the EIDs of the bench candidates")
    eid_df = candidates_df[[EID_COLUMN_NAME]]
    return eid_df.values

# Function to build the enterprise E-Mail addresses of the bench candidates


def _build_enterprise_emails_from_eid_list(eid_list_of_bench_candidates):
    logging.info("Building E-Mail addresses of bench candidates according to their EIDs ...")
    email_list_of_bench_candidates = [eid + ENTERPRISE_EMAIL_DOMAIN for eid in eid_list_of_bench_candidates]
    return np.concatenate(email_list_of_bench_candidates, axis=0)

# Function to send reminder E-Mails to bench candidates


def _send_reminder_email_to_bench_candidates(email_list_of_bench_candidates):
    _send_email_using_email_address(email_list_of_bench_candidates)

    # _send_email_using_outlook(email_list_of_bench_candidates)


# Function to send reminder E-Mails using a given email address


def _send_email_using_email_address(email_list_of_bench_candidates):
    # Retrieve sender Email credentials as ENV vars
    sender_email_address = os.getenv('SENDER_EMAIL_ADDRESS')
    sender_email_password = os.getenv('SENDER_EMAIL_PASSWORD')
    email_server_address = os.getenv('EMAIL_SERVER_ADDRESS')
    email_server_port = os.getenv('EMAIL_SERVER_PORT')

    if any([var is None for var in [sender_email_address, sender_email_password, email_server_address, email_server_port]]):
        logging.error('One or more required environment variables for sending E-Mails are missing! Exiting the process.')
        return

    for email in email_list_of_bench_candidates:
        message = MIMEMultipart()
        message["From"] = sender_email_address
        message["To"] = email
        message["Subject"] = EMAIL_SUBJECT
        message.attach(MIMEText(EMAIL_TEXT_CONTENT, "html"))
        try:
            # Connect to the SMTP server
            server = smtplib.SMTP(email_server_address, email_server_port)
            # Secure the connection
            server.starttls()
            # Log in to the email account
            server.login(sender_email_address, sender_email_password)
            # Send the email
            server.sendmail(sender_email_address, email, message.as_string())

            logging.info(f"Reminder E-Mail sent to: {email}")
        except Exception as error:
            print(f"Error when sending E-Mail: {error}")
        finally:
            # Close the connection
            server.quit()


# Function to send reminder E-Mails using Outlook
# Note: This function can only be used for local testing purposes!
# For production, use the function '_send_email_using_email_address' instead


def _send_email_using_outlook(email_list_of_bench_candidates):
    logging.info("Sending reminder E-Mails to bench candidates ...")
    outlook = app('Microsoft Outlook')
    for email in email_list_of_bench_candidates:
        msg = outlook.make(
            new=k.outgoing_message,
            with_properties={
                k.subject: EMAIL_SUBJECT,
                k.plain_text_content: EMAIL_TEXT_CONTENT
            })
        msg.make(
            new=k.recipient,
            with_properties={
                k.email_address: {
                    k.name: email,
                    k.address: email
                }
            })
        msg.open()
        # Uncomment the following line to send the emails
        # msg.send()
        logging.info(f"Reminder E-Mail sent to: {email}")

# Main function


def main():
    # Read the value of the environment variable 'AVAILABILITY_LISTS_SOURCE_FOLDER_URL'
    availability_lists_source_folder_url = os.getenv('AVAILABILITY_LISTS_SOURCE_FOLDER_URL')
    if availability_lists_source_folder_url is None:
        logging.error('No source location for availability lists provided as environment variable. Exiting the process.')
        return
    # Check availabilities and send reminder
    check_availabilities_and_send_reminder(availability_lists_source_folder_url)

# Entry point


if __name__ == "__main__":
    main()
