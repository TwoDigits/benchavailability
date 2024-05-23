# Bench Availability Reminder

This repository contains a Python script `bench_availability_reminder.py` that is used to process Excel files containing bench availability information and send email reminders to bench candidates to update their CVs and put them on the Availability Page.

## Dependencies

The script uses the following Python libraries:

- pandas
- glob
- os
- pathlib
- appscript
- numpy
- logging
- smtplib
- email

## How it works

The script scans a directory for Excel files, reads them into pandas dataframes, and filters the data based on availability criteria. It then sends an email reminder based on the filtered data to bench candidates' E-Mail addresses.

## Constants

The script uses the following constants:

- `LIST_FILE_EXTENSION_PATTERN`: The file extension pattern to look for in the directory.
- `SPLIT_CHAR`: The character used to split strings.
- `ORG_LEVEL_COLUMN_NAME`: The name of the column in the Excel file that contains the organization level.
- `ORG_LEVEL_COLUMN_FILTER_VALUE`: The value in the organization level column to filter on.
- `AVAILABILITY_STATUS_COLUMN_NAME`: The name of the column in the Excel file that contains the availability status.
- `AVAILABILITY_STATUS_COLUMN_FILTER_VALUES`: The values in the availability status column to filter on.
- `EID_COLUMN_NAME`: The name of the column in the Excel file that contains the enterprise ID.
- `ENTERPRISE_EMAIL_DOMAIN`: The domain for the enterprise email.
- `EMAIL_SUBJECT`: The subject of the email to be sent.
- `EMAIL_TEXT_CONTENT`: The content of the email to be sent.

To be able to retrieve availability lists, following environment variable **MUST** be provided:
- `AVAILABILITY_LISTS_SOURCE_FOLDER_URL`: The location folder where the availability lists are stored

To make the script able to send E-Mails, following environment varaibles **MUST** be provided:
- `SENDER_EMAIL_ADDRESS`: The E-Mail address of the sender
- `SENDER_EMAIL_PASSWORD`: The password for the above mentioned E-Mail address
- `EMAIL_SERVER_ADDRESS`: The E-Mail server address
- `EMAIL_SERVER_PORT`: The port of the above mentioned E-Mail server

## Usage

To use the script, simply run it in your Python environment. Make sure to adjust the constants and environment variables to match your specific use case.
