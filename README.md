# quote-generation-flask-app
# JIRA Webhook Receiver for Automated Document Handling

This Flask application listens for JIRA webhook events to automate the processing of issues, generating documents based on templates retrieved from Confluence, converting them to PDF, and uploading them to SharePoint. The application also interacts with JIRA to update issue details and attach generated documents.

## Features

- Listen for JIRA webhook POST requests.
- Fetch document templates from Confluence.
- Generate Word documents and convert them to PDF.
- Upload documents to SharePoint and attach them to JIRA issues.
- Update JIRA issues based on document processing results.

## Requirements

- Python 3.8+
- Flask
- Requests
- python-docx
- docx2pdf (Windows only)
- PyWin32 (Windows only for COM support)

## Installation

- Install the following libraries
- pip install Flask
- pip install python-docx
- pip install Werkzeug
- pip install docx2pdf
- pip install jira
- pip install requests
- pip install pywin32

## Create Config.py page

This script uses a separate python page to save sensative information such as API Tokens, passwords, etc.

To replicate the following flask app, convert config.txt file to config.py and define objects. 




