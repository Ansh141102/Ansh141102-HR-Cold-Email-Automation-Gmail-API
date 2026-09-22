# HR  Cold-Email Automation — Gmail API

A Python-based HR outreach automation system that reads HR contact data from Excel, personalizes an email for each HR contact, attaches a resume, and sends emails through Gmail using the official Gmail API.

The project is designed around a **controlled batch workflow** so emails are reviewed and sent in manageable groups instead of being sent blindly to an entire database.

---

## Table of Contents

1. [What This Project Does](#what-this-project-does)
2. [Project Workflow](#project-workflow)
3. [Production Project Structure](#production-project-structure)
4. [How Personalization Works](#how-personalization-works)
5. [Requirements](#requirements)
6. [Google Cloud and Gmail API Setup](#google-cloud-and-gmail-api-setup)
7. [Installation](#installation)
8. [Prepare the Excel File](#prepare-the-excel-file)
9. [Configure Your Resume](#configure-your-resume)
10. [Configure the Email Template](#configure-the-email-template)
11. [Gmail Authentication](#gmail-authentication)
12. [Testing Before Production](#testing-before-production)
13. [Production Sending](#production-sending)
14. [Batch System](#batch-system)
15. [Duplicate and Already-Sent Protection](#duplicate-and-already-sent-protection)
16. [Logs](#logs)
17. [Starting From the Beginning](#starting-from-the-beginning)
18. [Using the Project With Another Gmail Account](#using-the-project-with-another-gmail-account)
19. [Common Problems](#common-problems)
20. [Security](#security)
21. [Recommended `.gitignore`](#recommended-gitignore)
22. [Important Safety Rules](#important-safety-rules)
23. [Typical Workflow](#typical-workflow)
24. [Troubleshooting Checklist](#troubleshooting-checklist)

---

# What This Project Does

This project automates personalized job-outreach emails.

Instead of manually writing an email for every HR contact, the program reads contact information from Excel and inserts the recipient's details into a reusable email template.

For example, an Excel row might contain:

| HR Name | HR Email | Company |
|---|---|---|
| Hima Kulshrestha | hima@example.com | Example Technologies |

The generated email can then contain:

```text
Hi Hima Kulshrestha,

I am reaching out to explore entry-level opportunities at
Example Technologies...
```

The resume is automatically attached.

The production sender uses Gmail API authentication and sends through the Gmail account authorized by the user.

---

# Project Workflow

The complete workflow is:

```text
HR Excel
   |
   v
Data Cleaning
   |
   v
Company-name cleanup
   |
   v
Final Excel
   |
   v
Data Quality Check
   |
   v
Gmail OAuth Authentication
   |
   v
Personalized Email Generation
   |
   v
Batch Review
   |
   v
Human Confirmation
   |
   v
Gmail API
   |
   v
Email Sent
   |
   v
Send Log Updated
```

The production principle is:

```text
Excel -> Personalize -> Select Batch -> Review -> Confirm -> Send -> Log
```

---

# Production Project Structure

After the project has been cleaned up, the production folder can look like this:

```text
HR_Email_Automation/
|
+-- venv/
|
+-- Ansh_Srivastava_Resume.pdf
|
+-- credentials.json
+-- token.json
|
+-- HR_Email_Automation_Final.xlsx
|
+-- email_template.py
+-- gmail_auth.py
+-- gmail_send_production.py
|
+-- email_send_log.csv
|
+-- README.md
```

## File descriptions

### `venv/`

Python virtual environment containing the installed packages.

It is better to create a new virtual environment on another computer instead of copying this folder.

### `Ansh_Srivastava_Resume.pdf`

The resume attached to outgoing emails.

Another user should replace this with their own resume and update the filename/path in the script if necessary.

### `credentials.json`

Google OAuth client credentials downloaded from Google Cloud.

**Never publish this file publicly.**

### `token.json`

The authorization token generated after Gmail authentication.

This is account-specific.

**Never share your `token.json` with another user.**

### `HR_Email_Automation_Final.xlsx`

The production HR database used by the sender.

### `email_template.py`

Contains the email subject/body template and personalization logic.

### `gmail_auth.py`

Authenticates the Gmail account using Google OAuth.

### `gmail_send_production.py`

The main production sender. It reads the Excel file, skips already-sent contacts, creates personalized HTML emails, attaches the resume, sends the selected batch, and updates the send log.

### `email_send_log.csv`

Production history used to prevent the same address from being sent again.

**Do not delete this unless you intentionally want to reset the sending history.**

### `README.md`

Project documentation.

---

# How Personalization Works

The Excel data provides the values that change for every recipient.

Conceptually:

```python
hr_name = row["HR Name"]
hr_email = row["HR Email"]
company = row["Company"]
```

The template then uses those values.

For example:

```text
HR Name  = Hima Kulshrestha
HR Email = hima@example.com
Company  = Example Technologies
```

becomes:

```text
To: hima@example.com

Hi Hima Kulshrestha,

I am reaching out to explore entry-level opportunities at
Example Technologies...
```

The same template can therefore be reused for many contacts.

---

# Requirements

You need:

- Windows, macOS, or Linux
- Python 3
- A Gmail/Google account
- A Google Cloud project
- Gmail API enabled
- OAuth credentials
- An HR Excel file
- A PDF resume
- Internet access

The project uses packages such as:

```bash
python -m pip install pandas openpyxl google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

If the project already has a working environment, inspect installed packages with:

```bash
python -m pip list
```

---

# Google Cloud and Gmail API Setup

Each user should authorize **their own Gmail account**.

## 1. Create or select a Google Cloud project

Open:

https://console.cloud.google.com/

Create a project or select an existing project.

## 2. Enable Gmail API

Go to:

```text
APIs & Services
    -> Library
    -> Gmail API
    -> Enable
```

## 3. Configure OAuth consent

Go to:

```text
APIs & Services
    -> OAuth consent screen
```

Configure the application.

If Google requires test users for the selected configuration, add the Gmail account that will authorize the application.

## 4. Create OAuth credentials

Go to:

```text
APIs & Services
    -> Credentials
    -> Create Credentials
    -> OAuth client ID
```

Create credentials suitable for a desktop application.

Download the JSON file and rename it:

```text
credentials.json
```

Place it in the project directory:

```text
HR_Email_Automation/
+-- credentials.json
+-- gmail_auth.py
+-- gmail_send_production.py
```

---

# Installation

Copy or clone the project onto the new computer.

For example:

```bash
git clone YOUR_REPOSITORY_URL
cd HR_Email_Automation
```

If using a ZIP file, extract it and open a terminal inside the project directory.

## Create a virtual environment

Windows:

```bash
python -m venv venv
```

Git Bash:

```bash
source venv/Scripts/activate
```

Command Prompt:

```cmd
venv\Scripts\activate
```

PowerShell:

```powershell
venv\Scripts\Activate.ps1
```

Install dependencies:

```bash
python -m pip install pandas openpyxl google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

---

# Prepare the Excel File

The production sender uses:

```text
HR_Email_Automation_Final.xlsx
```

The worksheet name must exactly match the name expected by the Python script.

The current project uses:

```text
HR Email
```

Be careful about spaces.

These are different worksheet names:

```text
HR Email
HR Email 
```

A mismatch can produce:

```text
Worksheet named 'HR Email' not found
```

## Important columns

The current workflow uses these important fields:

```text
HR Name
HR Email
Company
```

Example:

```text
| HR Name      | HR Email          | Company              |
|--------------|-------------------|----------------------|
| Rahul Sharma | rahul@example.com | Example Technologies |
| Priya Singh  | priya@example.com | ABC Software         |
```

The exact column names must match what the Python scripts expect.

---

# Configure Your Resume

Replace the example resume:

```text
Ansh_Srivastava_Resume.pdf
```

with your own resume.

For example:

```text
Rahul_Sharma_Resume.pdf
```

Update the filename/path in the sender if it is hard-coded.

The PDF must exist in the expected project location.

---

# Configure the Email Template

Open:

```text
email_template.py
```

Customize the message with your own:

- name
- degree
- skills
- experience
- target roles
- portfolio
- GitHub
- LinkedIn
- phone number
- email signature

Keep the placeholders required by the code.

The design is:

```text
Your fixed information
        +
HR Name
        +
Company
        |
        v
Personalized email
```

The project uses HTML email formatting so supported content such as bold text can be rendered correctly in Gmail.

---

# Gmail Authentication

On a new computer/account, run:

```bash
python gmail_auth.py
```

The script will display a Google authorization URL.

Open it in a browser, sign in to the Gmail account that should send the emails, and approve the requested permission.

A successful authentication creates:

```text
token.json
```

The terminal should report successful Gmail authentication.

## Important

The Gmail account used by the sender is the Google account that authorized the application.

Therefore:

```text
Gmail account
     |
     v
gmail_auth.py
     |
     v
Google OAuth
     |
     v
token.json
     |
     v
gmail_send_production.py
     |
     v
Emails sent from that account
```

Another user should never copy your `token.json`.

---

# Testing Before Production

Before enabling actual sending, test the project.

A good testing sequence is:

1. Test the Excel reading.
2. Test HR name replacement.
3. Test company replacement.
4. Test email address selection.
5. Test subject.
6. Test HTML formatting.
7. Test bold formatting.
8. Test resume attachment.
9. Review Gmail drafts.
10. Only then enable actual sending.

The existing project followed this approach by generating Gmail drafts before moving to production sending.

Whenever the email template or personalization code is changed, repeat a small test before production sending.

---

# Production Sending

Run:

```bash
python gmail_send_production.py
```

The program displays the configuration and performs a safety check.

A production run should clearly indicate:

```text
MODE             : ACTUAL EMAIL SENDING
EMAIL SENDING    : ENABLED
```

The program then shows the contacts in the current batch.

Example:

```text
CURRENT SEND BATCH

01. HR Name | hr1@example.com | Company A
02. HR Name | hr2@example.com | Company B
...
25. HR Name | hr25@example.com | Company Y
```

It then asks for explicit confirmation:

```text
Type SEND to send this batch:
```

Only after the required confirmation should the batch be sent.

---

# Batch System

The production setup uses a batch size of:

```text
25
```

The idea is:

```text
All Excel rows
      |
      v
Remove already-sent addresses
      |
      v
Next 25 unsent contacts
      |
      v
Display batch
      |
      v
Review
      |
      v
Confirm SEND
      |
      v
Send
      |
      v
Update log
```

A smaller batch can be useful when changing templates or starting a new campaign.

Do not increase the batch size blindly. Keep the sending rate appropriate for the Gmail account and your intended outreach.

---

# Duplicate and Already-Sent Protection

The sender uses:

```text
email_send_log.csv
```

to track successful sends.

For example:

```text
First run:
a@example.com -> sent
b@example.com -> sent
c@example.com -> sent

Second run:
a@example.com -> SKIP
b@example.com -> SKIP
c@example.com -> SKIP
d@example.com -> send
e@example.com -> send
```

This allows the Excel database to contain duplicate rows without automatically sending the same address repeatedly.

## Excel duplicates vs duplicate sends

These are different concepts:

```text
Duplicate row in Excel
        !=
Duplicate email sent
```

The Excel file is not automatically rewritten just because a duplicate email exists.

The send log is responsible for production send-history protection.

---

# Logs

## `email_send_log.csv`

This is the main production history.

It is used to remember contacts that have already been processed/sent.

Keep it with the production project.

### Important

Do not delete or empty it casually.

Deleting the log can cause previously contacted addresses to become eligible for sending again.

---

# Starting From the Beginning

If you intentionally want to restart a campaign from the first row, the important point is that the sender uses the send log.

The Excel row number alone does not determine whether an email is eligible.

If you intentionally want a completely fresh sending history:

1. Stop the sender.
2. Back up `email_send_log.csv`.
3. Reset the active send log according to the format expected by the script.
4. Run a small test.
5. Review the batch.
6. Continue with production sending.

Recommended backup:

```text
email_send_log.csv
        |
        v
email_send_log_backup.csv
```

Do not reset the log simply because you deleted Gmail drafts. Gmail drafts and the production send log are separate things.

---

# Using the Project With Another Gmail Account

Suppose User A uses:

```text
userA@gmail.com
```

and User B wants to use:

```text
userB@gmail.com
```

User B should use their own authentication.

## Step 1

Set up Google Cloud/Gmail API access.

## Step 2

Place the appropriate `credentials.json` in the project.

## Step 3

Do **not** copy User A's:

```text
token.json
```

## Step 4

Run:

```bash
python gmail_auth.py
```

## Step 5

Sign in to:

```text
userB@gmail.com
```

## Step 6

Authorize Gmail access.

## Step 7

Verify the sender account with a controlled test.

## Step 8

Run:

```bash
python gmail_send_production.py
```

The emails will be sent through the account that authenticated with OAuth.

---

# Common Problems

## `Worksheet named 'HR Email' not found`

Check the exact worksheet name in Excel.

The following are not necessarily equivalent:

```text
HR Email
HR Email 
```

Also verify that the sender is opening the intended workbook.

---

## `credentials.json` not found

Make sure it is located in the project directory:

```text
HR_Email_Automation/
+-- credentials.json
+-- gmail_auth.py
+-- gmail_send_production.py
```

---

## Authentication problems

Run:

```bash
python gmail_auth.py
```

and complete the Google authorization flow.

Do not share another user's `token.json`.

---

## Resume not found

Verify that the configured PDF filename exactly matches the file on disk.

For example:

```text
Rahul_Sharma_Resume.pdf
```

is different from:

```text
Rahul_Sharma_Resume (1).pdf
```

---

## Permission denied while writing a CSV

If you see:

```text
PermissionError: [Errno 13] Permission denied
```

close the CSV file in Excel or another program and run the command again.

---

## Wrong Gmail account

The OAuth account controls the sending account.

Run authentication again with the intended Google account and verify the account before production sending.

---

## Recipient is skipped

If you see:

```text
[SKIP - ALREADY PROCESSED]
```

inspect:

```text
email_send_log.csv
```

The address is likely already recorded there.

This is intentional duplicate-send protection.

---

# Security

This project may contain:

- HR names
- HR email addresses
- your resume
- Google OAuth credentials
- Gmail authorization tokens
- email-sending history

Treat these as private data.

## Never publicly upload:

```text
credentials.json
token.json
email_send_log.csv
```

The Excel database and resume may also contain personal information and should normally remain private.

---

# Recommended `.gitignore`

If publishing the source code to GitHub, create:

```text
.gitignore
```

with:

```gitignore
# Python
venv/
__pycache__/
*.pyc

# Google OAuth secrets
credentials.json
token.json

# Personal/contact data
*.xlsx
*.csv

# Personal documents
*.pdf

# Environment files
.env
```

This prevents credentials, contact databases, logs, and personal documents from being accidentally committed.

---

# Important Safety Rules

## 1. Review every batch

Do not blindly send the entire database.

Use a controlled batch such as:

```text
25 contacts
```

Review the terminal output before confirming.

## 2. Test template changes

Whenever `email_template.py` changes, test:

- subject
- HR name
- company
- formatting
- links
- signature
- attachment

before actual sending.

## 3. Keep the send log

Do not delete:

```text
email_send_log.csv
```

unless you intentionally want to reset sending history.

## 4. Verify the Gmail account

Before the first production batch, make sure the OAuth account is the account you actually intend to use.

## 5. Use legitimate outreach

Use contact information and outreach practices that are appropriate for your situation and comply with applicable laws, platform policies, and organizational policies.

Do not use misleading personalization.

---

# Typical Workflow

Once setup is complete:

## Step 1 — Update the Excel database

Update:

```text
HR_Email_Automation_Final.xlsx
```

with new contacts.

## Step 2 — Run the sender

```bash
python gmail_send_production.py
```

## Step 3 — Review the batch

The terminal displays the contacts about to receive emails.

Example:

```text
01. HR Name | hr1@example.com | Company A
02. HR Name | hr2@example.com | Company B
...
25. HR Name | hr25@example.com | Company Y
```

## Step 4 — Confirm

If everything is correct:

```text
Type SEND to send this batch:
```

Enter:

```text
SEND
```

## Step 5 — Check the result

The program reports the result and updates the send log.

## Step 6 — Run again

Run:

```bash
python gmail_send_production.py
```

again to process the next eligible contacts.

Previously sent addresses are skipped.

---

# Troubleshooting Checklist

Before diagnosing a failure, check:

```text
[ ] Python virtual environment is active
[ ] Required packages are installed
[ ] credentials.json exists
[ ] token.json exists / authentication completed
[ ] Correct Gmail account was authorized
[ ] HR_Email_Automation_Final.xlsx exists
[ ] Worksheet name exactly matches the script
[ ] Required Excel columns exist
[ ] Resume PDF exists
[ ] email_template.py is configured
[ ] email_send_log.csv is present
[ ] Batch size is correct
[ ] No required CSV is open in Excel
[ ] Terminal is running from the project directory
```

---

# Quick Start for a New User

For someone who has copied the project source:

```bash
# 1. Open the project
cd HR_Email_Automation

# 2. Create a virtual environment
python -m venv venv

# 3. Activate it in Git Bash
source venv/Scripts/activate

# 4. Install dependencies
python -m pip install pandas openpyxl google-api-python-client google-auth-httplib2 google-auth-oauthlib

# 5. Put your own credentials.json in the project folder

# 6. Put your own resume in the project folder

# 7. Prepare HR_Email_Automation_Final.xlsx

# 8. Authenticate your Gmail account
python gmail_auth.py

# 9. Perform a controlled test

# 10. Start production sending
python gmail_send_production.py
```

The new user must authorize **their own Gmail account** during authentication.

---

# Project Architecture

The project can be understood as five layers:

```text
+------------------------------+
| 1. DATA                      |
| HR_Email_Automation_Final    |
| .xlsx                        |
+--------------+---------------+
               |
               v
+------------------------------+
| 2. TEMPLATE                  |
| email_template.py            |
| Personalization + HTML       |
+--------------+---------------+
               |
               v
+------------------------------+
| 3. AUTHENTICATION            |
| gmail_auth.py                |
| credentials.json/token.json  |
+--------------+---------------+
               |
               v
+------------------------------+
| 4. PRODUCTION SENDER         |
| gmail_send_production.py     |
| Batch + confirmation         |
+--------------+---------------+
               |
               v
+------------------------------+
| 5. HISTORY                   |
| email_send_log.csv           |
| Duplicate-send protection    |
+------------------------------+
```

---

# Design Philosophy

The key principle is:

```text
Automate repetitive work,
but keep the final sending decision under human control.
```

The program handles:

- Excel reading
- personalization
- HTML formatting
- resume attachment
- Gmail API communication
- batch selection
- send-history tracking

The user controls:

- Gmail account
- recipient database
- email template
- batch size
- whether a batch should actually be sent

---

# Final Production Workflow

```text
                  +-----------------------+
                  | HR Excel Database     |
                  | Final.xlsx            |
                  +-----------+-----------+
                              |
                              v
                  +-----------------------+
                  | Read HR Contact       |
                  | Name / Email / Company|
                  +-----------+-----------+
                              |
                              v
                  +-----------------------+
                  | Check Send Log        |
                  | Already sent?         |
                  +-------+---------+-----+
                          | YES     | NO
                          |         |
                          v         v
                        SKIP    Personalize
                                    |
                                    v
                              Attach Resume
                                    |
                                    v
                              Build HTML Email
                                    |
                                    v
                              Select Batch (25)
                                    |
                                    v
                              Human Review
                                    |
                         +----------+----------+
                         |                     |
                       SEND                DON'T SEND
                         |                     |
                         v                     v
                     Gmail API               Stop
                         |
                         v
                     Email Sent
                         |
                         v
                Update send log
```

---

## Final Notes

This README describes the production workflow and the files currently used by the project.

The exact Python implementation remains the source of truth for:

- Excel column names
- worksheet name
- file paths
- batch-selection behavior
- send-log format
- OAuth scope
- email-template placeholders

If any of those are changed in the code, update this README accordingly.
