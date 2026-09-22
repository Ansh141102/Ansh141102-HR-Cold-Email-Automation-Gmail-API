import os
import re
import csv
import time
import base64
from datetime import datetime

import pandas as pd

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


# ============================================================
# CONFIGURATION
# ============================================================

EXCEL_FILE = "HR_Email_Automation_Final.xlsx"
SHEET_NAME = "HR Email "

RESUME_FILE = "Ansh_Srivastava_Resume.pdf"

TOKEN_FILE = "token.json"

# NEW sending log.
# This is intentionally separate from email_draft_log.csv
SEND_LOG_FILE = "email_send_log.csv"

# IMPORTANT:
# First production batch = 25 emails
BATCH_SIZE = 25

# Small delay between emails
DELAY_SECONDS = 2

# Gmail API scope required for actual sending
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]


# ============================================================
# EMAIL TEMPLATE
# ============================================================

SUBJECT_TEMPLATE = (
    "Exploring Entry-Level Opportunities at {company_name} | Data & AI"
)


def create_email_body(hr_name, company_name):
    """
    Creates the personalized HTML email.
    Required portions are bolded using HTML <strong> tags.
    """

    return f"""
<html>
<body>

<p>Hi {hr_name},</p>

<p>I hope you're doing well.</p>

<p>
I’m reaching out to explore entry-level opportunities at
<strong>{company_name}</strong> in
<strong>Data Analytics, Data Engineering, Machine Learning or Python Development.</strong>
I’m a B.Tech graduate in Computer Science &amp; Engineering (Data Science)
and currently looking for an entry-level/associate/junior role where I can
apply my technical skills while learning from an experienced team.
</p>

<p>
My hands-on experience includes
<strong>
Python, SQL, Pandas, NumPy, Power BI, Tableau, Scikit-learn, XGBoost,
TensorFlow, PyTorch, ETL/ELT, Apache Spark and Apache Kafka.
</strong>
I’ve worked on projects involving data analysis, visualization,
machine learning and end-to-end data/ML workflows.
</p>

<p>
For example, I built a
<strong>Tableau dashboard using 130K+ EV registration records</strong>
to analyze adoption trends, performed end-to-end analysis of a
<strong>7,043-customer telecom dataset</strong>
and developed ML applications including a stock price predictor and a
heart failure prediction system.
</p>

<p>
I’m particularly interested in opportunities where I can work with
real-world data, engineering pipelines, analytics or AI/ML solutions.
</p>

<p>
I’ve attached my resume for your reference. If you’re aware of any
suitable openings for freshers or entry-level candidates, I would really
appreciate it if you could guide me toward the relevant opportunity or
application process.
</p>

<p>
Thank you for taking the time to read my message. I’d be grateful for any
guidance you can provide.
</p>

<p>
Best regards,<br>
Ansh Srivastava<br>
B.Tech – Computer Science &amp; Engineering (Data Science)<br>
+91-9599091751<br>
srivastavaansh171@gmail.com
</p>

</body>
</html>
"""


# ============================================================
# BASIC EMAIL VALIDATION
# ============================================================

def is_valid_email(email):
    if not isinstance(email, str):
        return False

    email = email.strip()

    pattern = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"

    return bool(re.match(pattern, email))


# ============================================================
# LOAD SEND LOG
# ============================================================

def load_send_log():
    """
    Reads the NEW sending log.

    The old email_draft_log.csv is deliberately ignored.
    """

    if not os.path.exists(SEND_LOG_FILE):
        return set()

    try:
        df = pd.read_csv(SEND_LOG_FILE)

        if "HR Email" not in df.columns:
            return set()

        emails = (
            df["HR Email"]
            .dropna()
            .astype(str)
            .str.strip()
            .str.lower()
        )

        return set(emails)

    except Exception as e:
        print(f"[WARNING] Could not read send log: {e}")
        return set()


# ============================================================
# APPEND TO SEND LOG
# ============================================================

def log_send(
    hr_name,
    hr_email,
    company,
    excel_row,
    status,
    details=""
):
    file_exists = os.path.exists(SEND_LOG_FILE)

    with open(
        SEND_LOG_FILE,
        "a",
        newline="",
        encoding="utf-8-sig"
    ) as f:

        writer = csv.writer(f)

        if not file_exists:
            writer.writerow([
                "Timestamp",
                "Excel Row",
                "HR Name",
                "HR Email",
                "Company",
                "Status",
                "Details"
            ])

        writer.writerow([
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            excel_row,
            hr_name,
            hr_email,
            company,
            status,
            details
        ])


# ============================================================
# GMAIL SERVICE
# ============================================================

def get_gmail_service():

    if not os.path.exists(TOKEN_FILE):
        raise FileNotFoundError(
            "token.json not found. Run gmail_auth.py first."
        )

    creds = Credentials.from_authorized_user_file(
        TOKEN_FILE,
        SCOPES
    )

    if not creds.valid:
        raise RuntimeError(
            "Gmail token is invalid or does not contain "
            "the gmail.send permission. Delete token.json "
            "and run gmail_auth.py again."
        )

    service = build(
        "gmail",
        "v1",
        credentials=creds
    )

    return service


# ============================================================
# CREATE MIME EMAIL
# ============================================================

def create_message(
    to_email,
    subject,
    html_body,
    resume_path
):

    message = MIMEMultipart("mixed")

    message["To"] = to_email
    message["Subject"] = subject

    # HTML body
    alternative = MIMEMultipart("alternative")

    html_part = MIMEText(
        html_body,
        "html",
        "utf-8"
    )

    alternative.attach(html_part)

    message.attach(alternative)

    # Resume attachment
    with open(resume_path, "rb") as attachment_file:

        attachment = MIMEBase(
            "application",
            "pdf"
        )

        attachment.set_payload(
            attachment_file.read()
        )

    encoders.encode_base64(attachment)

    attachment.add_header(
        "Content-Disposition",
        "attachment",
        filename=os.path.basename(resume_path)
    )

    message.attach(attachment)

    raw_message = base64.urlsafe_b64encode(
        message.as_bytes()
    ).decode()

    return {
        "raw": raw_message
    }


# ============================================================
# SEND EMAIL
# ============================================================

def send_email(
    service,
    hr_name,
    hr_email,
    company_name
):

    subject = SUBJECT_TEMPLATE.format(
        company_name=company_name
    )

    body = create_email_body(
        hr_name,
        company_name
    )

    message = create_message(
        to_email=hr_email,
        subject=subject,
        html_body=body,
        resume_path=RESUME_FILE
    )

    result = service.users().messages().send(
        userId="me",
        body=message
    ).execute()

    return result.get("id")


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print("=" * 80)
    print("HR EMAIL AUTOMATION — PRODUCTION SENDER")
    print("=" * 80)
    print()

    print(f"Excel file       : {EXCEL_FILE}")
    print(f"Sheet            : {SHEET_NAME}")
    print(f"Resume           : {RESUME_FILE}")
    print(f"Send log         : {SEND_LOG_FILE}")
    print(f"Batch size       : {BATCH_SIZE}")
    print()
    print("MODE             : ACTUAL EMAIL SENDING")
    print("EMAIL SENDING    : ENABLED")
    print()

    print("=" * 80)
    print("SAFETY CHECK")
    print("=" * 80)

    if not os.path.exists(EXCEL_FILE):
        print(f"[ERROR] Excel file not found: {EXCEL_FILE}")
        return

    if not os.path.exists(RESUME_FILE):
        print(f"[ERROR] Resume not found: {RESUME_FILE}")
        return

    if not os.path.exists(TOKEN_FILE):
        print("[ERROR] token.json not found.")
        print("Run gmail_auth.py first.")
        return

    # --------------------------------------------------------
    # LOAD EXCEL
    # --------------------------------------------------------

    try:
        df = pd.read_excel(
            EXCEL_FILE,
            sheet_name=SHEET_NAME
        )
    except Exception as e:
        print(f"[ERROR] Could not read Excel file: {e}")
        return

    print(f"Total Excel rows : {len(df)}")

    # --------------------------------------------------------
    # LOAD SEND HISTORY
    # --------------------------------------------------------

    sent_emails = load_send_log()

    print(
        f"Previously sent emails : {len(sent_emails)}"
    )

    print()

    # --------------------------------------------------------
    # REQUIRED COLUMNS
    # --------------------------------------------------------

    required_columns = [
        "HR Name",
        "HR Email",
        "Company"
    ]

    missing_columns = [
        col for col in required_columns
        if col not in df.columns
    ]

    if missing_columns:

        print(
            "[ERROR] Missing required columns:"
        )

        for col in missing_columns:
            print(f"  - {col}")

        return

    # --------------------------------------------------------
    # CONNECT TO GMAIL
    # --------------------------------------------------------

    try:

        service = get_gmail_service()

    except Exception as e:

        print()
        print("[ERROR] Gmail authentication problem:")
        print(e)
        print()
        print(
            "Make sure gmail_auth.py uses:"
        )
        print(
            'SCOPES = ["https://www.googleapis.com/auth/gmail.send"]'
        )
        print(
            "Then delete token.json and authenticate again."
        )

        return

    # --------------------------------------------------------
    # BUILD CURRENT BATCH
    # --------------------------------------------------------

    batch = []

    seen_in_batch = set()

    for index, row in df.iterrows():

        # Excel row = dataframe index + 2
        # because row 1 contains headers
        excel_row = index + 2

        hr_name = str(
            row["HR Name"]
        ).strip()

        hr_email = str(
            row["HR Email"]
        ).strip()

        company = str(
            row["Company"]
        ).strip()

        # --------------------------------------------
        # Missing values
        # --------------------------------------------

        if (
            not hr_name
            or hr_name.lower() == "nan"
        ):
            continue

        if (
            not hr_email
            or hr_email.lower() == "nan"
        ):
            continue

        if (
            not company
            or company.lower() == "nan"
        ):
            continue

        # --------------------------------------------
        # Invalid email
        # --------------------------------------------

        if not is_valid_email(hr_email):
            continue

        normalized_email = hr_email.lower()

        # --------------------------------------------
        # Already sent
        # --------------------------------------------

        if normalized_email in sent_emails:
            continue

        # --------------------------------------------
        # Duplicate email in current Excel
        # --------------------------------------------

        if normalized_email in seen_in_batch:
            continue

        seen_in_batch.add(normalized_email)

        batch.append({
            "excel_row": excel_row,
            "hr_name": hr_name,
            "hr_email": hr_email,
            "company": company
        })

        if len(batch) >= BATCH_SIZE:
            break

    # --------------------------------------------------------
    # NO NEW EMAILS
    # --------------------------------------------------------

    if not batch:

        print(
            "No new valid emails available to send."
        )

        print()
        print("=" * 80)

        return

    # --------------------------------------------------------
    # SHOW BATCH
    # --------------------------------------------------------

    print("-" * 80)
    print("CURRENT SEND BATCH")
    print("-" * 80)

    for i, item in enumerate(batch, start=1):

        print(
            f"{i:02d}. "
            f"{item['hr_name']} | "
            f"{item['hr_email']} | "
            f"{item['company']}"
        )

    print()
    print(
        f"Emails in this batch : {len(batch)}"
    )

    print()
    print("=" * 80)
    print("IMPORTANT")
    print("=" * 80)
    print()
    print(
        "These emails WILL ACTUALLY BE SENT."
    )
    print(
        "They will NOT be created as Gmail drafts."
    )
    print(
        "Resume will be attached."
    )
    print(
        "HTML/BOLD formatting will be used."
    )
    print()

    # --------------------------------------------------------
    # CONFIRMATION
    # --------------------------------------------------------

    confirmation = input(
        "Type SEND to send this batch: "
    ).strip()

    if confirmation != "SEND":

        print()
        print(
            "Sending cancelled."
        )
        print(
            "No emails were sent."
        )

        return

    print()
    print("=" * 80)
    print("STARTING SEND")
    print("=" * 80)
    print()

    successful = 0
    failed = 0

    # --------------------------------------------------------
    # SEND BATCH
    # --------------------------------------------------------

    for position, item in enumerate(
        batch,
        start=1
    ):

        excel_row = item["excel_row"]
        hr_name = item["hr_name"]
        hr_email = item["hr_email"]
        company = item["company"]

        print("-" * 80)

        print(
            f"Sending {position}/{len(batch)}"
        )

        print(
            f"Excel Row : {excel_row}"
        )

        print(
            f"HR Name   : {hr_name}"
        )

        print(
            f"HR Email  : {hr_email}"
        )

        print(
            f"Company   : {company}"
        )

        try:

            message_id = send_email(
                service,
                hr_name,
                hr_email,
                company
            )

            print(
                "[SUCCESS] Email sent"
            )

            print(
                f"Message ID : {message_id}"
            )

            log_send(
                hr_name=hr_name,
                hr_email=hr_email,
                company=company,
                excel_row=excel_row,
                status="SENT",
                details=message_id or ""
            )

            successful += 1

        except HttpError as e:

            print(
                "[ERROR] Gmail API error"
            )

            print(e)

            log_send(
                hr_name=hr_name,
                hr_email=hr_email,
                company=company,
                excel_row=excel_row,
                status="ERROR",
                details=str(e)
            )

            failed += 1

        except Exception as e:

            print(
                "[ERROR] Failed to send"
            )

            print(e)

            log_send(
                hr_name=hr_name,
                hr_email=hr_email,
                company=company,
                excel_row=excel_row,
                status="ERROR",
                details=str(e)
            )

            failed += 1

        # Delay between sends
        if position < len(batch):
            time.sleep(DELAY_SECONDS)

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    print()
    print("=" * 80)
    print("SEND BATCH COMPLETE")
    print("=" * 80)

    print()
    print(
        f"Emails attempted : {len(batch)}"
    )

    print(
        f"Successfully sent: {successful}"
    )

    print(
        f"Failed           : {failed}"
    )

    print()
    print(
        f"Send log         : {SEND_LOG_FILE}"
    )

    print()
    print("=" * 80)
    print("SAFETY INFORMATION")
    print("=" * 80)

    print(
        "Only unique email addresses were sent."
    )

    print(
        "Previously SENT addresses were skipped."
    )

    print(
        "Duplicate Excel rows were NOT deleted."
    )

    print(
        "The original Excel file was NOT modified."
    )

    print(
        "Successful sends were recorded in email_send_log.csv."
    )

    print("=" * 80)
    print()


if __name__ == "__main__":
    main()