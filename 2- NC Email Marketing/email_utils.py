from datetime import datetime, timedelta
from email.message import EmailMessage
import email as email_parser
import imaplib
import logging
import os
from pathlib import Path
import random
import re
import smtplib
import time
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

SENDER_EMAIL = os.getenv("GMAIL_USER2")
APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD2")
CC_EMAIL = "info@nirmanshilaconstruction.com"

EXCEL_FILE_PATH = os.getenv("email")
ATTACHMENT_PATH = os.getenv("Brochure")

SUBJECT = (
    "Slipform Construction Expert for RCC Chimney & Tall Structures |"
    " Nirmanshila Construction"
)


def get_html_body():
    """Returns the HTML body formatted for bulk emails."""
    return f"""\
<html>
  <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333333;">
    <p>Dear Sir,</p>

    <p>Hope this email finds you well.</p>

    <p>We are writing to introduce <strong>Nirmanshila Construction</strong> — an expert partner for your slipform construction requirements.</p>

    <p>If you are planning new expansions, modernizations, or structural maintenance, we offer end-to-end execution capabilities across the following core areas:</p>

    <ul>
      <li><strong>Tall Structural Construction:</strong> RCC Chimney, Silos, and Elevated Overhead Water Tanks.</li>
      <li><strong>Slipform Expertise:</strong> Experienced team for slipform work.</li>
      <li><strong>Thermal & Asset Protection:</strong> Refractory Brick Lining, Industrial Painting, and Protective Coatings.</li>
      <li><strong>RCC Chimney Repair/Maintenance and General Industrial Civil Works</strong>.</li>
    </ul>

    <p><strong>Why Partner with Nirmanshila?</strong></p>
    <ul>
      <li><strong>Technical Precision:</strong> From continuous-pour slipform setups to strict structural tolerances, we ensure flawless execution.</li>
      <li><strong>Operational Integrity:</strong> We pride ourselves on maintaining strict project timelines, safety standards, and transparent budgeting.</li>
      <li><strong>Turnkey Management:</strong> We handle the engineering complexities so your team can focus on core operations.</li>
    </ul>

    <p>For your review, I have attached our <strong>Company Brochure that includes company profile & project portfolio</strong>. You can also explore our capabilities online at <a href="https://www.nirmanshilaconstruction.com">www.nirmanshilaconstruction.com</a>.</p>

    <p>We would welcome the opportunity to connect to explore how we can support your upcoming projects. Let us know a good time to talk.</p>

    <p>Looking forward to hearing from you. Thanks!</p>

    <br>
    <p><strong>Rajeev Kumar</strong> &nbsp;|&nbsp; <strong>Amit Tayal</strong><br>
    +91 7500462001 &nbsp;|&nbsp; +91 9650744299</p>

    <hr style="border: none; border-top: 1px solid #cccccc; margin: 20px 0;">
    <p style="font-size: 0.9em; color: #555555;">
      <strong>M/s Nirmanshila Construction</strong><br>
      Registered Office: 17, Kushi Vihar, Shanti Nagar, Muzaffarnagar, 251001<br>
      Website: <a href="https://www.nirmanshilaconstruction.com">www.nirmanshilaconstruction.com</a><br>
      Email: info@nirmanshilaconstruction.com
    </p>
  </body>
</html>
"""


def create_batch_email(bcc_emails, pdf_data, pdf_name):
    """Creates a single EmailMessage targeting a batch of BCC recipients with CC included."""
    msg = EmailMessage()
    msg["Subject"] = SUBJECT
    msg["From"] = f"Nirmanshila Construction <{SENDER_EMAIL}>"
    msg["To"] = SENDER_EMAIL
    msg["Cc"] = CC_EMAIL
    msg["Bcc"] = ", ".join(bcc_emails)

    plain_text = "Dear Sir,\n\nPlease view this email in an HTML-compatible client."
    msg.set_content(plain_text)
    msg.add_alternative(get_html_body(), subtype="html")

    if pdf_data:
        msg.add_attachment(
            pdf_data,
            maintype="application",
            subtype="pdf",
            filename=pdf_name,
        )

    return msg


def send_bulk_emails_safely(excel_file_path=None, attachment_path=None, batch_size=100):
    """Saves a single draft email via IMAP containing up to batch_size eligible BCC recipients."""
    print("--- Starting Bulk Email Draft Creation Process ---")
    
    target_excel = Path(excel_file_path or EXCEL_FILE_PATH)
    target_attach = Path(attachment_path or ATTACHMENT_PATH)

    print(f"Target Excel Path: {target_excel}")
    print(f"Target Attachment Path: {target_attach}")

    if not target_excel.is_file():
        error_msg = f"Could not find Excel file at: {target_excel}"
        print(f"Error: {error_msg}")
        raise FileNotFoundError(error_msg)

    print("Reading Excel file...")
    df = pd.read_excel(target_excel)
    print(f"Total rows found in Excel: {len(df)}")

    if "Last Sent Date" not in df.columns:
        df["Last Sent Date"] = None
    

    if target_attach.is_file():
        print(f"Attachment '{target_attach.name}' loaded successfully.")
        pdf_data = target_attach.read_bytes()
        pdf_name = target_attach.name
    else:
        print("No attachment found or attachment path is invalid. Continuing without attachment.")
        pdf_data = None
        pdf_name = None

    now = datetime.now()
    cooldown_period = timedelta(days=7)

    eligible_items = []
    dnd_count = 0
    cooldown_count = 0

    print("Filtering eligible emails based on criteria (No DND, sent >= 7 days ago)...")
    for idx, row in df.iterrows():
        email = str(row["Email"]).strip() if pd.notna(row.get("Email")) else ""

        if not email or "@" not in email:
            continue

        # Skip rows marked as DND in 'Email ID Status'
        email_status = (
            str(row["Email ID Status"]).strip().upper()
            if "Email ID Status" in df.columns and pd.notna(row.get("Email ID Status"))
            else ""
        )
        if email_status == "DND":
            dnd_count += 1
            continue

        # Pick only emails sent 7 or more days ago (or never sent)
        last_sent_val = row["Last Sent Date"]
        if pd.notna(last_sent_val):
            try:
                last_sent_dt = pd.to_datetime(last_sent_val)
                if now - last_sent_dt < cooldown_period:
                    cooldown_count += 1
                    continue
            except Exception as e:
                pass

        eligible_items.append((idx, email))

    print(f"Filtering complete:")
    print(f" - Skipped (DND): {dnd_count}")
    print(f" - Skipped (Sent within last 7 days): {cooldown_count}")
    print(f" - Total Eligible Emails: {len(eligible_items)}")

    if not eligible_items:
        print("No eligible emails found to process. Exiting.")
        return

    # Select only the first batch of eligible emails
    first_batch = eligible_items[:batch_size]
    batch_indices = [item[0] for item in first_batch]
    batch_emails = [item[1] for item in first_batch]

    failed_batches = []

    try:
        print("Connecting to Gmail IMAP server (imap.gmail.com)...")
        with imaplib.IMAP4_SSL("imap.gmail.com") as imap:
            imap.login(SENDER_EMAIL, APP_PASSWORD)
            print("Successfully logged into Gmail via IMAP.")

            print(f"\nProcessing single draft ({len(batch_emails)} recipients)...")

            msg = create_batch_email(batch_emails, pdf_data, pdf_name)
            timestamp_str = now.strftime("%Y-%m-%d")

            try:
                # Append email to [Gmail]/Drafts folder
                print("Saving single draft to '[Gmail]/Drafts'...")
                imap.append(
                    "[Gmail]/Drafts",
                    "\\Draft",
                    imaplib.Time2Internaldate(time.time()),
                    msg.as_bytes(),
                )

                # Update Excel status upon saving draft
                print("Draft saved successfully. Updating Excel statuses to 'Sent' and current date...")
                for idx in batch_indices:
                    df.at[idx, "Last Sent Date"] = timestamp_str
                df.to_excel(target_excel, index=False)
                print("Excel file updated.")

            except Exception as err:
                print(f"Failed to save draft: {err}")
                for idx in batch_indices:
                    df.at[idx, "Last Sent Date"] = f"Failed: {err}"
                df.to_excel(target_excel, index=False)

                failed_batches.append(1)

    except Exception as e:
        print(f"An IMAP/Connection error occurred: {e}")

    print("\n--- Process Finished ---")
    if failed_batches:
        print("The draft creation encountered an error.")
    else:
        print("Draft saved to Drafts successfully!")