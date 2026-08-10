import email
from email.header import decode_header
import imaplib
import os
from pathlib import Path
import re
import pandas as pd
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# IMAP Configuration
IMAP_SERVER = "imap.gmail.com"
IMAP_PORT = 993


def parse_header_val(header_value):
    """Safely decodes encoded email headers (From, Subject, etc.)."""
    if not header_value:
        return ""
    decoded = decode_header(header_value)
    header_str = ""
    for text, encoding in decoded:
        if isinstance(text, bytes):
            header_str += text.decode(encoding or "utf-8", errors="ignore")
        else:
            header_str += str(text)
    return header_str


def update_excel_with_bounces(
    excel_file_path: str = None,
    sender_email: str = None,
    app_password: str = None,
):
    """Connects to Gmail via IMAP, processes inbox messages to:

    1. Set 'Email ID Status' to 'DND' for bounced/undelivered email addresses.
    2. Capture alternative email addresses from recipient replies into 'Status'.
    3. Delete inbox messages sent from or associated with DND email addresses.
    """
    print("[INFO] Starting email processing...")

    # Fallback to environment variables if parameters are not explicitly passed
    excel_file_path = excel_file_path or os.getenv("emailMDB")
    sender_email = sender_email or os.getenv("GMAIL_USER2")
    app_password = app_password or os.getenv("GMAIL_APP_PASSWORD2")

    if not excel_file_path or not sender_email or not app_password:
        print("[ERROR] Missing required credentials or Excel file path.")
        return

    target_excel = Path(excel_file_path)

    if not target_excel.is_file():
        print(f"[ERROR] Specified Excel file does not exist: {target_excel}")
        return

    print(f"[INFO] Reading Excel file: {target_excel}")
    df = pd.read_excel(target_excel)

    if "Status" not in df.columns:
        df["Status"] = None
    if "Email ID Status" not in df.columns:
        df["Email ID Status"] = None

    # Map normalized emails to row indices for lookup
    excel_email_map = {}
    for idx, row in df.iterrows():
        email_val = (
            str(row["Email"]).strip().lower()
            if pd.notna(row.get("Email"))
            else ""
        )
        if email_val and "@" in email_val:
            excel_email_map[email_val] = idx

    if not excel_email_map:
        print("[WARNING] No valid emails found in the Excel sheet.")
        return

    print(f"[INFO] Loaded {len(excel_email_map)} emails from Excel.")

    # Load pre-existing DND emails from Excel
    dnd_emails = set(
        df[df["Email ID Status"].astype(str).str.strip().str.upper() == "DND"]["Email"]
        .dropna()
        .astype(str)
        .str.strip()
        .str.lower()
    )

    alt_email_matches = {}

    email_regex = re.compile(
        r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
    )

    bounce_indicators = [
        "mailer-daemon",
        "postmaster",
        "mail delivery",
        "delivery status notification",
        "undelivered",
        "failure notice",
        "returned to sender",
        "mail administrator",
        "system administrator",
    ]

    try:
        print(f"[INFO] Connecting to Gmail IMAP ({IMAP_SERVER}:{IMAP_PORT}) for {sender_email}...")
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(sender_email, app_password)
        mail.select("INBOX")
        print("[INFO] Connected successfully. Searching INBOX...")

        status, response = mail.search(None, "ALL")

        if status != "OK" or not response[0]:
            print("[INFO] No messages found in INBOX.")
            mail.logout()
            return

        msg_ids = response[0].split()
        print(f"[INFO] Found {len(msg_ids)} messages in INBOX. Parsing messages...")

        # Store metadata to determine deletions after analyzing all messages
        msg_records = []

        for msg_id in msg_ids:
            status, msg_data = mail.fetch(msg_id, "(RFC822)")
            if status != "OK":
                continue

            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])

                    sender_from = parse_header_val(msg.get("From", "")).lower()
                    subject = parse_header_val(msg.get("Subject", "")).lower()

                    body_text = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            if content_type in [
                                "text/plain",
                                "text/html",
                                "message/delivery-status",
                            ]:
                                try:
                                    payload = part.get_payload(decode=True)
                                    if payload:
                                        body_text += (
                                            payload.decode(
                                                "utf-8", errors="ignore"
                                            )
                                            + "\n"
                                        )
                                except Exception:
                                    pass
                    else:
                        payload = msg.get_payload(decode=True)
                        if payload:
                            body_text = payload.decode("utf-8", errors="ignore")

                    is_bounce = any(
                        indicator in sender_from or indicator in subject
                        for indicator in bounce_indicators
                    )

                    extracted_emails = [
                        e.strip().lower() for e in email_regex.findall(body_text)
                    ]
                    sender_matches = email_regex.findall(sender_from)
                    sender_clean = sender_matches[0].strip().lower() if sender_matches else ""

                    if is_bounce:
                        for found_clean in extracted_emails:
                            if (
                                found_clean != sender_email.lower()
                                and found_clean in excel_email_map
                            ):
                                dnd_emails.add(found_clean)

                    else:
                        if sender_clean in excel_email_map:
                            for found_clean in extracted_emails:
                                if (
                                    found_clean != sender_clean
                                    and found_clean != sender_email.lower()
                                    and not any(
                                        domain in found_clean
                                        for domain in [
                                            "google.com",
                                            "gmail.com",
                                            "reply",
                                            "support",
                                        ]
                                    )
                                ):
                                    alt_email_matches[sender_clean] = (
                                        found_clean
                                    )
                                    break

                    msg_records.append({
                        "msg_id": msg_id,
                        "is_bounce": is_bounce,
                        "sender_clean": sender_clean,
                        "extracted_emails": set(extracted_emails),
                    })

        # Process IMAP deletions for bounce messages and messages associated with DND emails
        print("[INFO] Processing email deletions...")
        deleted_count = 0
        for record in msg_records:
            is_dnd_sender = record["sender_clean"] in dnd_emails
            is_dnd_bounce = record["is_bounce"] or bool(
                record["extracted_emails"].intersection(dnd_emails)
            )

            if is_dnd_sender or is_dnd_bounce:
                try:
                    mail.copy(record["msg_id"], "[Gmail]/Trash")
                except Exception:
                    pass
                mail.store(record["msg_id"], "+FLAGS", "\\Deleted")
                deleted_count += 1

        if deleted_count > 0:
            print(f"[INFO] Permanently expunging {deleted_count} message(s) from INBOX...")
            mail.expunge()
            print(f"[SUCCESS] Deleted {deleted_count} bounce/DND email(s).")
        else:
            print("[INFO] No bounce or DND emails found to delete.")

        mail.logout()
        print("[INFO] IMAP session closed successfully.")

    except Exception as e:
        print(f"[ERROR] An error occurred during execution: {e}")
        return

    # Update DataFrame and write back to Excel file
    print("[INFO] Updating Excel dataset...")
    dnd_count = 0
    alt_count = 0

    for email_addr, row_idx in excel_email_map.items():
        if email_addr in dnd_emails:
            df.at[row_idx, "Email ID Status"] = "DND"
            dnd_count += 1

        if email_addr in alt_email_matches:
            df.at[row_idx, "Status"] = alt_email_matches[email_addr]
            alt_count += 1

    if dnd_count > 0 or alt_count > 0:
        df.to_excel(target_excel, index=False)
        print(f"[SUCCESS] Excel file updated successfully with {dnd_count} DND statuses and {alt_count} alternate emails.")
    else:
        print("[INFO] No Excel updates required.")

    print("[SUCCESS] Execution completed.")