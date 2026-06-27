"""Handles the 'successful login' email notification."""
import os
import smtplib
from datetime import datetime
from email.mime.text import MIMEText

def send_login_alert(to_email, username):
    """Sends a 'successful login' email to the address used to log in.

    Returns (sent: bool, error: str | None).
    """
    email_address = os.environ.get("EMAIL_ADDRESS")
    email_password = os.environ.get("EMAIL_PASSWORD")
    if not email_address or not email_password:
        return False, "Email credentials are not set in .env"
    if not to_email:
        return False, "No email address on file for this user"

    timestamp = datetime.now().strftime("%d %b %Y, %I:%M %p")
    body = (
        f"Hi {username},\n\n"
        f"This confirms a successful login to your Employee Management "
        f"System account on {timestamp}.\n\n"
        f"If this wasn't you, please change your password immediately.\n\n"
        f"Regards,\nEmployee Management System"
    )
    msg = MIMEText(body)
    msg["Subject"] = "Successful Login Alert - Employee Management System"
    msg["From"] = email_address
    msg["To"] = to_email

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=10) as smtp:
            smtp.login(email_address, email_password)
            smtp.sendmail(email_address, to_email, msg.as_string())
        return True, None
    except Exception as e:
        return False, str(e)
