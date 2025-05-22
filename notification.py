# modules required
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from extract import data, logging
from update import insert_count, update_count, failed_syncs

# config parameters
email = data['email']
from_email = email['from']
password = email['password']
to_email= email['to']

# content of the notification
subject = "Data Sync Completed"
body = f"""
The data sync has been completed successfully.
- Items updated: {update_count}
- New items added: {insert_count}
- Failed syncs: {failed_syncs}

"""

try:
    logging.info("Preparing email message.")
    msg = MIMEMultipart("alternative")
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    logging.info("Connecting to SMTP server...")
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    logging.info("Logging into email account...")
    server.login(from_email, password)

    logging.info("Sending email...")
    server.sendmail(from_email, to_email, msg.as_string())
    server.quit()
    logging.info("Email sent successfully!")

except smtplib.SMTPAuthenticationError:
    logging.info("Failed to authenticate. Check your email/password or app-specific password.")
except smtplib.SMTPConnectError:
    logging.info("Failed to connect to the SMTP server.")
except smtplib.SMTPRecipientsRefused:
    logging.info("Recipient address was refused by the server.")
except smtplib.SMTPException as e:
    logging.info(f"SMTP error occurred: {e}")
except Exception as e:
    logging.info(f"An unexpected error occurred: {e}")


