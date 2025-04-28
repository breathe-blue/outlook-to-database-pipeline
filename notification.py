import pandas as pd
from arcgis.gis import GIS
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from os.path import dirname, abspath, join
from extract import data, logging


email = data['email']
from_email = email['from']
password = email['password']
to_email= email['to']


def send_email(updated_count, added_count, failed_syncs, from_email, password, to_email):


    subject = "Data Sync Completed"
    body = f"""
    The data sync has been completed successfully.
    - Items updated: {updated_count}
    - New items added: {added_count}
    - Failed syncs: {len(failed_syncs)}

    Failed syncs: {', '.join(failed_syncs)}
    """

    try:
        logging.info("Preparing email message...")
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

send_email(updated_count, added_count, failed_syncs)

