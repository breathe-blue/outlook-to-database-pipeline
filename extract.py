import os, subprocess, re, datetime, win32com.client, json, logging, shutil
from pathlib import Path
from win32com.client import gencache
from os.path import dirname, abspath, join, exists, isfile, islink, isdir
from datetime import datetime
from os import remove, listdir, unlink


main_script = dirname(abspath(__file__))
log_file = join(main_script, "script.log")

logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

config_path = join(main_script, "config.json")

with open(config_path) as config_file:
    data = json.load(config_file)

filter = data['filter']
sender = filter['sender']
subject = filter['subject']
file_dir = data['file_dir']

latest = join(main_script, "latest.txt")

if not exists(latest):
    with open(latest, 'w') as f:
        f.write("")

def create_directory(base_path, folder_name):
    try:
        new_dir = base_path / folder_name
        new_dir.mkdir(parents=True, exist_ok=True)
        return new_dir
    except Exception as e:
        logging.error(f"Error creating directory {folder_name}: {str(e)}")
        return None

def save_attachments(attachments, file_dir, email_time):
    download_path = join(file_dir, "file_downloads")

    if exists(download_path):
        for filename in listdir(download_path):
            file_path = join(download_path, filename)
            try:
                if isfile(file_path) or islink(file_path):
                    unlink(file_path)
                elif isdir(file_path):
                    shutil.rmtree(file_path)       
                logging.info(f"Successfully deleted : {file_path}")
            except Exception as e:
                logging.error(f"Failed to delete {file_path}. Reason: {e}")
    else:
        os.makedirs(download_path)

    counter = 1
    for attachment in attachments:
        try:
            attachment_name = re.sub(r'[^\w\s.]+', '', attachment.FileName)
            base_name, ext = os.path.splitext(attachment_name)

            if ext.lower() in [".csv", ".xlsx", ".xls"]:
                timestamped_name = f"{email_time}_{counter}{ext}"
                attachment_path = join(download_path, timestamped_name)
                counter += 1
            else:
                attachment_path = join(file_dir, attachment_name)

            attachment.SaveAsFile(attachment_path)
            logging.info(f"Saved attachment: {attachment_path}")

        except Exception as e:
            logging.error(f"Failed to save attachment {attachment.FileName}: {str(e)}")

def export_emails(folder, file_dir):
    messages = folder.Items
    try:
        messages.Sort("[CreationTime]", True)
    except Exception as e:
        logging.warning(f"Could not sort messages: {str(e)}")

    processed_count = 0
    last_exe_time = None

    if os.path.exists(latest):
        with open(latest, 'r') as f:
            last_exe_str = f.read().strip()
            if last_exe_str:
                last_exe_time = datetime.strptime(last_exe_str, "%Y-%m-%d_%H-%M-%S")
                logging.info(f"Last downloaded time: {last_exe_time}")

    logging.info(f"Number of messages in folder: {len(messages)}")

    latest_processed_time = last_exe_time

    for message in messages:
        try:
            if message.Class != win32com.client.constants.olMail:
                continue

            try:
                if message.SenderEmailType == "EX":
                    sender_raw = message.Sender.GetExchangeUser().PrimarySmtpAddress
                else:
                    sender_raw = message.SenderEmailAddress
            except Exception:
                sender_raw = message.SenderEmailAddress

            match = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', sender_raw)
            if match:
                sender_extracted = match.group(1).lower()
            else:
                sender_extracted = sender_raw.lower()

            if sender:
                sender_expected = sender.lower()
                logging.info(f"Expected sender: {sender_expected}")
                logging.info(f"Actual extracted sender: {sender_extracted}")
                if sender_extracted != sender_expected:
                    logging.info(f"Skipping due to sender mismatch: expected '{sender_expected}', got '{sender_extracted}'")
                    continue

            logging.info(f"Subject: {message.Subject}")
            if subject and subject.lower() not in message.Subject.lower():
                logging.info(f"Skipping due to subject mismatch: expected to contain '{subject.lower()}'")
                continue

            email_time = message.CreationTime.strftime("%Y-%m-%d_%H-%M-%S")
            logging.info(f"Generated email time: {email_time}")
            email_time_obj = datetime.strptime(email_time, "%Y-%m-%d_%H-%M-%S")

            if last_exe_time and email_time_obj <= last_exe_time:
                logging.info(f"Skipping email {message.Subject} as it was downloaded before {last_exe_time}")
                continue

            if message.Attachments.Count > 0:
                save_attachments(message.Attachments, file_dir, email_time)

            processed_count += 1

            if not latest_processed_time or email_time_obj > latest_processed_time:
                latest_processed_time = email_time_obj

        except Exception as e:
            logging.error(f"Error processing email: {str(e)}")


    if latest_processed_time:
        with open(latest, 'w') as f:
            f.write(latest_processed_time.strftime("%Y-%m-%d_%H-%M-%S"))

    return processed_count

def main():
    logging.info("Starting the script...")
    outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    base_dir = Path.cwd() / "EmailExports"
    base_dir.mkdir(parents=True, exist_ok=True)
    logging.info(f"base_dir: {base_dir}")

    all_folders = namespace.Folders
    total_processed = 0

    for store in all_folders:
        folders = store.Folders
        logging.info(f"folders: {folders}")

        for folder in folders:
            if folder.Name.lower() == "inbox":
                processed_count = export_emails(folder, file_dir)
                total_processed += processed_count

    logging.info(f"Total emails processed: {total_processed}")
    logging.info("Email export completed.")

    try:
        if base_dir:
            shutil.rmtree(base_dir)
            logging.info(f"Deleted original EmailExports folder at {base_dir}")
    except Exception as e:
        logging.error(f"Error deleting EmailExports folder: {str(e)}")
    
    
    try:
        script_path = Path(__file__).parent / "update.py"
        subprocess.run(["python", str(script_path)], check=True)
        logging.info(f"Successfully executed {script_path}")
    except Exception as e:
        logging.error(f"Error executing update.py: {str(e)}")


if __name__ == "__main__":
    main()
