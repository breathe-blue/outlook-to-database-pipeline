
# 📧 Outlook to Database Pipeline

This repository automates the workflow of **extracting Excel/CSV attachments from Outlook emails**, **processing them**, and **syncing the data into a database**. It also sends a **notification email** upon completion. This is ideal for environments where data is sent regularly via email and needs to be updated in a database without manual intervention.

<br>

## 🔁 Workflow Overview

### 1. **Email Processing**
- Connects to **Microsoft Outlook Inbox**
- Filters emails based on **sender** and **subject line**
- Uses `latest.txt` to process only those emails sent after the last run
- Downloads attachments to the `file_downloads` folder that gets created automatically

### 2. **Data Sync**
- Processes and reads all the downloaded Excel/CSV files
- Loads data into **dataframes** and cleans it
- Determines rows to **insert** or **update** based on the `id_field`
- Syns the data based on the appropriate tables in the database.
- Tracks the count of **inserts**, **updates**, and **failed syncs** if any

### 3. **Notification**
- Sends an email to a designated recipient after processing from a given email address.
- Email contains a summary of insert count, update count, and failure count which can be changed as per requirements  
<br>

## 📁 Repository Structure

| Files               | Description |
|---------------------|-------------|
| `requirements.txt`  | Contains all Python dependencies |
| `config.json`       | Dynamic parameters: email filters, database info, etc. |
| `extract.py`        | Extracts filtered Outlook emails and downloads attachments |
| `update.py`         | Parses the data files and updates the database |
| `notification.py`   | Sends notification email with sync summary |
| `latest.txt`        | Timestamp log to avoid reprocessing emails |
| `__init__.py`       | Initializes the script environment if needed and has the version |

<br>

## 💻 System Requirements

- **OS**: Windows (due to `pywin32` for Outlook integration)
- **Python**: 3.7 or later
- **Email**: Microsoft Outlook (installed and configured)
<br>


## 🔧 Installation & Setup

### 1. Clone the Repository
```bash
git clone https://github.com/your-username/outlook-to-database-pipeline.git
cd outlook-to-database-pipeline
```

### 2. Install Required Packages
```bash
pip install -r requirements.txt
```

### 3. Configure the Pipeline
Edit the `config.json` file to set:
- Email sender and subject filters
- Download and processing folders
- Database credentials and target tables
- Notification email details

### 4. Run the Script
After updating the configuration, simply run extract.py. If everything is set up correctly, it will automatically trigger update.py and notification.py using subprocesses.
<br>

📌 **Note**:
- Run scripts as Administrator for Outlook access and file permissions.
- Can be automated using task schedulers
<br>

<br>

## 🤝 Contributions

Feel free to fork this repo, customize it for your organization, and open pull requests for improvements or fixes.

---
MIT License

Copyright (c) 2025

Permission is hereby granted, free of charge, to any person obtaining a copy  
of this software and associated documentation files (the "Software"), to deal  
in the Software without restriction, including without limitation the rights  
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell  
copies of the Software, and to permit persons to whom the Software is  
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in  
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR  
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,  
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE  
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER  
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,  
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN  
THE SOFTWARE.
