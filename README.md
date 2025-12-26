# Blood Processing Automation Suite

This suite automates blood and serum data processing, file management, and monitoring via a user-friendly dashboard. All configuration is centralized in `config.json` for easy portability.

---

## Folder Structure

MainFolder/
│
├── watchdog_dashboard.py
├── scheduled_task.py
├── robocopy.py
├── PythonBPTask.py
├── config.json
├── requirements.txt
├── README.txt
└── logs/
    └── Master.log

---

## Setup Instructions

### 1. Install Python

- Download and install Python 3.8 or newer from:  
  https://www.python.org/downloads/

### 2. Install Required Packages

- Open a command prompt in your project folder.
- Run:

pip install -r requirements.txt

### 3. Configure Paths and Settings

- Edit `config.json` to set all source/destination folders and parameters for your environment.

### 4. Create the Logs Folder

- Ensure a folder named `logs` exists in your main project directory.
- The scripts will create it automatically if missing.

---

## How to Run

### Start the Dashboard GUI

- Double-click `watchdog_dashboard.py` or run:

python watchdog_dashboard.py

- This will launch the monitoring dashboard and start the scheduler.

### Run at Startup (Optional)

- Create a shortcut or batch file to run `watchdog_dashboard.py` and place it in your Windows Startup folder.
- Example batch file:

@echo off
cd /d "%~dp0"
pythonw watchdog_dashboard.py

---

## Scripts Overview

### watchdog_dashboard.py

- GUI dashboard to monitor and control the scheduler.
- Displays live logs and allows start/stop of automated tasks.

### scheduled_task.py

- Runs data processing and file copy scripts at regular intervals, as configured in `config.json`.

### robocopy.py

- Copies new or updated files between source and destination folders, logging all actions.

### PythonBPTask.py

- Processes CSV files, formats Excel outputs, applies sensitivity labels, and logs actions.

---

## Configuration

All settings (paths, intervals, filenames, etc.) are stored in `config.json`.  
**After moving the folder to a new computer, update only `config.json` to match your new environment.**

Example `config.json`:
```json
{
"raw_file_source": "data/raw",
"bp_process_dest": "data/bp",
"serum_process_dest": "data/serum",
"bp_copy_source": "data/bp_copy_src",
"serum_copy_source": "data/serum_copy_src",
"bp_copy_dest": "data/bp_copy_dest",
"serum_copy_dest": "data/serum_copy_dest",
"interval_minutes": 5,
"record_process": "data/RecordsSim.xlsx",
"log_folder": "logs"
}




## Troubleshooting

- **Missing Packages:**  
  If you see `ModuleNotFoundError`, run `pip install -r requirements.txt` again.

- **Permission Issues:**  
  Make sure you have read/write access to all folders specified in `config.json` and to the `logs` folder.

- **GUI Not Launching:**  
  Ensure you are running the script in a user session (not as a background service).

## Customization

- To change source/destination folders or processing intervals, edit `config.json`.
- To move the project to another computer, copy the entire folder, install Python and dependencies, and update `config.json`.

## Support

For questions or issues, contact your IT administrator or the script author.