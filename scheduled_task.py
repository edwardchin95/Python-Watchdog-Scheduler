import time
import signal
import sys
import subprocess
import os
import json
import logging
from datetime import datetime
from concurrent_log_handler import ConcurrentRotatingFileHandler

# --- CUSTOM TIMESTAMPED ROTATING HANDLER ---
class TimestampedConcurrentRotatingFileHandler(ConcurrentRotatingFileHandler):
    def doRollover(self):
        """
        Override to use timestamped filenames for rotated logs.
        """
        if self.stream:
            self.stream.close()
            self.stream = None

        # Get timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dirname, basename = os.path.split(self.baseFilename)
        new_logname = os.path.join(dirname, f"master_{timestamp}.log")

        # Rename current log to timestamped name
        if os.path.exists(self.baseFilename):
            os.rename(self.baseFilename, new_logname)

        # Remove old backups if exceeding backupCount
        # backups = sorted([
        #     f for f in os.listdir(dirname)
        #     if f.startswith("master_") and f.endswith(".log")
        # ])
        # while len(backups) > self.backupCount:
        #     os.remove(os.path.join(dirname, backups[0]))
        #     backups.pop(0)

        # Reopen the log file
        self.stream = self._open()

# --- CONFIG & LOGGER SETUP BLOCK ---
script_dir = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(script_dir, 'config.json')
log_path = os.path.join(script_dir,'logs','Master.log')

# Load config.json
try:
    with open(config_path, 'r') as f:
        config = json.load(f)
except Exception as e:
    print(f"Failed to load config: {e}")
    sys.exit(1)

interval_minutes = config.get('interval_minutes', 15)
process_script = config.get('process_script', 'PythonBPTask.py')
robocopy_script = config.get('robocopy_script', 'robocopy.py')
process_path = os.path.join(script_dir, process_script)
robocopy_path = os.path.join(script_dir, robocopy_script)

formatter = logging.Formatter('[%(asctime)s] [SCHEDULER][%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger("SCHEDULER")

# --- USE CUSTOM TIMESTAMPED HANDLER ---
file_handler = TimestampedConcurrentRotatingFileHandler(
    log_path, maxBytes=20*1024*1024, encoding='utf-8'
)
file_handler.setFormatter(formatter)
logger.setLevel(logging.INFO)
logger.addHandler(file_handler)

# (Optional) Also log to console
stream_handler = logging.StreamHandler(sys.stdout)
stream_handler.setFormatter(formatter)
logger.addHandler(stream_handler)

logger.propagate = False
# --- END CONFIG & LOGGER SETUP BLOCK ---

running = True

def graceful_exit(signum, frame):
    global running
    logger.info("Received exit signal. Shutting down gracefully...")
    running = False

# Register signal handler ONLY for SIGTERM
signal.signal(signal.SIGTERM, graceful_exit)
# signal.signal(signal.SIGINT, graceful_exit) # Uncomment if you want Ctrl+C to trigger graceful exit

def run_subprocess(script_path, script_label):
    try:
        logger.info(f"Running {script_label}...")
        result = subprocess.run([sys.executable, script_path])
        if result.returncode != 0:
            logger.warning(f"{script_label} exited with code {result.returncode}")
    except Exception as e:
        logger.error(f"Failed to run {script_label}: {e}")

def main(process_path, robocopy_path, interval_minutes):
    while running:
        run_subprocess(process_path, os.path.basename(process_path))
        run_subprocess(robocopy_path, os.path.basename(robocopy_path))
        logger.info(f"Waiting {interval_minutes} minutes before next run...")
        for i in range(interval_minutes * 60):
            if not running:
                break
            time.sleep(1)

if __name__ == "__main__":
    main(process_path, robocopy_path, interval_minutes)