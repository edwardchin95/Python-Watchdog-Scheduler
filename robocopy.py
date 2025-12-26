import shutil
import os
import json
import sys
import logging
from datetime import datetime
from concurrent_log_handler import ConcurrentRotatingFileHandler

# --- CUSTOM TIMESTAMPED ROTATING HANDLER ---
class TimestampedConcurrentRotatingFileHandler(ConcurrentRotatingFileHandler):
    def doRollover(self):
        if self.stream:
            self.stream.close()
            self.stream = None

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dirname, basename = os.path.split(self.baseFilename)
        new_logname = os.path.join(dirname, f"master_{timestamp}.log")

        if os.path.exists(self.baseFilename):
            os.rename(self.baseFilename, new_logname)

        self.stream = self._open()

# --- LOGGER SETUP BLOCK ---
script_dir = os.path.dirname(os.path.abspath(__file__))
log_path = os.path.join(script_dir,'logs','Master.log')
formatter = logging.Formatter('[%(asctime)s] [ROBOCOPY][%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger("ROBOCOPY")

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
# --- END LOGGER SETUP BLOCK ---

# Get the directory where the script is located
config_path = os.path.join(script_dir, 'config.json')

# Load config.json
try:
    with open(config_path, 'r') as f:
        config = json.load(f)
except Exception as e:
    logger.error(f"Failed to load config: {e}")
    sys.exit(1)

bp_copy_source = config['bp_copy_source']
serum_copy_source = config['serum_copy_source']
bp_copy_dest = config['bp_copy_dest']
serum_copy_dest = config['serum_copy_dest']

# --- Copy missing file or file with outdated timestamp ---
def copy_missing_or_updated_files(source, destination):
    copied_count = 0
    if not os.path.exists(source):
        logger.warning(f"Source directory missing: {source}")
        return copied_count
    for root, dirs, files in os.walk(source):
        rel_path = os.path.relpath(root, source)
        rel_path = "" if rel_path == "." else rel_path
        dest_dir = os.path.normpath(os.path.join(destination, rel_path))
        os.makedirs(dest_dir, exist_ok=True)
        for file in files:
            src_file = os.path.join(root, file)
            dest_file = os.path.join(dest_dir, file)
            try:
                if not os.path.exists(dest_file):
                    shutil.copy2(src_file, dest_file)
                    logger.info(f"Copied new file: {src_file} -> {dest_file}")
                    copied_count += 1
                else:
                    if os.path.getmtime(src_file) > os.path.getmtime(dest_file):
                        shutil.copy2(src_file, dest_file)
                        logger.info(f"Updated file: {src_file} -> {dest_file}")
                        copied_count += 1
            except Exception as e:
                logger.error(f"Failed to copy {src_file} to {dest_file}: {e}")
    return copied_count

def main():
    total_copied = 0
    total_copied += copy_missing_or_updated_files(bp_copy_source, bp_copy_dest)
    total_copied += copy_missing_or_updated_files(serum_copy_source, serum_copy_dest)
    if total_copied == 0:
        logger.info("Nothing new to copy over.")

if __name__ == "__main__":
    main()