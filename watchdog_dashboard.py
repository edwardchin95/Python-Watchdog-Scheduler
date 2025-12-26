import tkinter as tk
from tkinter import scrolledtext
import subprocess
import threading
import os
import queue
import json
import logging
from datetime import datetime
import sys
import re
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

        # --- Backup deletion logic removed ---
        # All timestamped logs will be kept

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

scheduler_script_path = os.path.join(script_dir, config.get('scheduler_script', 'scheduled_task.py'))
log_file_path = log_path  # Use Master.log for consistency

formatter = logging.Formatter('[%(asctime)s] [WATCHDOG][%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
logger = logging.getLogger("WATCHDOG")

# --- USE CUSTOM TIMESTAMPED HANDLER ---
file_handler = TimestampedConcurrentRotatingFileHandler(
    log_file_path, maxBytes=20*1024*1024, encoding='utf-8'
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

# --- REGEX FOR TIMESTAMP DETECTION ---
TIMESTAMP_REGEX = r'^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\]'

# --- GUI creation ---
class SchedulerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Blood Processing Formatting Task Monitor")
        self.root.geometry("600x400")
        # GUI Components
        self.status_label = tk.Label(root, text="Scheduler Status: Not Running", fg="red", font=("Arial", 12))
        self.status_label.pack(pady=10)
        self.start_button = tk.Button(root, text="Start Scheduler", command=self.start_scheduler)
        self.start_button.pack(pady=5)
        self.stop_button = tk.Button(root, text="Stop Scheduler", command=self.stop_scheduler, state=tk.DISABLED)
        self.stop_button.pack(pady=5)
        self.output_box = scrolledtext.ScrolledText(root, wrap=tk.WORD, height=15, width=70, state=tk.DISABLED)
        self.output_box.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        # State
        self.process = None
        self.running = False
        self.user_stopped = False
        self.output_queue = queue.Queue()
        self.lines_buffer = []  # Store last 30 lines
        # Periodic check for output
        self.root.after(500, self.update_output)

# --- GUI start scheduler button---
    def start_scheduler(self):
        if not self.running:
            try:
                self.running = True
                self.user_stopped = False  # Reset flag
                self.process = subprocess.Popen(
                    ["python", "-u", scheduler_script_path],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1
                )
                self.status_label.config(text="Scheduler Status: Running", fg="green")
                self.start_button.config(state=tk.DISABLED)
                self.stop_button.config(state=tk.NORMAL)
                threading.Thread(target=self.monitor_process, daemon=True).start()
                threading.Thread(target=self.read_output, daemon=True).start()
                logger.info("Scheduler started.")
            except Exception as e:
                self.append_output(f"Error starting scheduler: {e}\n")
                logger.error(f"Error starting scheduler: {e}")

# --- GUI stop scheduler button---
    def stop_scheduler(self):
        if self.process and self.running:
            self.user_stopped = True  # Mark user stop
            try:
                self.process.terminate()
                self.process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self.process.kill()
            except Exception as e:
                logger.error(f"Error stopping scheduler: {e}")
            finally:
                self.running = False
                self.status_label.config(text="Scheduler Status: Stopped", fg="red")
                self.start_button.config(state=tk.NORMAL)
                self.stop_button.config(state=tk.DISABLED)
                logger.info("Scheduler stopped.")

# --- If crashed, will reboot after 10 seconds---
    def monitor_process(self):
        while self.running:
            if self.process.poll() is not None:
                if not self.user_stopped:  # Only restart if NOT user stopped
                    self.output_queue.put("Scheduler crashed. Restarting in 10 seconds...\n")
                    self.status_label.config(text="Scheduler Status: Crashed - Restarting", fg="orange")
                    self.running = False
                    logger.warning("Scheduler crashed. Restarting in 10 seconds...")
                    self.root.after(10000, self.start_scheduler)
                break

    def read_output(self):
        # Only display output in GUI, do NOT log child process output to Master.log
        for line in self.process.stdout:
            self.output_queue.put(line)

# --- GUI log output---
    def update_output(self):
        while not self.output_queue.empty():
            raw_text = self.output_queue.get().rstrip('\n')
            # --- Only add timestamp if not already present ---
            if re.match(TIMESTAMP_REGEX, raw_text):
                formatted_text = f"{raw_text}\n"
            else:
                timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
                formatted_text = f"{timestamp} {raw_text}\n"
            # Maintain last 30 lines
            self.lines_buffer.append(formatted_text)
            if len(self.lines_buffer) > 30:
                self.lines_buffer.pop(0)
            # Update GUI
            self.output_box.config(state=tk.NORMAL)
            self.output_box.delete(1.0, tk.END)  # Clear current text
            self.output_box.insert(tk.END, "".join(self.lines_buffer))
            self.output_box.see(tk.END)
            self.output_box.config(state=tk.DISABLED)
        self.root.after(500, self.update_output)

    def append_output(self, text):
        self.output_box.config(state=tk.NORMAL)
        self.output_box.insert(tk.END, text)
        self.output_box.see(tk.END)
        self.output_box.config(state=tk.DISABLED)

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = SchedulerGUI(root)
        app.start_scheduler()  # Start immediately on launch
        logger.info("Watchdog GUI started.")
        root.mainloop()
    except Exception as e:
        logger.error(f"Watchdog GUI failed to start: {e}")