import hashlib
import time
import os

def calculate_hash(file_path):
    """Calculate the MD5 hash of a file."""
    hash_md5 = hashlib.md5()
    with open(file_path, 'rb') as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()

def monitor_powerpoint_changes(file_path, log_file):
    """Monitor a PowerPoint file for changes and log to a text file."""
    last_hash = calculate_hash(file_path)
    while True:
        time.sleep(10)  # wait for 10 seconds before checking again
        current_hash = calculate_hash(file_path)
        if current_hash != last_hash:
            with open(log_file, 'a') as log:
                log.write(f"File changed at {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            last_hash = current_hash

if __name__ == "__main__":
    PPT_FILE = 'path_to_powerpoint_file.pptx'  # Update this with the path to your PowerPoint file
    LOG_FILE = 'changes_log.txt'  # Name of the log file
    monitor_powerpoint_changes(PPT_FILE, LOG_FILE)
