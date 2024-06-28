import hashlib
import time
import os

def calculate_hash(file_path):
    """Calculate the MD5 hash of a file."""
    hash_md5 = hashlib.md5()
    try:
        with open(file_path, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_md5.update(chunk)
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        return None
    return hash_md5.hexdigest()

def monitor_powerpoint_changes(file_path, log_file, interval=10):
    """Monitor a PowerPoint file for changes and log to a text file."""
    last_hash = calculate_hash(file_path)
    if last_hash is None:
        return

    print(f"Monitoring changes to {file_path}...")

    while True:
        time.sleep(interval)
        current_hash = calculate_hash(file_path)
        if current_hash is None:
            continue
        if current_hash != last_hash:
            with open(log_file, 'a') as log:
                log.write(f"File changed at {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                print(f"Change detected and logged at {time.strftime('%Y-%m-%d %H:%M:%S')}")
            last_hash = current_hash

if __name__ == "__main__":
    PPT_FILE = 'path_to_powerpoint_file.pptx'  # Update this with the path to your PowerPoint file
    LOG_FILE = 'changes_log.txt'  # Name of the log file

    if not os.path.exists(PPT_FILE):
        print(f"The specified PowerPoint file does not exist: {PPT_FILE}")
    else:
        monitor_powerpoint_changes(PPT_FILE, LOG_FILE)
