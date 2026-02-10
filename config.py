import os

# URLs
BASE_URL = "https://digital-ons.bizagi.com"

# Settings
HEADLESS = False  # Must be False for manual login
TIMEOUT = 30000   # 30 seconds default timeout

# Paths
PROCESS_DIR = os.path.dirname(os.path.abspath(__file__))

# User requested: C:\Users\CLIENTE\Downloads (Donwloads typo fixed)
# Robust way:
try:
    # DOWNLOAD_DIR = os.path.join(os.path.expanduser("~"), "Downloads")
    # Forcing workspace directory to avoid permission issues
    DOWNLOAD_DIR = os.path.join(PROCESS_DIR, "downloads")
except:
    DOWNLOAD_DIR = os.path.join(PROCESS_DIR, "downloads")

# Ensure download directory exists
os.makedirs(DOWNLOAD_DIR, exist_ok=True)
