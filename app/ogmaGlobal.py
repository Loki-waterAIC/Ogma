import os
import sys

APP_VERSION = "2.0.0"

# Locked access to word
LOCK_FILE_PATH: str = os.path.join(os.path.abspath("."), os.path.join("tmp", "ogma_lock.lock"))
