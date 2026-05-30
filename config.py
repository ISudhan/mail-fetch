"""
config.py
---------
Central configuration for the Outlook Attachment Automation System (Linux/IMAP).
Edit ONLY this file to adapt the system to your environment.
"""

# ---------------------------------------------------------------------------
# IMAP Connection — OAuth2 / XOAUTH2 (no password needed)
# ---------------------------------------------------------------------------
IMAP_HOST: str = "imap-mail.outlook.com"   # Personal Outlook.com / Hotmail
IMAP_PORT: int = 993
IMAP_USE_SSL: bool = True

IMAP_USERNAME: str = "nsudhan@outlook.com"   # Your full Outlook email

# OAuth2 — uses Microsoft public client (no Azure app registration required)
# First run will open a browser for one-time login; token is cached after that.
OAUTH2_CLIENT_ID: str = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"  # Thunderbird public client ID
OAUTH2_TENANT: str    = "consumers"   # 'consumers' for @outlook.com / @hotmail.com
OAUTH2_SCOPES: list   = ["https://outlook.office.com/IMAP.AccessAsUser.All"]
OAUTH2_TOKEN_CACHE: str = "oauth_token_cache.json"   # persisted token cache file

# Mailbox folder to monitor (IMAP folder name)
IMAP_INBOX_FOLDER: str = "INBOX"

# IMAP IDLE timeout in seconds (reconnect cycle); keep ≤ 28 min per RFC
IMAP_IDLE_TIMEOUT: int = 1500   # 25 minutes

# ---------------------------------------------------------------------------
# Excel Configuration File
# ---------------------------------------------------------------------------
EXCEL_CONFIG_PATH: str = "/home/sudhan/workspace/mail-fetch/Automation_Config.xlsx"

SHEET_CLIENTS: str = "Clients"   # Columns: Client Name | Server
SHEET_EMAILS: str  = "Emails"    # Columns: Email Address | Type | Client Name | Domain

# ---------------------------------------------------------------------------
# Base Folder & HNI Paths  (Linux paths — use mounted network share if needed)
# ---------------------------------------------------------------------------
BASE_FOLDER: str = "/home/sudhan/workspace/mail-fetch/test_clients"

HNI_MAP: dict[str, str] = {
    "HNI1": "/home/sudhan/workspace/mail-fetch/test_clients/Tax HNI 1",
    "HNI2": "/home/sudhan/workspace/mail-fetch/test_clients/Tax HNI 2",
}

# Default HNI folder when creating a brand-new client
DEFAULT_HNI_KEY: str = "HNI1"

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
import os as _os
LOG_FILE: str = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "automation.log")
LOG_MAX_BYTES: int  = 10 * 1024 * 1024   # 10 MB
LOG_BACKUP_COUNT: int = 5

# ---------------------------------------------------------------------------
# Attachment Handling
# ---------------------------------------------------------------------------
DUPLICATE_TIMESTAMP_FORMAT: str = "%Y%m%d_%H%M%S"
SKIP_EXTENSIONS: list[str] = []   # e.g. [".exe", ".bat"] — empty = save all

# Folder for emails whose sender is not in Excel.
# Saves to: <BASE_FOLDER>/UNKNOWN_SENDER_FOLDER/<sender_email>/FY YYYY-YY/
UNKNOWN_SENDER_FOLDER: str = "new"

# ---------------------------------------------------------------------------
# Auto-install packages
# ---------------------------------------------------------------------------
REQUIRED_PACKAGES: list[str] = [
    "imapclient",
    "pandas",
    "openpyxl",
    "msal",
]
