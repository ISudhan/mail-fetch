# 📬 Outlook Attachment Automation System

A production-ready Python automation system that monitors a Microsoft Outlook / Microsoft 365 mailbox via **IMAP IDLE** (real-time push, no polling) and automatically saves email attachments into a structured client folder hierarchy — organised by client name, Indian Financial Year, and broker.

---

## ✨ Features

- 🔔 **Real-time monitoring** via IMAP IDLE — no polling loops
- 📁 **Auto folder creation** — HNI → Client → FY → Broker
- 📅 **Auto Financial Year** calculation (Indian FY: Apr → Mar)
- 👥 **Multi-client routing** — one email, multiple client folders
- 🔍 **Smart email matching** — exact address first, then domain fallback
- 📎 **Duplicate filename handling** — timestamp suffix appended
- 🔄 **Auto-reconnect** on network/IMAP failures
- 📝 **Rotating logs** — structured timestamped entries
- ⚙️ **systemd auto-start** — starts at login, no root required
- 📦 **Auto dependency install** on first run
- 📂 **Unknown sender fallback** — saves to `new/<email>/FY/` if not in Excel

---

## 📁 Project Structure

```
mail-fetch/
├── outlook_sharepoint_sync.py   # Main application
├── config.py                    # All settings (edit this first)
├── install_startup.py           # systemd service installer
├── requirements.txt             # Python dependencies
├── Automation_Config.xlsx       # Client/email mapping (you create this)
├── automation.log               # Created at runtime
├── oauth_token_cache.json       # OAuth2 token cache (auto-created, never commit)
└── README.md
```

> ⚠️ `oauth_token_cache.json` and `config.py` contain sensitive credentials. Both are listed in `.gitignore` and must **never** be committed to version control.

---

## ⚙️ Configuration

Edit **`config.py`** before running:

```python
# IMAP host for personal Outlook.com
IMAP_HOST     = "imap-mail.outlook.com"
IMAP_USERNAME = "your_email@outlook.com"

# OAuth2 (no password needed — browser login once)
OAUTH2_CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"
OAUTH2_TENANT    = "consumers"

# Excel config file path
EXCEL_CONFIG_PATH = "/home/youruser/Automation_Config.xlsx"

# Folder structure
HNI_MAP = {
    "HNI1": "/path/to/Tax HNI 1",
    "HNI2": "/path/to/Tax HNI 2",
}

# Unknown senders saved here: <BASE_FOLDER>/new/<sender@email>/FY YYYY-YY/
UNKNOWN_SENDER_FOLDER = "new"
```

---

## 📊 Excel Configuration (`Automation_Config.xlsx`)

### Sheet 1 — `Clients`

| Client Name    | Server |
|----------------|--------|
| John Doe       | HNI1   |
| ABC Industries | HNI2   |

### Sheet 2 — `Emails`

| Email Address           | Type   | Client Name  | Domain          |
|-------------------------|--------|--------------|-----------------|
| broker@zerodha.com      | Broker | Zerodha      | zerodha.com     |
| john@gmail.com          | Client | John Doe     |                 |
| support@icicidirect.com | Broker | ICICI Direct | icicidirect.com |

> **Rules:**
> - Leave *Domain* blank for client rows.
> - Fill *Domain* for broker rows to match all emails from that broker automatically.
> - Unknown senders (not in Excel) are automatically saved to `new/<sender_email>/FY/`.

---

## 🚀 Quick Start

### 1. Install dependencies

```bash
cd /path/to/mail-fetch
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Enable IMAP on your Outlook account

- Go to [outlook.office.com](https://outlook.office.com) → Settings → Mail → Sync email → Enable IMAP
- For personal Outlook.com: Settings → View all Outlook settings → Mail → Sync email → POP and IMAP

### 3. Edit `config.py` with your email and paths

### 4. Run the monitor

```bash
source .venv/bin/activate
python outlook_sharepoint_sync.py
```

**First run only:** A Microsoft login prompt will appear in the terminal:
```
============================================================
 Microsoft Login Required
============================================================
 1. Open: https://www.microsoft.com/link
 2. Enter code: XXXXXXXX
============================================================
```
Open the URL, enter the code, and sign in. The token is cached — **you won't need to do this again**.

---

## 🔁 Auto-Start with systemd (run once)

Register the monitor to start automatically at every login:

```bash
python install_startup.py
```

### Manage the service

```bash
# Check if it's running
systemctl --user status outlook-attachment-monitor

# View live logs
journalctl --user -u outlook-attachment-monitor -f

# Stop the monitor
systemctl --user stop outlook-attachment-monitor

# Start manually
systemctl --user start outlook-attachment-monitor

# Remove auto-start completely
python install_startup.py --remove
```

---

## 📂 Folder Structure Created

```
/path/to/clients/
├── Tax HNI 1/
│   └── John Doe/
│       └── FY 2026-27/
│           ├── Zerodha/
│           │   └── statement.pdf      ← broker identified
│           └── report.pdf             ← no broker
└── new/
    └── unknown@gmail.com/
        └── FY 2026-27/
            └── attachment.jpg         ← sender not in Excel
```

---

## 🧠 Business Logic

| Scenario | Condition | Result |
|----------|-----------|--------|
| 1 | Broker in FROM, Client in CC/TO | Save to client folder under broker subfolder |
| 2 | Client in FROM, Broker in CC | Save to client folder under broker subfolder |
| 3 | Only client identified | Save directly under FY folder |
| 4 | Only broker identified | Resolve linked client from Excel |
| 5 | Multiple clients identified | Copy attachment to **all** matched client folders |
| 6 | No broker identified | Save directly under FY (no broker subfolder) |
| 7 | **Sender not in Excel** | Save to `new/<sender_email>/FY/` automatically |

### Client Folder Resolution

1. Search **Tax HNI 1** for existing folder
2. Search **Tax HNI 2** for existing folder
3. If not found → create under **Tax HNI 1** (default)
4. Auto-create missing FY and Broker subfolders

### Financial Year Logic

```
April 2026 – March 2027  →  FY 2026-27
April 2027 – March 2028  →  FY 2027-28
```

Calculated automatically. No config change needed at year-end.

---

## 📋 Logs

```bash
tail -f automation.log
```

Example output:
```
2026-05-30 13:25:21 INFO     Email received: Trade Confirmation
2026-05-30 13:25:21 INFO     From: ['broker@zerodha.com']
2026-05-30 13:25:21 INFO     Matched clients: ['John Doe']
2026-05-30 13:25:21 INFO     Matched brokers: ['Zerodha']
2026-05-30 13:25:21 INFO     Attachment saved: .../John Doe/FY 2026-27/Zerodha/statement.pdf
```

---

## 🛠️ Troubleshooting

| Symptom | Fix |
|---------|-----|
| `AUTHENTICATE failed` | Re-run script; complete the browser device-code login |
| IMAP login loop | Delete `oauth_token_cache.json` and re-authenticate |
| Excel not found | Check `EXCEL_CONFIG_PATH` in config.py |
| Folders not created | Check write permissions on the destination path |
| systemd service failing | `journalctl --user -u outlook-attachment-monitor -xe` |
| Token expired | Delete `oauth_token_cache.json` — will re-authenticate on next run |

---

## 📦 Dependencies

| Package | Purpose |
|---------|---------|
| `imapclient` | IMAP IDLE real-time monitoring |
| `msal` | Microsoft OAuth2 authentication |
| `pandas` | Excel config parsing |
| `openpyxl` | Excel file reading backend |

---

## 🔒 Security Notes

- **Never commit** `oauth_token_cache.json` or `config.py` — both are in `.gitignore`
- OAuth2 tokens are stored locally and refresh automatically
- No passwords are stored anywhere in the codebase

---

## 📄 License

MIT — free to use and modify.
