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

---

## 📁 Project Structure

```
mail-fetch/
├── outlook_sharepoint_sync.py   # Main application
├── config.py                    # All settings (edit this first)
├── install_startup.py           # systemd service installer
├── requirements.txt             # Python dependencies
├── automation.log               # Created at runtime
└── README.md
```

---

## ⚙️ Configuration

Edit **`config.py`** before running:

```python
# IMAP Credentials
IMAP_HOST     = "outlook.office365.com"   # or "imap-mail.outlook.com"
IMAP_USERNAME = "your_email@domain.com"
IMAP_PASSWORD = "your_app_password"       # Use App Password if MFA is on

# Excel config file path
EXCEL_CONFIG_PATH = "/home/sudhan/Automation_Config.xlsx"

# Folder structure
HNI_MAP = {
    "HNI1": "/mnt/clients/Tax HNI 1",
    "HNI2": "/mnt/clients/Tax HNI 2",
}
```

---

## 📊 Excel Configuration (`Automation_Config.xlsx`)

### Sheet 1 — `Clients`

| Client Name    | Server |
|----------------|--------|
| John Doe       | HNI1   |
| ABC Industries | HNI2   |

### Sheet 2 — `Emails`

| Email Address           | Type   | Client Name  | Domain        |
|-------------------------|--------|--------------|---------------|
| broker@zerodha.com      | Broker | Zerodha      | zerodha.com   |
| john@gmail.com          | Client | John Doe     |               |
| support@icicidirect.com | Broker | ICICI Direct | icicidirect.com |

> **Rules:** Leave *Domain* blank for client rows. Fill *Domain* for broker rows to match all emails from that broker automatically.

---

## 🚀 Quick Start

### 1. Clone / download the project

```bash
cd /home/sudhan/workspace/mail-fetch
```

### 2. Create virtual environment & install dependencies

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 3. Enable IMAP on Outlook

- **Microsoft 365:** outlook.office.com → Settings → Mail → Sync email → Enable IMAP
- **Outlook.com:** Settings → View all Outlook settings → Mail → Sync email → POP and IMAP
- **With MFA:** Create an App Password at [account.microsoft.com](https://account.microsoft.com) → Security → App passwords

### 4. Edit `config.py` with your credentials

### 5. Run

```bash
source .venv/bin/activate
python outlook_sharepoint_sync.py
```

---

## 🔁 Auto-Start with systemd

Register the monitor to start automatically at login:

```bash
python install_startup.py           # install
python install_startup.py --status  # check status
python install_startup.py --logs    # live logs
python install_startup.py --remove  # uninstall
```

---

## 📂 Folder Structure Created

```
/mnt/clients/
└── Tax HNI 1/
│   └── John Doe/
│       └── FY 2026-27/
│           ├── Zerodha/
│           │   └── statement.pdf
│           └── report.pdf        ← no broker identified
└── Tax HNI 2/
    └── ABC Industries/
        └── FY 2026-27/
            └── circular.pdf
```

---

## 🧠 Business Logic

| Scenario | Condition | Result |
|----------|-----------|--------|
| 1 | Broker in FROM, Client in CC/TO/BCC | Save to client's folder under broker subfolder |
| 2 | Client in FROM, Broker in CC/TO/BCC | Save to client's folder under broker subfolder |
| 3 | Only client identified | Save to client's FY folder (no broker subfolder) |
| 4 | Only broker identified | Resolve linked client from Excel, save there |
| 5 | Multiple clients identified | Copy attachment to **all** matched client folders |
| 6 | No broker identified | Save directly under FY folder |

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

Example:
```
2026-05-30 14:30:01 INFO     Email received: Trade confirmation
2026-05-30 14:30:01 INFO     From: ['broker@zerodha.com']
2026-05-30 14:30:01 INFO     Matched clients: ['John Doe']
2026-05-30 14:30:01 INFO     Matched brokers: ['Zerodha']
2026-05-30 14:30:01 INFO     Attachment saved: /mnt/clients/Tax HNI 1/John Doe/FY 2026-27/Zerodha/statement.pdf
```

---

## 🛠️ Troubleshooting

| Symptom | Fix |
|---------|-----|
| `AUTHENTICATE failed` | Use App Password if MFA is enabled |
| `[AUTHENTICATIONFAILED]` | Enable IMAP in Outlook account settings |
| `ConnectionRefusedError` | Verify `IMAP_HOST` and `IMAP_PORT` |
| Excel not found | Check `EXCEL_CONFIG_PATH` in config.py |
| Folders not created | Check write permissions on the mount point |
| systemd service failing | `journalctl --user -u outlook-attachment-monitor -xe` |

---

## 📦 Dependencies

| Package | Purpose |
|---------|---------|
| `imapclient` | IMAP IDLE real-time monitoring |
| `pandas` | Excel config parsing |
| `openpyxl` | Excel file reading backend |

---

## 📄 License

MIT — free to use and modify.



# Check if it's running
systemctl --user status outlook-attachment-monitor

# View live logs
journalctl --user -u outlook-attachment-monitor -f

# Stop it
systemctl --user stop outlook-attachment-monitor

# Start it manually
systemctl --user start outlook-attachment-monitor

# Remove auto-start completely
python install_startup.py --remove
