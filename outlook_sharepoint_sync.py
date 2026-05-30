"""
outlook_sharepoint_sync.py
--------------------------
Outlook Attachment Automation System — Linux / IMAP Edition.

Monitors a Microsoft Outlook (Office 365 / Exchange) mailbox via IMAP IDLE
(real-time push — no polling loop) and saves email attachments into a
structured client folder hierarchy.

Requirements:
    - Linux (any distro with Python 3.10+)
    - pip install -r requirements.txt
    - IMAP must be enabled on the Outlook/Exchange account
    - Fill in credentials in config.py before running

Usage:
    python outlook_sharepoint_sync.py
"""

from __future__ import annotations

# ── 0. Auto-install missing dependencies ─────────────────────────────────────
import importlib
import subprocess
import sys


def _auto_install(packages: list[str]) -> None:
    name_map = {"imapclient": "imapclient", "pandas": "pandas", "openpyxl": "openpyxl"}
    for pkg in packages:
        mod = name_map.get(pkg, pkg)
        try:
            importlib.import_module(mod)
        except ImportError:
            print(f"[SETUP] Installing {pkg} …")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])
            print(f"[SETUP] {pkg} installed.")


import config  # noqa: E402
_auto_install(config.REQUIRED_PACKAGES)

# ── 1. Imports ────────────────────────────────────────────────────────────────
import email
import email.policy
import logging
import logging.handlers
import os
import re
import socket
import time
from datetime import date, datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Optional

import imapclient
import msal
import pandas as pd

# ── 2. Logging ────────────────────────────────────────────────────────────────

def _setup_logging() -> logging.Logger:
    logger = logging.getLogger("OutlookSync")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter("%(asctime)s %(levelname)-8s %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")

    fh = logging.handlers.RotatingFileHandler(
        config.LOG_FILE,
        maxBytes=config.LOG_MAX_BYTES,
        backupCount=config.LOG_BACKUP_COUNT,
        encoding="utf-8",
    )
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


log = _setup_logging()

# ── 3. Financial Year ─────────────────────────────────────────────────────────

def get_financial_year(d: date) -> str:
    """
    Indian Financial Year string for *d*.
    April 1 → March 31.  E.g.  2026-07-01 → 'FY 2026-27'
    """
    fy_start = d.year if d.month >= 4 else d.year - 1
    fy_end   = (fy_start + 1) % 100
    return f"FY {fy_start}-{fy_end:02d}"


# ── 4. Excel Config ───────────────────────────────────────────────────────────

class ExcelConfig:
    """Loads client/email mappings from the Excel workbook."""

    def __init__(self, path: str) -> None:
        self.path = path
        self._clients: pd.DataFrame = pd.DataFrame()
        self._emails:  pd.DataFrame = pd.DataFrame()
        self.reload()

    def reload(self) -> None:
        if not os.path.exists(self.path):
            log.error("Excel config not found: %s", self.path)
            return
        try:
            self._clients = pd.read_excel(
                self.path, sheet_name=config.SHEET_CLIENTS, dtype=str).fillna("")
            self._emails  = pd.read_excel(
                self.path, sheet_name=config.SHEET_EMAILS,  dtype=str).fillna("")
            self._clients.columns = [c.strip() for c in self._clients.columns]
            self._emails.columns  = [c.strip() for c in self._emails.columns]
            log.info("Excel loaded — %d clients, %d email rules",
                     len(self._clients), len(self._emails))
        except Exception as exc:
            log.error("Failed to load Excel config: %s", exc)

    def lookup_email(self, address: str) -> Optional[dict]:
        """Exact match first, then domain match (case-insensitive)."""
        if self._emails.empty:
            return None
        addr = address.strip().lower()

        # Exact
        mask = self._emails["Email Address"].str.strip().str.lower() == addr
        rows = self._emails[mask]
        if not rows.empty:
            return rows.iloc[0].to_dict()

        # Domain
        if "@" in addr:
            domain = addr.split("@", 1)[1]
            mask_d = self._emails["Domain"].str.strip().str.lower() == domain
            rows = self._emails[mask_d]
            if not rows.empty:
                return rows.iloc[0].to_dict()
        return None

    def lookup_client(self, client_name: str) -> Optional[dict]:
        """Return the Clients sheet row for *client_name* (case-insensitive)."""
        if self._clients.empty:
            return None
        name = client_name.strip().lower()
        mask = self._clients["Client Name"].str.strip().str.lower() == name
        rows = self._clients[mask]
        return rows.iloc[0].to_dict() if not rows.empty else None

    @property
    def emails_df(self) -> pd.DataFrame:
        return self._emails


# ── 5. Folder Resolver ────────────────────────────────────────────────────────

class FolderResolver:
    """Resolves and creates the correct destination folder for attachments."""

    def __init__(self, excel: ExcelConfig) -> None:
        self.excel = excel

    def _client_folder(self, client_name: str) -> str:
        """Resolve the client folder using the 'Server' column as source of truth."""
        # ── 1. Look up the Server key from the Clients sheet ─────────────────
        client_row = self.excel.lookup_client(client_name)
        server_key = ""
        if client_row:
            server_key = client_row.get("Server", "").strip()

        if not server_key:
            raise ValueError(
                f"Client '{client_name}' has no 'Server' value in the Clients sheet. "
                f"Please set it to one of: {list(config.HNI_MAP.keys())}"
            )
        if server_key not in config.HNI_MAP:
            raise ValueError(
                f"Client '{client_name}' has Server='{server_key}' which is not in "
                f"HNI_MAP. Valid keys: {list(config.HNI_MAP.keys())}"
            )

        # ── 2. Use that HNI path — find or create the folder there ───────────
        base_path = config.HNI_MAP[server_key]
        client_path = os.path.join(base_path, client_name)
        if os.path.isdir(client_path):
            log.debug("Found existing client folder: %s", client_path)
        else:
            log.info("Creating new client folder: %s", client_path)
            Path(client_path).mkdir(parents=True, exist_ok=True)
        return client_path

    def resolve(self, client_name: str, broker_name: Optional[str],
                received_date: date) -> str:
        """Return fully resolved and created destination path."""
        fy          = get_financial_year(received_date)
        client_dir  = self._client_folder(client_name)
        dest        = os.path.join(client_dir, fy, broker_name) if broker_name \
                      else os.path.join(client_dir, fy)
        Path(dest).mkdir(parents=True, exist_ok=True)
        log.debug("Resolved destination: %s", dest)
        return dest


# ── 6. Attachment Saver ───────────────────────────────────────────────────────

class AttachmentSaver:
    """Saves email attachments to disk with duplicate-name handling."""

    @staticmethod
    def _safe_filename(name: str) -> str:
        return re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip()

    @staticmethod
    def _unique_path(folder: str, filename: str) -> str:
        target = os.path.join(folder, filename)
        if not os.path.exists(target):
            return target
        stem, ext = os.path.splitext(filename)
        ts = datetime.now().strftime(config.DUPLICATE_TIMESTAMP_FORMAT)
        new_name   = f"{stem}_{ts}{ext}"
        new_target = os.path.join(folder, new_name)
        log.warning("Duplicate — renaming to: %s", new_name)
        return new_target

    def save(self, filename: str, payload: bytes, dest_folder: str) -> Optional[str]:
        try:
            if config.SKIP_EXTENSIONS:
                ext = os.path.splitext(filename)[1].lower()
                if ext in [e.lower() for e in config.SKIP_EXTENSIONS]:
                    log.info("Skipping (filtered extension): %s", filename)
                    return None
            safe = self._safe_filename(filename)
            path = self._unique_path(dest_folder, safe)
            with open(path, "wb") as f:
                f.write(payload)
            log.info("Attachment saved: %s", path)
            return path
        except Exception as exc:
            log.error("Failed to save attachment '%s': %s", filename, exc)
            return None


# ── 7. Email Processor ────────────────────────────────────────────────────────

class EmailProcessor:
    """Parses a raw email, matches clients/brokers, saves attachments."""

    def __init__(self, excel: ExcelConfig) -> None:
        self.excel    = excel
        self.resolver = FolderResolver(excel)
        self.saver    = AttachmentSaver()

    # ------------------------------------------------------------------
    @staticmethod
    def _extract_addrs(header_val: str) -> list[str]:
        """Parse all email addresses from a header value string."""
        if not header_val:
            return []
        addrs = []
        for part in re.split(r"[;,]", header_val):
            part = part.strip()
            m = re.search(r"<([^>]+)>", part)
            addr = m.group(1) if m else part
            if "@" in addr:
                addrs.append(addr.strip().lower())
        return addrs

    def _collect_all_addresses(self, msg: EmailMessage) -> dict[str, list[str]]:
        return {
            "From": self._extract_addrs(msg.get("From", "")),
            "To":   self._extract_addrs(msg.get("To",   "")),
            "CC":   self._extract_addrs(msg.get("CC",   "")),
            "BCC":  self._extract_addrs(msg.get("BCC",  "")),
        }

    def _classify(self, fields: dict[str, list[str]]) -> tuple[list[dict], list[dict]]:
        clients, brokers = [], []
        seen: set[str] = set()
        for addrs in fields.values():
            for addr in addrs:
                if addr in seen:
                    continue
                seen.add(addr)
                result = self.excel.lookup_email(addr)
                if not result:
                    continue
                rtype = result.get("Type", "").strip().lower()
                if rtype == "client":
                    clients.append(result)
                elif rtype == "broker":
                    brokers.append(result)
        return clients, brokers

    # ------------------------------------------------------------------
    def process(self, raw_bytes: bytes) -> None:
        try:
            # Reload Excel config before each email so any new/edited rows
            # (clients, email addresses) take effect without a restart.
            self.excel.reload()

            msg: EmailMessage = email.message_from_bytes(
                raw_bytes, policy=email.policy.default)

            subject = msg.get("Subject", "(no subject)")
            log.info("─" * 60)
            log.info("Email received: %s", subject)

            # Collect attachments first — skip if none
            # Grab every leaf part that has a filename (covers PDFs, Word, Excel,
            # images, etc. regardless of Content-Disposition value).
            # Exception: skip image/* parts that are NOT explicitly marked as
            # "attachment" — those are typically inline logos/icons embedded in
            # HTML email bodies.
            attachments: list[tuple[str, bytes]] = []
            for part in msg.walk():
                # Skip multipart containers — they hold no payload themselves
                if part.get_content_maintype() == "multipart":
                    continue
                fname = part.get_filename()
                if not fname:
                    continue
                disp = (part.get_content_disposition() or "").lower()
                maintype = part.get_content_maintype().lower()
                # Skip inline images (embedded logos / email decorations)
                if maintype == "image" and "attachment" not in disp:
                    log.debug("Skipping inline image: %s", fname)
                    continue
                payload = part.get_payload(decode=True)
                if payload:
                    attachments.append((fname, payload))
                    log.debug("Queued attachment: %s (%s/%s, disp=%s)",
                              fname, maintype, part.get_content_subtype(), disp or "none")

            if not attachments:
                log.info("No attachments — skipping.")
                return

            fields = self._collect_all_addresses(msg)
            log.info("From: %s", fields["From"])
            log.info("To  : %s", fields["To"])
            log.info("CC  : %s", fields["CC"])
            log.info("BCC : %s", fields["BCC"])

            clients, brokers = self._classify(fields)
            client_names = [r.get("Client Name", "").strip() for r in clients]
            broker_names = [r.get("Client Name", "").strip() for r in brokers]
            log.info("Matched clients: %s", client_names or "none")
            log.info("Matched brokers: %s", broker_names or "none")

            # Received date for FY
            try:
                date_str = msg.get("Date", "")
                received_dt: date = email.utils.parsedate_to_datetime(date_str).date()
            except Exception:
                received_dt = date.today()

            broker_display = broker_names[0] if broker_names else None

            # Build target list: (client_name, broker_name)
            targets: list[tuple[str, Optional[str]]] = []
            if client_names:
                for cname in client_names:
                    targets.append((cname, broker_display))
            elif broker_names:
                raw_linked = [r.get("Client Name", "").strip() for r in brokers if r.get("Client Name")]
                # Only use names that actually exist as clients in the Clients sheet
                linked = [n for n in raw_linked if self.excel.lookup_client(n)]
                if linked:
                    for cname in linked:
                        targets.append((cname, broker_display))
                else:
                    log.warning(
                        "Only broker '%s' found but no valid linked client — "
                        "routing to unknown-sender folder.", broker_names
                    )
            else:
                # ── Fallback: unknown sender — save to new/<sender_email>/FY/ ──
                sender_addrs = fields.get("From", [])
                sender = sender_addrs[0] if sender_addrs else "unknown"
                # Sanitise email for use as folder name (keep @ and . — valid on Linux)
                safe_sender = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", sender).strip()
                dest = os.path.join(
                    config.BASE_FOLDER,
                    config.UNKNOWN_SENDER_FOLDER,
                    safe_sender,
                    get_financial_year(received_dt),
                )
                Path(dest).mkdir(parents=True, exist_ok=True)
                log.warning(
                    "Unknown sender '%s' — saving to fallback folder: %s", sender, dest
                )
                for fname, payload in attachments:
                    self.saver.save(fname, payload, dest)
                return

            # Save attachments to every matched target
            for client_name, broker_name in targets:
                if not client_name:
                    continue
                dest = self.resolver.resolve(client_name, broker_name, received_dt)
                for fname, payload in attachments:
                    self.saver.save(fname, payload, dest)

        except Exception as exc:
            log.error("Unhandled error processing email: %s", exc, exc_info=True)


# ── 8. IMAP IDLE Monitor ──────────────────────────────────────────────────────

class IMAPMonitor:
    """
    Connects to the IMAP server and uses IDLE to receive real-time
    notifications when new messages arrive.  Reconnects automatically
    after timeouts or network errors.
    """

    def __init__(self, processor: EmailProcessor) -> None:
        self.processor = processor
        self._client: Optional[imapclient.IMAPClient] = None

    # ------------------------------------------------------------------
    def _get_access_token(self) -> str:
        """
        Acquire an OAuth2 access token via MSAL.
        - First run: opens device-code flow (prints a URL + code to the terminal).
        - Subsequent runs: silently refreshes from the cached token file.
        """
        cache = msal.SerializableTokenCache()
        cache_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            config.OAUTH2_TOKEN_CACHE,
        )
        if os.path.exists(cache_path):
            cache.deserialize(open(cache_path).read())

        app = msal.PublicClientApplication(
            client_id=config.OAUTH2_CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{config.OAUTH2_TENANT}",
            token_cache=cache,
        )

        accounts = app.get_accounts(username=config.IMAP_USERNAME)
        result = None

        # Try silent refresh first
        if accounts:
            result = app.acquire_token_silent(config.OAUTH2_SCOPES, account=accounts[0])

        # Fall back to device-code flow (requires user interaction once)
        if not result:
            log.info("No cached token — starting device-code login …")
            flow = app.initiate_device_flow(scopes=config.OAUTH2_SCOPES)
            if "user_code" not in flow:
                raise RuntimeError(f"Device flow failed: {flow.get('error_description')}") 
            print("\n" + "="*60)
            print(" Microsoft Login Required")
            print("="*60)
            print(f" 1. Open: {flow['verification_uri']}")
            print(f" 2. Enter code: {flow['user_code']}")
            print("="*60 + "\n")
            result = app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            raise RuntimeError(
                f"OAuth2 token acquisition failed: {result.get('error_description', result)}"
            )

        # Persist token cache
        if cache.has_state_changed:
            with open(cache_path, "w") as f:
                f.write(cache.serialize())
            log.debug("Token cache saved: %s", cache_path)

        return result["access_token"]

    def _connect(self) -> imapclient.IMAPClient:
        log.info("Connecting to %s:%s …", config.IMAP_HOST, config.IMAP_PORT)
        token = self._get_access_token()
        # Build XOAUTH2 string: base64("user=<email>\x01auth=Bearer <token>\x01\x01")
        import base64
        xoauth2 = base64.b64encode(
            f"user={config.IMAP_USERNAME}\x01auth=Bearer {token}\x01\x01".encode()
        ).decode()
        client = imapclient.IMAPClient(
            config.IMAP_HOST,
            port=config.IMAP_PORT,
            ssl=config.IMAP_USE_SSL,
        )
        client.oauth2_login(config.IMAP_USERNAME, token)
        client.select_folder(config.IMAP_INBOX_FOLDER)
        log.info("IMAP connected via OAuth2 and '%s' selected.", config.IMAP_INBOX_FOLDER)
        return client

    def _fetch_and_process(self, uid: int) -> None:
        try:
            response = self._client.fetch([uid], ["RFC822"])
            if uid in response:
                raw = response[uid][b"RFC822"]
                self.processor.process(raw)
        except Exception as exc:
            log.error("Error fetching message UID %s: %s", uid, exc)

    # ------------------------------------------------------------------
    def run(self) -> None:
        """Main loop — connects, enters IDLE, processes new mail forever."""
        while True:
            try:
                self._client = self._connect()

                # Process any unseen messages that arrived while offline
                unseen = self._client.search(["UNSEEN"])
                if unseen:
                    log.info("Processing %d unseen message(s) from offline period.", len(unseen))
                    for uid in unseen:
                        self._fetch_and_process(uid)

                log.info("Entering IMAP IDLE mode (timeout %ss) …", config.IMAP_IDLE_TIMEOUT)
                self._client.idle()

                while True:
                    # idle_check blocks up to timeout seconds
                    responses = self._client.idle_check(timeout=config.IMAP_IDLE_TIMEOUT)

                    if not responses:
                        # Timeout — send DONE + re-issue IDLE to keep connection alive
                        log.debug("IDLE timeout — refreshing …")
                        self._client.idle_done()
                        self._client.idle()
                        continue

                    # Check for EXISTS responses (new mail notification)
                    has_new = any(
                        isinstance(r, tuple) and len(r) >= 2 and r[1] == b"EXISTS"
                        for r in responses
                    )
                    if has_new:
                        self._client.idle_done()
                        new_uids = self._client.search(["UNSEEN"])
                        log.debug("New UIDs: %s", new_uids)
                        for uid in new_uids:
                            self._fetch_and_process(uid)
                        # Re-enter IDLE
                        self._client.idle()

            except (imapclient.IMAPClient.AbortError,
                    imapclient.IMAPClient.ReadOnlyError,
                    ConnectionResetError,
                    OSError,
                    socket.error) as exc:
                log.warning("IMAP connection lost (%s) — reconnecting in 30s …", exc)
                self._safe_logout()
                time.sleep(30)

            except KeyboardInterrupt:
                log.info("Shutdown requested.")
                self._safe_logout()
                break

            except Exception as exc:
                log.error("Unexpected error: %s — reconnecting in 60s …", exc, exc_info=True)
                self._safe_logout()
                time.sleep(60)

    def _safe_logout(self) -> None:
        try:
            if self._client:
                self._client.logout()
        except Exception:
            pass
        self._client = None


# ── 9. Bootstrap ──────────────────────────────────────────────────────────────

def main() -> None:
    log.info("=" * 60)
    log.info("Outlook Attachment Automation System (Linux/IMAP) — Starting")
    log.info("Log file : %s", config.LOG_FILE)
    log.info("Excel    : %s", config.EXCEL_CONFIG_PATH)
    log.info("IMAP     : %s@%s (OAuth2)", config.IMAP_USERNAME, config.IMAP_HOST)
    log.info("=" * 60)

    excel     = ExcelConfig(config.EXCEL_CONFIG_PATH)
    processor = EmailProcessor(excel)
    monitor   = IMAPMonitor(processor)
    monitor.run()

    log.info("Outlook Attachment Automation System — Stopped.")


if __name__ == "__main__":
    main()
