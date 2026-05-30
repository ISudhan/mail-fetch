"""
install_startup.py
------------------
Run this script ONCE to install a systemd user service so the monitor
starts automatically on login (no root required).

Usage:
    python install_startup.py           # install & enable
    python install_startup.py --remove  # disable & remove
    python install_startup.py --status  # show current status
    python install_startup.py --logs    # tail service logs
"""

from __future__ import annotations
import os
import subprocess
import sys
from pathlib import Path

SCRIPT_PATH  = Path(__file__).resolve().parent / "outlook_sharepoint_sync.py"
PYTHON_BIN   = sys.executable
SERVICE_NAME = "outlook-attachment-monitor"
SERVICE_DIR  = Path.home() / ".config" / "systemd" / "user"
SERVICE_FILE = SERVICE_DIR / f"{SERVICE_NAME}.service"

SERVICE_CONTENT = f"""[Unit]
Description=Outlook Attachment Automation Monitor
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
ExecStart={PYTHON_BIN} {SCRIPT_PATH}
WorkingDirectory={SCRIPT_PATH.parent}
Restart=on-failure
RestartSec=15
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=default.target
"""


def _run(cmd: list[str], check: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, check=check, capture_output=True, text=True)


def install() -> None:
    if not SCRIPT_PATH.exists():
        print(f"[ERROR] Script not found: {SCRIPT_PATH}")
        sys.exit(1)

    SERVICE_DIR.mkdir(parents=True, exist_ok=True)
    SERVICE_FILE.write_text(SERVICE_CONTENT)
    print(f"[OK] Service file written: {SERVICE_FILE}")

    _run(["systemctl", "--user", "daemon-reload"])
    _run(["systemctl", "--user", "enable", SERVICE_NAME])
    _run(["systemctl", "--user", "start",  SERVICE_NAME])

    # Enable lingering so service starts even without a GUI login
    try:
        _run(["loginctl", "enable-linger", os.getenv("USER", "")])
        print("[OK] Linger enabled — service will start at boot.")
    except Exception:
        print("[WARN] Could not enable linger (run as root if needed): loginctl enable-linger $USER")

    print(f"[OK] Service '{SERVICE_NAME}' installed and started.")
    print("     Run:  systemctl --user status", SERVICE_NAME)


def remove() -> None:
    _run(["systemctl", "--user", "stop",    SERVICE_NAME], check=False)
    _run(["systemctl", "--user", "disable", SERVICE_NAME], check=False)
    if SERVICE_FILE.exists():
        SERVICE_FILE.unlink()
        print(f"[OK] Service file removed: {SERVICE_FILE}")
    _run(["systemctl", "--user", "daemon-reload"])
    print(f"[OK] Service '{SERVICE_NAME}' removed.")


def status() -> None:
    result = _run(["systemctl", "--user", "status", SERVICE_NAME], check=False)
    print(result.stdout or result.stderr)


def logs() -> None:
    os.execvp("journalctl", ["journalctl", "--user", "-u", SERVICE_NAME, "-f", "--no-pager"])


if __name__ == "__main__":
    arg = sys.argv[1].lower() if len(sys.argv) > 1 else ""
    if   arg == "--remove": remove()
    elif arg == "--status": status()
    elif arg == "--logs":   logs()
    else:                   install()
