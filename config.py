# ============================================================
#  ContactSync — configuration
#  All secrets are read from environment variables.
#  For local development, put values in a .env file.
#  For Azure App Services, set them as App Settings.
# ============================================================
import os
from dotenv import load_dotenv

# Load .env file when running locally.
# On Azure App Services, App Settings are already in the environment
# and load_dotenv() is a harmless no-op (no .env file present).
load_dotenv()

def _require(name):
    """Read a required environment variable; raise clearly if missing."""
    val = os.environ.get(name)
    if not val:
        raise EnvironmentError(
            f"Required environment variable '{name}' is not set. "
            f"Add it to your .env file (local) or App Settings (Azure)."
        )
    return val

# ── Microsoft 365 / Azure App Registration ────────────────────────────────────
TENANT_ID     = _require("TENANT_ID")
CLIENT_ID     = _require("CLIENT_ID")
CLIENT_SECRET = _require("CLIENT_SECRET")

# ── Mailbox to sync ───────────────────────────────────────────────────────────
MAILBOX_EMAIL = _require("MAILBOX_EMAIL")

# ── Flask session security ────────────────────────────────────────────────────
SECRET_KEY = _require("SECRET_KEY")

# ── Sync settings ─────────────────────────────────────────────────────────────
# Default: 30 minutes (1800 seconds)
SYNC_INTERVAL_SECONDS = int(os.environ.get("SYNC_INTERVAL_SECONDS", "1800"))

# ── Timezone ──────────────────────────────────────────────────────────────────
# All logs and timestamps use Mountain Time (US/Mountain)
TIMEZONE = "US/Mountain"

# ── Web server ────────────────────────────────────────────────────────────────
WEB_PORT = int(os.environ.get("WEB_PORT", "5000"))

# ── Database ──────────────────────────────────────────────────────────────────
# Local:  instance/contacts.db  (default)
# Azure:  /home/data/contacts.db  (set DB_PATH in App Settings)
DB_PATH = os.environ.get("DB_PATH", "instance/contacts.db")

# ── User accounts ─────────────────────────────────────────────────────────────
# Stored as: username:password:role,username:password:role,...
# Example:   viewer:secret1:readonly,editor:secret2:editor,admin:secret3:admin
def _parse_users(raw):
    users = {}
    for entry in raw.split(","):
        entry = entry.strip()
        if not entry:
            continue
        parts = entry.split(":", 2)
        if len(parts) == 3:
            uname, pwd, role = parts
            users[uname.strip()] = {"password": pwd.strip(), "role": role.strip()}
    return users

_users_raw = os.environ.get("USERS", "")
USERS = _parse_users(_users_raw) if _users_raw else {}
