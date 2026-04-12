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

# ── Web server ────────────────────────────────────────────────────────────────
WEB_PORT = int(os.environ.get("WEB_PORT", "5000"))

# ── Database ──────────────────────────────────────────────────────────────────
# Local:  instance/contacts.db  (default)
# Azure:  /home/data/contacts.db  (set DB_PATH in App Settings)
DB_PATH = os.environ.get("DB_PATH", "instance/contacts.db")

# ── Application constants ─────────────────────────────────────────────────────
# Allowed email domain for Easy Auth / Microsoft 365 login
DOMAIN      = os.environ.get("APP_DOMAIN", "invisionvail.com")
# Email address that gets the 'admin' role on first login
ADMIN_EMAIL = os.environ.get("ADMIN_EMAIL", "johnh@invisionvail.com")
# IANA timezone name used for display in the admin UI
TIMEZONE    = os.environ.get("APP_TIMEZONE", "America/Denver")
# Contacts per page on the main list
PAGE_SIZE   = int(os.environ.get("PAGE_SIZE", "50"))
# Short abbreviations shown in category bubbles (key = full name, value = 3-char code)
CATEGORY_ABBREVS = {
    "Architects":                  "ARC",
    "Builder":                     "BLD",
    "Current Customer":            "CC",
    "Electricians":                "ELC",
    "General":                     "GEN",
    "Interior Designer":           "IND",
    "Mfg Reps and Distributors":   "MFG",
    "Past Customer":               "PC",
    "Personal":                    "PRS",
    "Property Manager":            "PMG",
    "Real Estate Agent":           "REA",
    "Sub Contractor":              "SUB",
}

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
