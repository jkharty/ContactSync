"""
database.py — creates and manages the SQLite database.
"""
import os
import sqlite3
import config

# Ensure the directory for the database file exists (required on Azure App Services
# where /home/data/ is not created automatically).
_db_dir = os.path.dirname(config.DB_PATH)
if _db_dir:
    os.makedirs(_db_dir, exist_ok=True)

def get_db():
    # timeout=30: wait up to 30 s for a write lock instead of failing immediately.
    conn = sqlite3.connect(config.DB_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    # WAL mode allows concurrent reads during writes — essential when the sync
    # scheduler thread and web request handlers access the DB simultaneously.
    conn.execute("PRAGMA journal_mode=WAL")
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS contacts (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            exchange_id     TEXT    UNIQUE NOT NULL,
            change_key      TEXT,
            first_name      TEXT,
            last_name       TEXT,
            display_name    TEXT,
            company         TEXT,
            job_title       TEXT,
            email1          TEXT,
            email2          TEXT,
            email3          TEXT,
            phone_business  TEXT,
            phone_mobile    TEXT,
            phone_home      TEXT,
            address_street  TEXT,
            address_city    TEXT,
            address_state   TEXT,
            address_zip     TEXT,
            address_country TEXT,
            home_street     TEXT,
            home_city       TEXT,
            home_state      TEXT,
            home_zip        TEXT,
            home_country    TEXT,
            other_street    TEXT,
            other_city      TEXT,
            other_state     TEXT,
            other_zip       TEXT,
            other_country   TEXT,
            categories      TEXT,
            notes_rtf       BLOB,
            notes_html      TEXT,
            notes_plain     TEXT,
            last_modified   TEXT,
            synced_at       TEXT
        );

        CREATE TABLE IF NOT EXISTS sync_log (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            started_at  TEXT,
            finished_at TEXT,
            added       INTEGER DEFAULT 0,
            updated     INTEGER DEFAULT 0,
            deleted     INTEGER DEFAULT 0,
            errors      INTEGER DEFAULT 0,
            message     TEXT
        );

        CREATE TABLE IF NOT EXISTS pending_writes (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            exchange_id     TEXT NOT NULL,
            notes_html      TEXT,
            notes_rtf       BLOB,
            field_type      TEXT DEFAULT 'notes',
            field_data      TEXT,
            requested_by    TEXT,
            requested_at    TEXT,
            status          TEXT DEFAULT 'pending'
        );

        CREATE INDEX IF NOT EXISTS idx_contacts_display_name
            ON contacts(display_name COLLATE NOCASE);
        CREATE INDEX IF NOT EXISTS idx_contacts_company
            ON contacts(company COLLATE NOCASE);
        CREATE INDEX IF NOT EXISTS idx_contacts_email1
            ON contacts(email1 COLLATE NOCASE);

        CREATE TABLE IF NOT EXISTS sync_status (
            id           INTEGER PRIMARY KEY CHECK (id = 1),
            state        TEXT DEFAULT 'idle',
            started_at   TEXT,
            finished_at  TEXT,
            last_message TEXT
        );

        INSERT OR IGNORE INTO sync_status (id, state) VALUES (1, 'idle');

        CREATE TABLE IF NOT EXISTS conflicts (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            exchange_id     TEXT NOT NULL,
            display_name    TEXT,
            app_html        TEXT,
            outlook_html    TEXT,
            app_edited_by   TEXT,
            app_edited_at   TEXT,
            outlook_modified_at TEXT,
            detected_at     TEXT,
            resolved_at     TEXT,
            resolved_by     TEXT,
            winner          TEXT,
            status          TEXT DEFAULT 'pending'
        );

        CREATE TABLE IF NOT EXISTS sync_errors (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            exchange_id  TEXT,
            display_name TEXT,
            error_type   TEXT,
            error_detail TEXT,
            first_seen   TEXT,
            last_seen    TEXT,
            resolved     INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS audit_log (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            contact_id   INTEGER,
            exchange_id  TEXT,
            display_name TEXT,
            changed_by   TEXT NOT NULL,
            change_type  TEXT NOT NULL,
            field_name   TEXT,
            old_value    TEXT,
            new_value    TEXT,
            changed_at   TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_audit_contact ON audit_log(contact_id);
        CREATE INDEX IF NOT EXISTS idx_audit_changed ON audit_log(changed_at DESC);

        CREATE TABLE IF NOT EXISTS users (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            username   TEXT UNIQUE NOT NULL,
            role       TEXT NOT NULL DEFAULT 'readonly',
            created_at TEXT,
            updated_at TEXT
        );
    """)
    conn.commit()

    # Migrations for existing databases — safe to run repeatedly
    migrations = [
        "ALTER TABLE contacts ADD COLUMN home_street TEXT",
        "ALTER TABLE contacts ADD COLUMN home_city TEXT",
        "ALTER TABLE contacts ADD COLUMN home_state TEXT",
        "ALTER TABLE contacts ADD COLUMN home_zip TEXT",
        "ALTER TABLE contacts ADD COLUMN home_country TEXT",
        "ALTER TABLE contacts ADD COLUMN other_street TEXT",
        "ALTER TABLE contacts ADD COLUMN other_city TEXT",
        "ALTER TABLE contacts ADD COLUMN other_state TEXT",
        "ALTER TABLE contacts ADD COLUMN other_zip TEXT",
        "ALTER TABLE contacts ADD COLUMN other_country TEXT",
        "ALTER TABLE contacts ADD COLUMN categories TEXT",
        "ALTER TABLE pending_writes ADD COLUMN field_type TEXT DEFAULT 'notes'",
        "ALTER TABLE pending_writes ADD COLUMN field_data TEXT",
        # audit_log and users are created in the main schema block above;
        # these no-ops ensure older DBs that pre-date the schema block get the tables too.
        """CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT, contact_id INTEGER,
            exchange_id TEXT, display_name TEXT, changed_by TEXT NOT NULL,
            change_type TEXT NOT NULL, field_name TEXT,
            old_value TEXT, new_value TEXT, changed_at TEXT NOT NULL)""",
        "CREATE INDEX IF NOT EXISTS idx_audit_contact ON audit_log(contact_id)",
        "CREATE INDEX IF NOT EXISTS idx_audit_changed ON audit_log(changed_at DESC)",
        """CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL,
            role TEXT NOT NULL DEFAULT 'readonly', created_at TEXT, updated_at TEXT)""",
        "ALTER TABLE sync_log ADD COLUMN sync_type TEXT",
    ]
    for sql in migrations:
        try:
            conn.execute(sql)
            conn.commit()
        except Exception:
            pass  # Column already exists

    conn.close()
    print("[DB] Database initialised.")
