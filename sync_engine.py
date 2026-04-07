"""
sync_engine.py — connects to Exchange via EWS (OAuth2), reads contacts,
detects conflicts, writes edits back, and maintains the local database.
"""
import datetime
import logging
import re
import config
from database import get_db
from exchangelib import (
    Account, Configuration, ExtendedProperty, Contact, DELEGATE, IMPERSONATION,
    HTMLBody, EWSDateTime, UTC
)
from exchangelib.credentials import OAuth2AuthorizationCodeCredentials

logging.basicConfig(level=logging.INFO,
    format="%(asctime)s [SYNC] %(levelname)s %(message)s")
log = logging.getLogger("sync")

class RtfBody(ExtendedProperty):
    property_tag  = 0x1009
    property_type = "Binary"

Contact.register("rtf_body", RtfBody)

# ── Token cache ───────────────────────────────────────────────────────────────
_cached_account = None

from html.parser import HTMLParser
import html as _html_mod

class _QuillToRTF(HTMLParser):
    """
    Converts Quill-generated HTML to simple RTF that Outlook's contact
    notes viewer can render. Supports bold, italic, underline, strikethrough,
    inline colors, paragraphs, line breaks, headings, lists, and tables.
    Keeps RTF intentionally minimal — Outlook's contact notes parser is basic.
    """
    _H_SIZES = {"h1": 40, "h2": 32, "h3": 28}

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self._colors = []      # (r,g,b) tuples; 1-based index in RTF color table
        self._parts  = []      # RTF fragments
        self._stack  = []      # tag stack for tracking open formatting
        self._td_sep = False   # whether next <td> should prepend \tab

    # ── color table ───────────────────────────────────────────────────────────
    def _ci(self, hex_color):
        hex_color = hex_color.lstrip("#")
        if len(hex_color) != 6:
            return 0
        rgb = (int(hex_color[0:2],16), int(hex_color[2:4],16), int(hex_color[4:6],16))
        if rgb not in self._colors:
            self._colors.append(rgb)
        return self._colors.index(rgb) + 1   # 1-based

    # ── RTF escaping ──────────────────────────────────────────────────────────
    @staticmethod
    def _esc(text):
        out = []
        for ch in text:
            if   ch == "\\": out.append("\\\\")
            elif ch == "{":  out.append("\\{")
            elif ch == "}":  out.append("\\}")
            elif ch == "\n": out.append("\\line\n")
            elif ord(ch) > 127:
                out.append(f"\\u{ord(ch)}?")
            else:
                out.append(ch)
        return "".join(out)

    # ── tag handlers ─────────────────────────────────────────────────────────
    def handle_starttag(self, tag, attrs):
        tag  = tag.lower()
        attr = dict(attrs)
        p    = self._parts
        stk  = self._stack

        if tag == "p":
            p.append("\\pard ")
            stk.append("p")

        elif tag in self._H_SIZES:
            sz = self._H_SIZES[tag]
            p.append(f"\\pard\\b\\fs{sz} ")
            stk.append(tag)

        elif tag in ("strong", "b"):
            p.append("\\b ")
            stk.append("b")

        elif tag in ("em", "i"):
            p.append("\\i ")
            stk.append("i")

        elif tag == "u":
            p.append("\\ul ")
            stk.append("u")

        elif tag in ("s", "strike", "del"):
            p.append("\\strike ")
            stk.append("s")

        elif tag == "span":
            style     = attr.get("style", "")
            # handle rgb() or #hex in color
            color_m   = re.search(r'color\s*:\s*rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', style)
            hex_m     = re.search(r'color\s*:\s*(#[0-9A-Fa-f]{6})', style)
            if color_m:
                r,g,b = int(color_m.group(1)), int(color_m.group(2)), int(color_m.group(3))
                idx = self._ci(f"{r:02X}{g:02X}{b:02X}")
                p.append(f"\\cf{idx} ")
                stk.append(("span","cf"))
            elif hex_m:
                idx = self._ci(hex_m.group(1))
                p.append(f"\\cf{idx} ")
                stk.append(("span","cf"))
            else:
                stk.append(("span",""))

        elif tag == "br":
            p.append("\\line\n")

        elif tag == "ul":
            stk.append("ul")

        elif tag == "ol":
            stk.append("ol")

        elif tag == "li":
            p.append("\\pard\\li360\\bullet\\tx360 ")
            stk.append("li")

        elif tag == "table":
            stk.append("table")

        elif tag == "tr":
            p.append("\\pard ")
            self._td_sep = False
            stk.append("tr")

        elif tag in ("td", "th"):
            if self._td_sep:
                p.append("\\tab ")
            self._td_sep = True
            stk.append(tag)

        elif tag == "a":
            stk.append("a")    # render text only

    def handle_endtag(self, tag):
        tag = tag.lower()
        p   = self._parts
        stk = self._stack

        if tag == "p":
            p.append("\\par\n")
            self._pop(stk, "p")

        elif tag in self._H_SIZES:
            p.append(f"\\b0\\fs24\\par\n")
            self._pop(stk, tag)

        elif tag in ("strong", "b"):
            p.append("\\b0 ")
            self._pop(stk, "b")

        elif tag in ("em", "i"):
            p.append("\\i0 ")
            self._pop(stk, "i")

        elif tag == "u":
            p.append("\\ulnone ")
            self._pop(stk, "u")

        elif tag in ("s", "strike", "del"):
            p.append("\\strike0 ")
            self._pop(stk, "s")

        elif tag == "span":
            # find and pop matching span entry
            for i in range(len(stk)-1, -1, -1):
                if isinstance(stk[i], tuple) and stk[i][0] == "span":
                    _, kind = stk.pop(i)
                    if kind == "cf":
                        p.append("\\cf0 ")
                    break

        elif tag == "li":
            p.append("\\par\n")
            self._pop(stk, "li")

        elif tag == "tr":
            p.append("\\par\n")
            self._td_sep = False
            self._pop(stk, "tr")

    def handle_data(self, data):
        self._parts.append(self._esc(data))

    @staticmethod
    def _pop(stack, val):
        for i in range(len(stack)-1, -1, -1):
            if stack[i] == val:
                stack.pop(i)
                return

    # ── assemble final RTF ────────────────────────────────────────────────────
    def get_rtf(self):
        if self._colors:
            entries   = "".join(f"\\red{r}\\green{g}\\blue{b};" for r,g,b in self._colors)
            colortbl  = "{\\colortbl ;" + entries + "}"
        else:
            colortbl  = "{\\colortbl ;}"
        body = "".join(self._parts)
        return (
            "{\\rtf1\\ansi\\deff0\n"
            "{\\fonttbl{\\f0\\fswiss\\fcharset0 Calibri;}}\n"
            + colortbl + "\n"
            "\\f0\\fs24\\pard\n"
            + body
            + "}"
        )


def html_to_rtf(html_string):
    """Convert Quill HTML to RTF using the native parser (no pandoc required)."""
    if not html_string or not html_string.strip():
        return r"{\rtf1\ansi }"
    parser = _QuillToRTF()
    parser.feed(html_string)
    return parser.get_rtf()


def _compress_rtf(rtf_str):
    """
    Compress RTF string to Exchange's LZFu format (PR_RTF_COMPRESSED).
    Returns compressed bytes, or None if compressed-rtf is not installed.
    Install with: pip install compressed-rtf
    """
    try:
        import compressed_rtf as crtf
        rtf_bytes = rtf_str.encode("latin-1", errors="replace")
        return crtf.compress(rtf_bytes)
    except Exception:
        return None

def get_account():
    global _cached_account
    if _cached_account is not None:
        return _cached_account

    # Use client credentials (app-only) flow — no browser required.
    # Requires the App Registration to have the Office 365 Exchange Online
    # application permission: full_access_as_app (admin-consented).
    from msal import ConfidentialClientApplication
    msal_app = ConfidentialClientApplication(
        client_id         = config.CLIENT_ID,
        client_credential = config.CLIENT_SECRET,
        authority         = f"https://login.microsoftonline.com/{config.TENANT_ID}",
    )
    result = msal_app.acquire_token_for_client(
        scopes=["https://outlook.office365.com/.default"]
    )
    if "access_token" not in result:
        raise Exception(f"Auth failed: {result.get('error_description', result)}")

    credentials = OAuth2AuthorizationCodeCredentials(
        client_id     = config.CLIENT_ID,
        client_secret = config.CLIENT_SECRET,
        tenant_id     = config.TENANT_ID,
        access_token  = result,
    )
    cfg = Configuration(
        server      = "outlook.office365.com",
        credentials = credentials,
        auth_type   = "OAuth 2.0",
    )
    account = Account(
        primary_smtp_address = config.MAILBOX_EMAIL,
        config               = cfg,
        autodiscover         = False,
        access_type          = IMPERSONATION,
    )
    _cached_account = account
    return account

def contact_to_dict(c):
    def s(v):
        return str(v).strip() if v else ""
    def ea(i):
        return s(c.email_addresses[i].email) if c.email_addresses and len(c.email_addresses) > i else ""
    def ph(k):
        if c.phone_numbers:
            for p in c.phone_numbers:
                if p.label == k:
                    return s(p.phone_number)
        return ""
    def addr(k):
        if c.physical_addresses:
            for a in c.physical_addresses:
                if a.label == k:
                    return a
        return None

    biz   = addr("Business")
    home  = addr("Home")
    other = addr("Other")
    # primary address for backward compat — prefer business, fall back to home/other
    primary = biz or home or other

    rtf_raw = bytes(c.rtf_body) if c.rtf_body else b""

    html_str = str(c.body) if c.body else ""
    body_match = re.search(r"<body[^>]*>(.*?)</body>", html_str, re.DOTALL | re.IGNORECASE)
    if body_match:
        html_str = body_match.group(1).strip()

    cats = ""
    if c.categories:
        cats = "|".join(str(cat) for cat in c.categories)
        log.debug(f"Categories for {s(c.display_name)}: {cats}")

    return {
        "exchange_id":     s(c.id),
        "change_key":      s(c.changekey),
        "first_name":      s(c.given_name),
        "last_name":       s(c.surname),
        "display_name":    s(c.display_name),
        "company":         s(c.company_name),
        "job_title":       s(c.job_title),
        "email1":          ea(0),
        "email2":          ea(1),
        "email3":          ea(2),
        "phone_business":  ph("BusinessPhone"),
        "phone_mobile":    ph("MobilePhone"),
        "phone_home":      ph("HomePhone"),
        "address_street":  s(biz.street)   if biz   else (s(primary.street)  if primary else ""),
        "address_city":    s(biz.city)     if biz   else (s(primary.city)    if primary else ""),
        "address_state":   s(biz.state)    if biz   else (s(primary.state)   if primary else ""),
        "address_zip":     s(biz.zipcode)  if biz   else (s(primary.zipcode) if primary else ""),
        "address_country": s(biz.country)  if biz   else (s(primary.country) if primary else ""),
        "home_street":     s(home.street)  if home  else "",
        "home_city":       s(home.city)    if home  else "",
        "home_state":      s(home.state)   if home  else "",
        "home_zip":        s(home.zipcode) if home  else "",
        "home_country":    s(home.country) if home  else "",
        "other_street":    s(other.street)  if other else "",
        "other_city":      s(other.city)    if other else "",
        "other_state":     s(other.state)   if other else "",
        "other_zip":       s(other.zipcode) if other else "",
        "other_country":   s(other.country) if other else "",
        "categories":      cats,
        "notes_rtf":       rtf_raw,
        "notes_html":      html_str,
        "notes_plain":     s(c.body),
        "last_modified":   str(c.last_modified_time) if c.last_modified_time else "",
        "synced_at":       datetime.datetime.utcnow().isoformat(),
    }

# ── Sync status helpers ───────────────────────────────────────────────────────
def _set_sync_status(state, message=""):
    try:
        db = get_db()
        now = datetime.datetime.utcnow().isoformat()
        if state == "running":
            db.execute("UPDATE sync_status SET state=?, started_at=?, last_message=? WHERE id=1",
                       (state, now, message))
        else:
            db.execute("UPDATE sync_status SET state=?, finished_at=?, last_message=? WHERE id=1",
                       (state, now, message))
        db.commit()
        db.close()
    except Exception:
        pass

def get_sync_status():
    try:
        db = get_db()
        row = db.execute("SELECT * FROM sync_status WHERE id=1").fetchone()
        db.close()
        return dict(row) if row else {}
    except Exception:
        return {}

# ── Conflict detection ────────────────────────────────────────────────────────
def _check_and_handle_conflict(db, d, pending_write):
    """
    Called when Exchange shows a contact has changed BUT we also have a
    pending write for it. This is a conflict — store both versions for admin.
    Returns True if conflict detected (skip normal update), False otherwise.
    """
    now = datetime.datetime.utcnow().isoformat()

    # Check if there's already an unresolved conflict for this contact
    existing = db.execute(
        "SELECT id FROM conflicts WHERE exchange_id=? AND status='pending'",
        (d["exchange_id"],)
    ).fetchone()

    if existing:
        # Update the Outlook version in the existing conflict
        db.execute(
            "UPDATE conflicts SET outlook_html=?, outlook_modified_at=?, detected_at=? WHERE id=?",
            (d["notes_html"], d["last_modified"], now, existing["id"])
        )
        return True

    # Create new conflict record
    db.execute("""
        INSERT INTO conflicts
          (exchange_id, display_name, app_html, outlook_html,
           app_edited_by, app_edited_at, outlook_modified_at, detected_at, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'pending')
    """, (
        d["exchange_id"],
        d["display_name"],
        pending_write["notes_html"],
        d["notes_html"],
        pending_write["requested_by"],
        pending_write["requested_at"],
        d["last_modified"],
        now,
    ))
    log.warning(f"Conflict detected for: {d['display_name']}")
    return True

# ── Full sync ─────────────────────────────────────────────────────────────────
def full_sync():
    log.info("Starting full sync ...")
    _set_sync_status("running", "Full sync in progress…")
    start = datetime.datetime.utcnow().isoformat()
    added = updated = errors = 0
    try:
        account  = get_account()
        target   = account.root / 'Top of Information Store' / 'Contacts'
        contacts = target.all()
        db = get_db()
        for c in contacts:
            try:
                if not isinstance(c, Contact):
                    error_type = type(c).__name__
                    error_id   = getattr(c, 'id', 'unknown')
                    log.warning(f"Skipping corrupt contact: {error_type}")
                    _log_contact_error(db, str(error_id), "(unknown — MIME corrupt)", error_type,
                                       "Exchange could not convert this contact's MIME content. "
                                       "Open and re-save this contact in Outlook to fix it.")
                    errors += 1
                    continue
                d = contact_to_dict(c)
                existing = db.execute(
                    "SELECT id, change_key FROM contacts WHERE exchange_id=?",
                    (d["exchange_id"],)
                ).fetchone()

                if existing:
                    if existing["change_key"] != d["change_key"]:
                        # Check for pending write on this contact (conflict check)
                        pending = db.execute(
                            "SELECT * FROM pending_writes WHERE exchange_id=? AND status='pending'",
                            (d["exchange_id"],)
                        ).fetchone()
                        if pending:
                            _check_and_handle_conflict(db, d, pending)
                        else:
                            db.execute("""
                                UPDATE contacts SET
                                    change_key=:change_key, first_name=:first_name,
                                    last_name=:last_name, display_name=:display_name,
                                    company=:company, job_title=:job_title,
                                    email1=:email1, email2=:email2, email3=:email3,
                                    phone_business=:phone_business, phone_mobile=:phone_mobile,
                                    phone_home=:phone_home, address_street=:address_street,
                                    address_city=:address_city, address_state=:address_state,
                                    address_zip=:address_zip, address_country=:address_country,
                                    home_street=:home_street, home_city=:home_city,
                                    home_state=:home_state, home_zip=:home_zip,
                                    home_country=:home_country, other_street=:other_street,
                                    other_city=:other_city, other_state=:other_state,
                                    other_zip=:other_zip, other_country=:other_country,
                                    categories=:categories,
                                    notes_rtf=:notes_rtf, notes_html=:notes_html,
                                    notes_plain=:notes_plain, last_modified=:last_modified,
                                    synced_at=:synced_at
                                WHERE exchange_id=:exchange_id
                            """, d)
                            updated += 1
                    else:
                        # change_key matches — contact unchanged in Exchange, but backfill
                        # categories in case this contact was synced before categories support
                        if d["categories"]:
                            db.execute(
                                "UPDATE contacts SET categories=:categories WHERE exchange_id=:exchange_id",
                                {"categories": d["categories"], "exchange_id": d["exchange_id"]}
                            )
                else:
                    db.execute("""
                        INSERT INTO contacts (
                            exchange_id, change_key, first_name, last_name, display_name,
                            company, job_title, email1, email2, email3,
                            phone_business, phone_mobile, phone_home,
                            address_street, address_city, address_state, address_zip, address_country,
                            home_street, home_city, home_state, home_zip, home_country,
                            other_street, other_city, other_state, other_zip, other_country,
                            categories, notes_rtf, notes_html, notes_plain, last_modified, synced_at
                        ) VALUES (
                            :exchange_id, :change_key, :first_name, :last_name, :display_name,
                            :company, :job_title, :email1, :email2, :email3,
                            :phone_business, :phone_mobile, :phone_home,
                            :address_street, :address_city, :address_state, :address_zip, :address_country,
                            :home_street, :home_city, :home_state, :home_zip, :home_country,
                            :other_street, :other_city, :other_state, :other_zip, :other_country,
                            :categories, :notes_rtf, :notes_html, :notes_plain, :last_modified, :synced_at
                        )
                    """, d)
                    added += 1
            except Exception as e:
                log.error(f"Error processing contact: {e}")
                errors += 1
        _process_pending_writes(db, account)
        db.commit()
        db.close()
    except Exception as e:
        log.error(f"Full sync failed: {e}")
        errors += 1

    finish = datetime.datetime.utcnow().isoformat()
    msg = f"Full sync complete — added:{added} updated:{updated} errors:{errors}"
    _set_sync_status("idle", msg)
    _log_sync(start, finish, added, updated, 0, errors, "full")
    log.info(f"Full sync done — added:{added} updated:{updated} errors:{errors}")

def force_refresh():
    log.info("Force refresh starting ...")
    _set_sync_status("running", "Force refresh in progress…")
    start = datetime.datetime.utcnow().isoformat()
    updated = errors = 0
    try:
        account  = get_account()
        target   = account.root / 'Top of Information Store' / 'Contacts'
        contacts = target.all()
        db = get_db()
        for c in contacts:
            try:
                if not isinstance(c, Contact):
                    continue
                d = contact_to_dict(c)
                db.execute("""
                    UPDATE contacts SET
                        change_key=:change_key, first_name=:first_name,
                        last_name=:last_name, display_name=:display_name,
                        company=:company, job_title=:job_title,
                        email1=:email1, email2=:email2, email3=:email3,
                        phone_business=:phone_business, phone_mobile=:phone_mobile,
                        phone_home=:phone_home, address_street=:address_street,
                        address_city=:address_city, address_state=:address_state,
                        address_zip=:address_zip, address_country=:address_country,
                        home_street=:home_street, home_city=:home_city,
                        home_state=:home_state, home_zip=:home_zip,
                        home_country=:home_country, other_street=:other_street,
                        other_city=:other_city, other_state=:other_state,
                        other_zip=:other_zip, other_country=:other_country,
                        categories=:categories,
                        notes_rtf=:notes_rtf, notes_html=:notes_html,
                        notes_plain=:notes_plain, last_modified=:last_modified,
                        synced_at=:synced_at
                    WHERE exchange_id=:exchange_id
                """, d)
                updated += 1
            except Exception as e:
                log.error(f"Error refreshing contact: {e}")
                errors += 1
        db.commit()
        db.close()
    except Exception as e:
        log.error(f"Force refresh failed: {e}")
        errors += 1

    finish = datetime.datetime.utcnow().isoformat()
    _set_sync_status("idle", f"Force refresh complete — updated:{updated} errors:{errors}")
    _log_sync(start, finish, 0, updated, 0, errors, "force_refresh")
    log.info(f"Force refresh done — updated:{updated} errors:{errors}")

# ── Incremental sync ──────────────────────────────────────────────────────────
def incremental_sync():
    log.info("Incremental sync ...")
    _set_sync_status("running", "Incremental sync in progress…")
    start = datetime.datetime.utcnow().isoformat()
    added = updated = errors = 0
    try:
        db  = get_db()
        row = db.execute("SELECT MAX(last_modified) as lm FROM contacts").fetchone()
        since_str = row["lm"] if row and row["lm"] else "2000-01-01T00:00:00"
        since_naive = datetime.datetime.fromisoformat(since_str[:19])
        since_dt    = EWSDateTime.from_datetime(since_naive.replace(tzinfo=UTC))

        account  = get_account()
        target   = account.root / 'Top of Information Store' / 'Contacts'
        contacts = target.filter(last_modified_time__gt=since_dt)

        for c in contacts:
            try:
                if not isinstance(c, Contact):
                    error_type = type(c).__name__
                    error_id   = getattr(c, 'id', 'unknown')
                    _log_contact_error(db, str(error_id), "(unknown — MIME corrupt)", error_type,
                                       "Exchange could not convert this contact's MIME content. "
                                       "Open and re-save this contact in Outlook to fix it.")
                    errors += 1
                    continue
                d = contact_to_dict(c)
                existing = db.execute(
                    "SELECT id FROM contacts WHERE exchange_id=?",
                    (d["exchange_id"],)
                ).fetchone()
                if existing:
                    # Conflict check
                    pending = db.execute(
                        "SELECT * FROM pending_writes WHERE exchange_id=? AND status='pending'",
                        (d["exchange_id"],)
                    ).fetchone()
                    if pending:
                        _check_and_handle_conflict(db, d, pending)
                    else:
                        db.execute("""
                            UPDATE contacts SET
                                change_key=:change_key, first_name=:first_name,
                                last_name=:last_name, display_name=:display_name,
                                company=:company, job_title=:job_title,
                                email1=:email1, email2=:email2, email3=:email3,
                                phone_business=:phone_business, phone_mobile=:phone_mobile,
                                phone_home=:phone_home, address_street=:address_street,
                                address_city=:address_city, address_state=:address_state,
                                address_zip=:address_zip, address_country=:address_country,
                                home_street=:home_street, home_city=:home_city,
                                home_state=:home_state, home_zip=:home_zip,
                                home_country=:home_country, other_street=:other_street,
                                other_city=:other_city, other_state=:other_state,
                                other_zip=:other_zip, other_country=:other_country,
                                categories=:categories,
                                notes_rtf=:notes_rtf, notes_html=:notes_html,
                                notes_plain=:notes_plain, last_modified=:last_modified,
                                synced_at=:synced_at
                            WHERE exchange_id=:exchange_id
                        """, d)
                        updated += 1
                else:
                    db.execute("""
                        INSERT INTO contacts (
                            exchange_id, change_key, first_name, last_name, display_name,
                            company, job_title, email1, email2, email3,
                            phone_business, phone_mobile, phone_home,
                            address_street, address_city, address_state, address_zip, address_country,
                            home_street, home_city, home_state, home_zip, home_country,
                            other_street, other_city, other_state, other_zip, other_country,
                            categories, notes_rtf, notes_html, notes_plain, last_modified, synced_at
                        ) VALUES (
                            :exchange_id, :change_key, :first_name, :last_name, :display_name,
                            :company, :job_title, :email1, :email2, :email3,
                            :phone_business, :phone_mobile, :phone_home,
                            :address_street, :address_city, :address_state, :address_zip, :address_country,
                            :home_street, :home_city, :home_state, :home_zip, :home_country,
                            :other_street, :other_city, :other_state, :other_zip, :other_country,
                            :categories, :notes_rtf, :notes_html, :notes_plain, :last_modified, :synced_at
                        )
                    """, d)
                    added += 1
            except Exception as e:
                log.error(f"Error: {e}")
                errors += 1

        _process_pending_writes(db, account)
        db.commit()
        db.close()
    except Exception as e:
        log.error(f"Incremental sync error: {e}")
        errors += 1

    finish = datetime.datetime.utcnow().isoformat()
    msg = f"Last sync: added:{added} updated:{updated} errors:{errors}"
    _set_sync_status("idle", msg)
    _log_sync(start, finish, added, updated, 0, errors, "incremental")

# ── Write-back ────────────────────────────────────────────────────────────────
def _write_fields_to_exchange(c, field_data):
    """Apply a field_data dict (from pending_writes) to an exchangelib Contact object."""
    import json
    from exchangelib import EmailAddress, PhoneNumber, PhysicalAddress
    data = json.loads(field_data)

    update_fields = []
    if "first_name" in data:
        c.given_name = data["first_name"] or None
        update_fields.append("given_name")
    if "last_name" in data:
        c.surname = data["last_name"] or None
        update_fields.append("surname")
    if "company" in data:
        c.company_name = data["company"] or None
        update_fields.append("company_name")
    if "job_title" in data:
        c.job_title = data["job_title"] or None
        update_fields.append("job_title")

    if any(k in data for k in ("email1", "email2", "email3")):
        emails = []
        for label, key in [("EmailAddress1","email1"),("EmailAddress2","email2"),("EmailAddress3","email3")]:
            val = data.get(key, "")
            if val:
                emails.append(EmailAddress(email=val, label=label))
        c.email_addresses = emails or None
        update_fields.append("email_addresses")

    phone_keys = ("phone_business", "phone_mobile", "phone_home")
    if any(k in data for k in phone_keys):
        phones = []
        for label, key in [("BusinessPhone","phone_business"),("MobilePhone","phone_mobile"),("HomePhone","phone_home")]:
            val = data.get(key, "")
            if val:
                phones.append(PhoneNumber(phone_number=val, label=label))
        c.phone_numbers = phones or None
        update_fields.append("phone_numbers")

    addr_prefixes = [("Business","address"), ("Home","home"), ("Other","other")]
    if any(f"{p}_{s}" in data for _,p in addr_prefixes for s in ("street","city","state","zip","country")):
        addrs = []
        for label, prefix in addr_prefixes:
            street = data.get(f"{prefix}_street", "")
            if street:
                addrs.append(PhysicalAddress(
                    street=street,
                    city=data.get(f"{prefix}_city") or None,
                    state=data.get(f"{prefix}_state") or None,
                    zipcode=data.get(f"{prefix}_zip") or None,
                    country=data.get(f"{prefix}_country") or None,
                    label=label,
                ))
        c.physical_addresses = addrs or None
        update_fields.append("physical_addresses")

    if "categories" in data:
        cats = data["categories"]  # list
        c.categories = cats if cats else None
        update_fields.append("categories")

    if update_fields:
        c.save(update_fields=update_fields)


def _process_pending_writes(db, account):
    from exchangelib import ItemId, HTMLBody
    pending = db.execute(
        "SELECT * FROM pending_writes WHERE status='pending'"
    ).fetchall()
    for row in pending:
        # Skip if there's an unresolved conflict — wait for admin decision
        conflict = db.execute(
            "SELECT id FROM conflicts WHERE exchange_id=? AND status='pending'",
            (row["exchange_id"],)
        ).fetchone()
        if conflict:
            log.info(f"Skipping write-back — conflict pending for {row['exchange_id'][:30]}")
            continue
        try:
            items = list(account.fetch(ids=[ItemId(row["exchange_id"])]))
            if not items:
                raise ValueError("Contact not found in Exchange")
            c = items[0]

            field_type = row["field_type"] if row["field_type"] else "notes"

            if field_type == "fields":
                _write_fields_to_exchange(c, row["field_data"])
                log.info(f"Wrote back fields: {row['exchange_id'][:40]}")
            elif field_type == "categories":
                import json as _json
                data = _json.loads(row["field_data"])
                cats = data.get("categories", [])
                c.categories = cats if cats else None
                c.save(update_fields=["categories"])
                log.info(f"Wrote back categories: {row['exchange_id'][:40]}")
            else:
                html = row["notes_html"] or ""
                rtf_str    = html_to_rtf(html)
                compressed = _compress_rtf(rtf_str)
                if compressed:
                    c.rtf_body = compressed
                    c.save(update_fields=["rtf_body"])
                    log.info(f"Wrote back (RTF): {row['exchange_id'][:40]}")
                else:
                    c.body = HTMLBody(f"<html><body>{html}</body></html>")
                    c.save(update_fields=["body"])
                    log.info(f"Wrote back (HTML fallback): {row['exchange_id'][:40]}")

            db.execute(
                "UPDATE pending_writes SET status='done' WHERE id=?", (row["id"],)
            )
        except Exception as e:
            log.error(f"Write-back failed: {e}")
            db.execute(
                "UPDATE pending_writes SET status='error' WHERE id=?", (row["id"],)
            )

def delete_contact(exchange_id, display_name, deleted_by):
    """Delete a contact from both Exchange and the local database."""
    from exchangelib import ItemId
    now = datetime.datetime.utcnow().isoformat()
    errors = []

    # Delete from Exchange
    try:
        account = get_account()
        items = list(account.fetch(ids=[ItemId(exchange_id)]))
        if items:
            items[0].delete()
            log.info(f"Deleted from Exchange: {display_name} (by {deleted_by})")
        else:
            log.warning(f"Contact not found in Exchange during delete: {exchange_id[:40]}")
    except Exception as e:
        log.error(f"Failed to delete from Exchange: {e}")
        errors.append(str(e))

    # Delete from local database regardless (even if Exchange delete failed)
    try:
        db = get_db()
        db.execute("DELETE FROM contacts WHERE exchange_id=?", (exchange_id,))
        db.execute("DELETE FROM pending_writes WHERE exchange_id=?", (exchange_id,))
        db.execute("DELETE FROM conflicts WHERE exchange_id=?", (exchange_id,))
        db.execute(
            "UPDATE sync_errors SET resolved=1 WHERE exchange_id=?", (exchange_id,)
        )
        db.execute("""
            INSERT INTO sync_log (started_at, finished_at, added, updated, deleted, errors, message)
            VALUES (?, ?, 0, 0, 1, ?, ?)
        """, (now, now, len(errors), f"delete:{display_name} by {deleted_by}"))
        db.commit()
        db.close()
        log.info(f"Deleted from local DB: {display_name}")
    except Exception as e:
        log.error(f"Failed to delete from local DB: {e}")
        errors.append(str(e))

    return errors  # empty list = success

def resolve_conflict(conflict_id, winner, resolved_by):
    """
    Called when admin picks a winner. winner = 'app' or 'outlook'.
    Applies the winning version and cleans up.
    """
    db = get_db()
    conflict = db.execute(
        "SELECT * FROM conflicts WHERE id=?", (conflict_id,)
    ).fetchone()
    if not conflict:
        db.close()
        return False

    now = datetime.datetime.utcnow().isoformat()

    if winner == "app":
        # Keep app version — mark pending write as still pending so it gets pushed
        winning_html = conflict["app_html"]
        log.info(f"Conflict {conflict_id}: app version wins")
    else:
        # Keep Outlook version — cancel the pending write and update local DB
        winning_html = conflict["outlook_html"]
        db.execute(
            "UPDATE pending_writes SET status='cancelled' WHERE exchange_id=? AND status='pending'",
            (conflict["exchange_id"],)
        )
        db.execute(
            "UPDATE contacts SET notes_html=? WHERE exchange_id=?",
            (winning_html, conflict["exchange_id"])
        )
        log.info(f"Conflict {conflict_id}: Outlook version wins")

    db.execute("""
        UPDATE conflicts SET
            status='resolved', resolved_at=?, resolved_by=?, winner=?
        WHERE id=?
    """, (now, resolved_by, winner, conflict_id))
    db.commit()
    db.close()

    if winner == "app":
        # Push app version to Exchange now
        try:
            account = get_account()
            from exchangelib import ItemId, HTMLBody
            items = list(account.fetch(ids=[ItemId(conflict["exchange_id"])]))
            if items:
                rtf_str    = html_to_rtf(winning_html)
                compressed = _compress_rtf(rtf_str)
                if compressed:
                    items[0].rtf_body = compressed
                    items[0].save(update_fields=["rtf_body"])
                else:
                    items[0].body = HTMLBody(f"<html><body>{winning_html}</body></html>")
                    items[0].save(update_fields=["body"])
                inner_db = get_db()
                inner_db.execute(
                    "UPDATE pending_writes SET status='done' WHERE exchange_id=? AND status='pending'",
                    (conflict["exchange_id"],)
                )
                inner_db.commit()
                inner_db.close()
        except Exception as e:
            log.error(f"Conflict resolution write-back failed: {e}")
    return True

# ── Helpers ───────────────────────────────────────────────────────────────────
def _log_contact_error(db, exchange_id, display_name, error_type, error_detail):
    now = datetime.datetime.utcnow().isoformat()
    existing = db.execute(
        "SELECT id FROM sync_errors WHERE exchange_id=? AND resolved=0",
        (exchange_id,)
    ).fetchone()
    if existing:
        db.execute(
            "UPDATE sync_errors SET last_seen=?, error_detail=? WHERE id=?",
            (now, error_detail, existing["id"])
        )
    else:
        db.execute("""
            INSERT INTO sync_errors
              (exchange_id, display_name, error_type, error_detail, first_seen, last_seen, resolved)
            VALUES (?, ?, ?, ?, ?, ?, 0)
        """, (exchange_id, display_name, error_type, error_detail, now, now))

def _html_to_plain(html):
    text = re.sub(r"<br\s*/?>", "\n", html, flags=re.IGNORECASE)
    text = re.sub(r"<p[^>]*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    return text.strip()

def _log_sync(start, finish, added, updated, deleted, errors, mode):
    try:
        db = get_db()
        db.execute("""
            INSERT INTO sync_log (started_at, finished_at, added, updated, deleted, errors, message)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (start, finish, added, updated, deleted, errors, mode))
        db.commit()
        db.close()
    except Exception:
        pass

def run_scheduler():
    import time
    log.info(f"Sync scheduler started. Interval: {config.SYNC_INTERVAL_SECONDS}s")
    while True:
        try:
            incremental_sync()
        except Exception as e:
            log.error(f"Scheduler error: {e}")
        time.sleep(config.SYNC_INTERVAL_SECONDS)
