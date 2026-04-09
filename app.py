"""
app.py — Flask web application.
Serves the contact browser/editor to users on the local network.
"""
import os, datetime, threading, functools, json as _json
import pytz
from flask import (Flask, render_template, request, redirect,
                   url_for, session, jsonify, g)
import config
from database import get_db, init_db
from sync_engine import full_sync, run_scheduler, html_to_rtf

app = Flask(__name__)
app.secret_key = config.SECRET_KEY

# Mountain Time timezone for all timestamps
MT = pytz.timezone(config.TIMEZONE)

def now_iso():
    """Return current time in Mountain Time as ISO string."""
    return datetime.datetime.now(MT).isoformat()

# ── Easy Auth / Microsoft 365 login ───────────────────────────────────────────
_DOMAIN      = "invisionvail.com"
_ADMIN_EMAIL = "johnh@invisionvail.com"

@app.before_request
def load_azure_user():
    """
    When Azure Easy Auth is active, every request arrives with the header
    X-MS-CLIENT-PRINCIPAL-NAME set to the signed-in user's email address.

    On each request:
      - Domain-check the email (deny non-invisionvail.com addresses).
      - First login: insert user into the users table with role 'readonly'
        (or 'admin' for the designated admin email).
      - Subsequent requests: read role from the users table so that admin
        role changes take effect on the user's very next request.

    When running locally (no Easy Auth), the header is absent and this
    function is a no-op — the normal /login route handles authentication.
    """
    if request.path.startswith("/.auth"):
        return

    email = request.headers.get("X-MS-CLIENT-PRINCIPAL-NAME", "")
    if not email:
        return

    email_lower = email.lower()

    # Domain gate — deny anyone outside invisionvail.com.
    if email_lower != _ADMIN_EMAIL.lower() and not email_lower.endswith("@" + _DOMAIN):
        return "Access denied: your Microsoft account is not authorised for this application.", 403

    username  = email.split("@")[0]
    now       = now_iso()
    default_role = "admin" if email_lower == _ADMIN_EMAIL.lower() else "readonly"

    # Look up user in DB; create on first login.
    db = get_db()
    try:
        row = db.execute("SELECT role FROM users WHERE username=?", (username,)).fetchone()
        if row:
            # Use whatever role the admin has set in the DB.
            role = row["role"]
            db.execute(
                "UPDATE users SET last_login=?, email=? WHERE username=?",
                (now, email, username)
            )
        else:
            role = default_role
            db.execute(
                "INSERT INTO users (username, email, role, created_at, last_login) "
                "VALUES (?,?,?,?,?)",
                (username, email, role, now, now)
            )
        db.commit()
    finally:
        db.close()

    session["username"] = username
    session["email"]    = email
    session["role"]     = role

# ── Startup tasks (runs under both gunicorn and direct `python app.py`) ───────
# init_db() is safe to call multiple times — all tables use CREATE IF NOT EXISTS.
with app.app_context():
    init_db()
    # One-time migration: set display_name to company for contacts with no name
    with get_db() as _db:
        _db.execute("""
            UPDATE contacts SET display_name = company
            WHERE (display_name IS NULL OR display_name = '')
              AND company IS NOT NULL AND company != ''
        """)
        # Pre-populate users table so admins can manage roles before first login
        now = now_iso()
        for uname, cfg in config.USERS.items():
            _db.execute(
                "INSERT OR IGNORE INTO users (username, role, created_at) VALUES (?,?,?)",
                (uname, cfg["role"], now)
            )

# Start the background sync scheduler. Each gunicorn worker runs its own thread;
# the sync engine's database transactions prevent data corruption.
_scheduler_thread = threading.Thread(target=run_scheduler, daemon=True)
_scheduler_thread.start()

@app.context_processor
def inject_config():
    def get_sync_errors():
        db = get_request_db()
        return db.execute(
            "SELECT * FROM sync_errors WHERE resolved=0 ORDER BY last_seen DESC"
        ).fetchall()
    def get_pending_conflicts():
        db = get_request_db()
        return db.execute(
            "SELECT COUNT(*) FROM conflicts WHERE status='pending'"
        ).fetchone()[0]
    def get_sync_status():
        from sync_engine import get_sync_status as _gss
        return _gss()
    return {
        "config": config,
        "get_sync_errors": get_sync_errors,
        "get_pending_conflicts": get_pending_conflicts,
        "get_sync_status": get_sync_status,
    }

# ── Auth helpers ──────────────────────────────────────────────────────────────
def login_required(f):
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        if "username" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated

def role_required(*roles):
    def decorator(f):
        @functools.wraps(f)
        def decorated(*args, **kwargs):
            if session.get("role") not in roles:
                return render_template("error.html",
                    message="You don't have permission to do that."), 403
            return f(*args, **kwargs)
        return decorated
    return decorator

# ── DB per request ────────────────────────────────────────────────────────────
@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db: db.close()

def get_request_db():
    if "db" not in g:
        g.db = get_db()
    return g.db

# ── Audit helper ──────────────────────────────────────────────────────────────
def _audit(db, contact_id, exchange_id, display_name, changed_by,
           change_type, field_name=None, old_value=None, new_value=None):
    now = now_iso()
    db.execute("""
        INSERT INTO audit_log
          (contact_id, exchange_id, display_name, changed_by, change_type,
           field_name, old_value, new_value, changed_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (contact_id, exchange_id, display_name, changed_by, change_type,
          field_name,
          str(old_value) if old_value is not None else None,
          str(new_value) if new_value is not None else None,
          now))

# ── Shared contact WHERE builder ──────────────────────────────────────────────
_SORT_COLS = {
    "name":    "display_name COLLATE NOCASE",
    "company": "company COLLATE NOCASE",
    "title":   "job_title COLLATE NOCASE",
    "email":   "email1 COLLATE NOCASE",
}

def _build_word_prefix_pattern(query_words):
    """Build SQL pattern for word-prefix matching.

    For each word in the query, generate a pattern that matches it at the
    start of the field or after whitespace (word boundary).

    Example: ["Rob", "Clas"] matches "Robert B. Clasen" because:
      - "Robert B. Clasen" starts with "Rob" (matches Rob%)
      - "Robert B. Clasen" contains " Clas" (matches % Clas%)
    """
    patterns = []
    params = []
    for word in query_words:
        # Match at start OR after space
        patterns.append(f"(field LIKE ? OR field LIKE ?)")
        params.extend([f"{word}%", f"% {word}%"])
    return patterns, params

def _contact_where(query, category, is_admin):
    parts, params = [], []
    if not is_admin:
        parts.append("(categories IS NULL OR categories NOT LIKE '%PERSONAL%')")

    if query:
        # Split query into words for word-prefix matching
        query_words = query.split()

        if query_words:
            # Build word-prefix patterns for searchable fields
            fields = [
                "display_name", "company", "email1", "email2", "email3",
                "phone_business", "phone_mobile", "phone_home", "notes_plain",
                "address_street", "address_city", "address_state",
                "home_street", "home_city", "home_state",
                "other_street", "other_city", "other_state"
            ]

            field_conditions = []
            for field in fields:
                field_patterns = []
                field_params = []
                for word in query_words:
                    # Each word must match at start of field or after whitespace
                    field_patterns.append(f"({field} LIKE ? OR {field} LIKE ?)")
                    field_params.extend([f"{word}%", f"% {word}%"])

                # All words must match in this field
                if field_patterns:
                    field_condition = " AND ".join(field_patterns)
                    field_conditions.append(f"({field_condition})")
                    params.extend(field_params)

            # Any of the fields can match
            if field_conditions:
                parts.append("(" + " OR ".join(field_conditions) + ")")

    if category:
        parts.append("categories LIKE ?")
        params.append(f"%{category}%")

    where = ("WHERE " + " AND ".join(parts)) if parts else ""
    return where, params

def _available_cats(db, where_sql, params):
    """Return sorted list of categories visible under the given WHERE."""
    extra = "AND categories IS NOT NULL AND categories != ''"
    cat_where = (where_sql + " " + extra) if where_sql else "WHERE categories IS NOT NULL AND categories != ''"
    rows = db.execute(f"SELECT categories FROM contacts {cat_where}", params).fetchall()
    cat_set = set()
    for r in rows:
        for c in r["categories"].split("|"):
            c = c.strip()
            if c:
                cat_set.add(c)
    return sorted(cat_set)

# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/login", methods=["GET", "POST"])
def login():
    # If Easy Auth already set the session (M365 login), skip the login page.
    if "username" in session:
        return redirect(url_for("index"))
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user_cfg = config.USERS.get(username)
        if user_cfg and user_cfg["password"] == password:
            now = now_iso()
            db  = get_request_db()
            # Check DB for user — admin may have changed their role
            db_user = db.execute(
                "SELECT * FROM users WHERE username=?", (username,)
            ).fetchone()
            if db_user:
                role = db_user["role"]
            else:
                role = user_cfg["role"]
                db.execute(
                    "INSERT INTO users (username, role, created_at) VALUES (?,?,?)",
                    (username, role, now)
                )
                db.commit()
            session["username"] = username
            session["role"]     = role
            return redirect(url_for("index"))
        error = "Invalid username or password."
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    # When Easy Auth is active, /.auth/logout signs out of Microsoft too.
    # When running locally (no Easy Auth), it redirects harmlessly to /.auth/logout
    # which doesn't exist and falls back to the login page via the 404 handler —
    # so we redirect to / which login_required will redirect to /login.
    return redirect("/.auth/logout?post_logout_redirect_uri=/")

@app.route("/")
@login_required
def index():
    return redirect(url_for("contacts"))

@app.route("/contacts")
@login_required
def contacts():
    db       = get_request_db()
    query    = request.args.get("q", "").strip()
    category = request.args.get("cat", "").strip()
    sort_key = request.args.get("sort", "name")
    sort_dir = request.args.get("dir", "asc")
    page     = max(1, int(request.args.get("page", 1)))
    per      = 50
    off      = (page - 1) * per
    is_admin = session["role"] == "admin"

    where_sql, params = _contact_where(query, category, is_admin)
    sort_col  = _SORT_COLS.get(sort_key, "display_name COLLATE NOCASE")
    sort_dir_sql = "DESC" if sort_dir == "desc" else "ASC"

    rows = db.execute(f"""
        SELECT id, display_name, company, job_title, email1,
               phone_business, phone_mobile, phone_home, categories,
               home_street, home_city, home_state,
               address_street, address_city, address_state,
               other_street, other_city, other_state
        FROM contacts {where_sql}
        ORDER BY {sort_col} {sort_dir_sql}
        LIMIT ? OFFSET ?
    """, params + [per, off]).fetchall()

    total = db.execute(
        f"SELECT COUNT(*) FROM contacts {where_sql}", params
    ).fetchone()[0]

    pages = (total + per - 1) // per
    return render_template("contacts.html",
        contacts=rows, query=query, category=category,
        sort_key=sort_key, sort_dir=sort_dir,
        page=page, pages=pages, total=total,
        available_categories=_available_cats(db, where_sql, params),
        role=session["role"], username=session["username"])

@app.route("/contacts/<int:cid>")
@login_required
def contact_detail(cid):
    db  = get_request_db()
    row = db.execute("SELECT * FROM contacts WHERE id=?", (cid,)).fetchone()
    if not row:
        return render_template("error.html", message="Contact not found."), 404
    return render_template("contact.html", c=row,
        role=session["role"], username=session["username"])

@app.route("/contacts/<int:cid>/edit", methods=["GET", "POST"])
@login_required
@role_required("viewedit", "admin")
def edit_contact(cid):
    db  = get_request_db()
    row = db.execute("SELECT * FROM contacts WHERE id=?", (cid,)).fetchone()
    if not row:
        return render_template("error.html", message="Contact not found."), 404

    if request.method == "POST":
        new_html = request.form.get("notes_html", "")
        new_rtf  = html_to_rtf(new_html).encode("latin-1", errors="replace")
        now      = now_iso()
        # Update local DB immediately
        db.execute("""
            UPDATE contacts SET notes_html=?, notes_rtf=?, synced_at=?
            WHERE id=?
        """, (new_html, new_rtf, now, cid))
        # Queue write-back to Exchange
        db.execute("""
            INSERT INTO pending_writes
              (exchange_id, notes_html, notes_rtf, requested_by, requested_at, status)
            VALUES (?, ?, ?, ?, ?, 'pending')
        """, (row["exchange_id"], new_html, new_rtf,
              session["username"], now))
        db.commit()
        return redirect(url_for("contact_detail", cid=cid))

    return render_template("edit.html", c=row,
        role=session["role"], username=session["username"])

# ── Edit contact fields (identity / contact info / address / categories / notes)
@app.route("/contacts/<int:cid>/edit-fields", methods=["GET", "POST"])
@login_required
@role_required("viewedit", "admin")
def edit_fields(cid):
    import json
    db  = get_request_db()
    row = db.execute("SELECT * FROM contacts WHERE id=?", (cid,)).fetchone()
    if not row:
        return render_template("error.html", message="Contact not found."), 404

    if request.method == "POST":
        now     = now_iso()
        f       = request.form
        section = f.get("active_section", "identity")

        if section == "notes":
            new_html = f.get("notes_html", "")
            new_rtf  = html_to_rtf(new_html).encode("latin-1", errors="replace")
            _audit(db, cid, row["exchange_id"], row["display_name"],
                   session["username"], "edit_notes", "notes_html",
                   row["notes_html"] or "", new_html)
            db.execute("""
                UPDATE contacts SET notes_html=?, notes_rtf=?, synced_at=? WHERE id=?
            """, (new_html, new_rtf, now, cid))
            db.execute("""
                INSERT INTO pending_writes
                  (exchange_id, notes_html, notes_rtf, requested_by, requested_at, status)
                VALUES (?, ?, ?, ?, ?, 'pending')
            """, (row["exchange_id"], new_html, new_rtf, session["username"], now))

        else:
            # identity / contact / address — categories submitted from identity tab dropdown
            cats    = [v.strip() for v in f.getlist("category") if v.strip()]
            cat_str = "|".join(cats)
            field_data = {
                "first_name":      f.get("first_name", "").strip(),
                "last_name":       f.get("last_name", "").strip(),
                "company":         f.get("company", "").strip(),
                "job_title":       f.get("job_title", "").strip(),
                "email1":          f.get("email1", "").strip(),
                "email2":          f.get("email2", "").strip(),
                "email3":          f.get("email3", "").strip(),
                "phone_business":  f.get("phone_business", "").strip(),
                "phone_mobile":    f.get("phone_mobile", "").strip(),
                "phone_home":      f.get("phone_home", "").strip(),
                "address_street":  f.get("address_street", "").strip(),
                "address_city":    f.get("address_city", "").strip(),
                "address_state":   f.get("address_state", "").strip(),
                "address_zip":     f.get("address_zip", "").strip(),
                "address_country": f.get("address_country", "").strip(),
                "home_street":     f.get("home_street", "").strip(),
                "home_city":       f.get("home_city", "").strip(),
                "home_state":      f.get("home_state", "").strip(),
                "home_zip":        f.get("home_zip", "").strip(),
                "home_country":    f.get("home_country", "").strip(),
                "other_street":    f.get("other_street", "").strip(),
                "other_city":      f.get("other_city", "").strip(),
                "other_state":     f.get("other_state", "").strip(),
                "other_zip":       f.get("other_zip", "").strip(),
                "other_country":   f.get("other_country", "").strip(),
                "categories":      cats,
            }
            first = field_data["first_name"]
            last  = field_data["last_name"]
            company = field_data.get("company", "").strip()
            display_name = f"{first} {last}".strip() if (first or last) else (company or row["display_name"])
            db.execute("""
                UPDATE contacts SET
                    first_name=:first_name, last_name=:last_name, display_name=:display_name,
                    company=:company, job_title=:job_title,
                    email1=:email1, email2=:email2, email3=:email3,
                    phone_business=:phone_business, phone_mobile=:phone_mobile, phone_home=:phone_home,
                    address_street=:address_street, address_city=:address_city,
                    address_state=:address_state, address_zip=:address_zip, address_country=:address_country,
                    home_street=:home_street, home_city=:home_city,
                    home_state=:home_state, home_zip=:home_zip, home_country=:home_country,
                    other_street=:other_street, other_city=:other_city,
                    other_state=:other_state, other_zip=:other_zip, other_country=:other_country,
                    categories=:categories, synced_at=:synced_at
                WHERE id=:id
            """, {**field_data, "categories": cat_str,
                  "display_name": display_name, "synced_at": now, "id": cid})
            # Audit per-changed-field
            field_map = {
                "first_name": "First name", "last_name": "Last name",
                "company": "Company", "job_title": "Job title",
                "email1": "Email 1", "email2": "Email 2", "email3": "Email 3",
                "phone_business": "Business phone", "phone_mobile": "Mobile",
                "phone_home": "Home phone", "address_street": "Street (Business)",
                "address_city": "City (Business)", "address_state": "State (Business)",
                "address_zip": "ZIP (Business)", "address_country": "Country (Business)",
                "home_street": "Street (Home)", "home_city": "City (Home)",
                "home_state": "State (Home)", "home_zip": "ZIP (Home)",
                "home_country": "Country (Home)", "other_street": "Street (Other)",
                "other_city": "City (Other)", "other_state": "State (Other)",
                "other_zip": "ZIP (Other)", "other_country": "Country (Other)",
            }
            for fld, label in field_map.items():
                old_v = row[fld] or ""
                new_v = field_data.get(fld, "")
                if old_v != new_v:
                    _audit(db, cid, row["exchange_id"], display_name,
                           session["username"], "edit_fields", label, old_v, new_v)
            old_cats = row["categories"] or ""
            if old_cats != cat_str:
                _audit(db, cid, row["exchange_id"], display_name,
                       session["username"], "edit_fields", "Categories", old_cats, cat_str)
            db.execute("""
                INSERT INTO pending_writes
                  (exchange_id, field_type, field_data, requested_by, requested_at, status)
                VALUES (?, 'fields', ?, ?, ?, 'pending')
            """, (row["exchange_id"], _json.dumps(field_data), session["username"], now))

        db.commit()
        return redirect(url_for("contact_detail", cid=cid))

    # GET — collect all categories known to the system
    section = request.args.get("section", "identity")
    cat_rows = db.execute(
        "SELECT categories FROM contacts WHERE categories IS NOT NULL AND categories != ''"
    ).fetchall()
    cat_set = set()
    for r in cat_rows:
        for cat in r["categories"].split("|"):
            if cat.strip():
                cat_set.add(cat.strip())
    # Also include the contact's own categories (in case they're unique)
    if row["categories"]:
        for cat in row["categories"].split("|"):
            if cat.strip():
                cat_set.add(cat.strip())
    all_categories = sorted(cat_set)

    return render_template("edit_fields.html", c=row, section=section,
        role=session["role"], username=session["username"],
        all_categories=all_categories)

# ── Create new contact ───────────────────────────────────────────────────────
@app.route("/contacts/new", methods=["GET", "POST"])
@login_required
@role_required("viewedit", "admin")
def new_contact():
    import uuid, json as _json
    db = get_request_db()

    if request.method == "POST":
        f    = request.form
        cats = [v.strip() for v in f.getlist("category") if v.strip()]
        cat_str  = "|".join(cats)
        first    = f.get("first_name", "").strip()
        last     = f.get("last_name", "").strip()
        company  = f.get("company", "").strip()
        display_name = f"{first} {last}".strip() if (first or last) else (company or "New Contact")
        now      = now_iso()
        local_id = f"LOCAL-{uuid.uuid4()}"

        field_data = {
            "first_name": first, "last_name": last, "company": company,
            "job_title":       f.get("job_title", "").strip(),
            "email1":          f.get("email1", "").strip(),
            "email2":          f.get("email2", "").strip(),
            "email3":          f.get("email3", "").strip(),
            "phone_business":  f.get("phone_business", "").strip(),
            "phone_mobile":    f.get("phone_mobile", "").strip(),
            "phone_home":      f.get("phone_home", "").strip(),
            "address_street":  f.get("address_street", "").strip(),
            "address_city":    f.get("address_city", "").strip(),
            "address_state":   f.get("address_state", "").strip(),
            "address_zip":     f.get("address_zip", "").strip(),
            "address_country": f.get("address_country", "").strip(),
            "home_street":     f.get("home_street", "").strip(),
            "home_city":       f.get("home_city", "").strip(),
            "home_state":      f.get("home_state", "").strip(),
            "home_zip":        f.get("home_zip", "").strip(),
            "home_country":    f.get("home_country", "").strip(),
            "other_street":    f.get("other_street", "").strip(),
            "other_city":      f.get("other_city", "").strip(),
            "other_state":     f.get("other_state", "").strip(),
            "other_zip":       f.get("other_zip", "").strip(),
            "other_country":   f.get("other_country", "").strip(),
            "categories":      cats,
            "notes_html":      f.get("notes_html", ""),
        }

        db.execute("""
            INSERT INTO contacts (
                exchange_id, first_name, last_name, display_name, company, job_title,
                email1, email2, email3, phone_business, phone_mobile, phone_home,
                address_street, address_city, address_state, address_zip, address_country,
                home_street, home_city, home_state, home_zip, home_country,
                other_street, other_city, other_state, other_zip, other_country,
                categories, notes_html, synced_at
            ) VALUES (
                :exchange_id, :first_name, :last_name, :display_name, :company, :job_title,
                :email1, :email2, :email3, :phone_business, :phone_mobile, :phone_home,
                :address_street, :address_city, :address_state, :address_zip, :address_country,
                :home_street, :home_city, :home_state, :home_zip, :home_country,
                :other_street, :other_city, :other_state, :other_zip, :other_country,
                :categories, :notes_html, :synced_at
            )
        """, {**field_data, "exchange_id": local_id, "display_name": display_name,
              "categories": cat_str, "synced_at": now})

        new_id = db.execute("SELECT last_insert_rowid()").fetchone()[0]
        db.execute("""
            INSERT INTO pending_writes
              (exchange_id, field_type, field_data, requested_by, requested_at, status)
            VALUES (?, 'create', ?, ?, ?, 'pending')
        """, (local_id, _json.dumps(field_data), session["username"], now))
        _audit(db, new_id, local_id, display_name,
               session["username"], "create", None, None, display_name)
        db.commit()
        return redirect(url_for("contact_detail", cid=new_id))

    # GET — empty form
    cat_rows = db.execute(
        "SELECT categories FROM contacts WHERE categories IS NOT NULL AND categories != ''"
    ).fetchall()
    cat_set = set()
    for r in cat_rows:
        for cat in r["categories"].split("|"):
            if cat.strip():
                cat_set.add(cat.strip())
    empty = {k: "" for k in [
        "id", "display_name", "first_name", "last_name", "company", "job_title",
        "email1", "email2", "email3", "phone_business", "phone_mobile", "phone_home",
        "address_street", "address_city", "address_state", "address_zip", "address_country",
        "home_street", "home_city", "home_state", "home_zip", "home_country",
        "other_street", "other_city", "other_state", "other_zip", "other_country",
        "categories", "notes_html",
    ]}
    return render_template("edit_fields.html", c=empty, section="identity",
        role=session["role"], username=session["username"],
        all_categories=sorted(cat_set), new_contact=True)

# ── Delete contact ────────────────────────────────────────────────────────────
@app.route("/contacts/<int:cid>/delete", methods=["POST"])
@login_required
@role_required("admin")
def delete_contact(cid):
    from sync_engine import delete_contact as _delete
    db  = get_request_db()
    row = db.execute(
        "SELECT exchange_id, display_name FROM contacts WHERE id=?", (cid,)
    ).fetchone()
    if not row:
        return render_template("error.html", message="Contact not found."), 404
    errors = _delete(row["exchange_id"], row["display_name"], session["username"])
    if errors:
        return render_template("error.html",
            message=f"Delete partially failed: {'; '.join(errors)}"), 500
    return redirect(url_for("contacts"))

# ── Conflict routes ───────────────────────────────────────────────────────────
@app.route("/admin/conflicts")
@login_required
@role_required("admin")
def conflicts():
    db  = get_request_db()
    pending  = db.execute(
        "SELECT * FROM conflicts WHERE status='pending' ORDER BY detected_at DESC"
    ).fetchall()
    resolved = db.execute(
        "SELECT * FROM conflicts WHERE status='resolved' ORDER BY resolved_at DESC LIMIT 20"
    ).fetchall()
    return render_template("conflicts.html", pending=pending, resolved=resolved,
        role=session["role"], username=session["username"])

@app.route("/admin/conflicts/<int:cid>/resolve", methods=["POST"])
@login_required
@role_required("admin")
def resolve_conflict(cid):
    from sync_engine import resolve_conflict as _resolve
    winner = request.form.get("winner")
    if winner not in ("app", "outlook"):
        return jsonify({"error": "Invalid winner"}), 400
    success = _resolve(cid, winner, session["username"])
    return jsonify({"status": "resolved" if success else "error"})

# ── Sync status API ────────────────────────────────────────────────────────────
@app.route("/api/sync-status")
@login_required
def api_sync_status():
    from sync_engine import get_sync_status
    status = get_sync_status()
    db = get_request_db()
    conflicts = db.execute(
        "SELECT COUNT(*) FROM conflicts WHERE status='pending'"
    ).fetchone()[0]
    errors = db.execute(
        "SELECT COUNT(*) FROM sync_errors WHERE resolved=0"
    ).fetchone()[0]
    status["pending_conflicts"] = conflicts
    status["sync_errors"] = errors
    return jsonify(status)

# ── Admin routes ──────────────────────────────────────────────────────────────
@app.route("/admin")
@login_required
@role_required("admin", "viewedit")
def admin():
    db      = get_request_db()
    tab     = request.args.get("tab", "overview")
    logs    = db.execute("SELECT * FROM sync_log ORDER BY started_at DESC LIMIT 50").fetchall()
    total   = db.execute("SELECT COUNT(*) FROM contacts").fetchone()[0]
    pending = db.execute("SELECT COUNT(*) FROM pending_writes WHERE status='pending'").fetchone()[0]
    users   = db.execute("SELECT * FROM users ORDER BY username COLLATE NOCASE").fetchall()
    if session["role"] == "viewedit":
        # viewedit cannot see audit entries for contacts categorised as personal
        history = db.execute("""
            SELECT al.* FROM audit_log al
            LEFT JOIN contacts c ON al.contact_id = c.id
            WHERE c.categories IS NULL OR c.categories NOT LIKE '%personal%'
            ORDER BY al.changed_at DESC LIMIT 200
        """).fetchall()
    else:
        history = db.execute("""
            SELECT * FROM audit_log ORDER BY changed_at DESC LIMIT 200
        """).fetchall()
    errors  = db.execute(
        "SELECT * FROM sync_errors WHERE resolved=0 ORDER BY last_seen DESC"
    ).fetchall()
    conflicts_pending  = db.execute(
        "SELECT * FROM conflicts WHERE status='pending' ORDER BY detected_at DESC"
    ).fetchall()
    conflicts_resolved = db.execute(
        "SELECT * FROM conflicts WHERE status='resolved' ORDER BY resolved_at DESC LIMIT 20"
    ).fetchall()
    return render_template("admin.html",
        logs=logs, total=total, pending=pending,
        users=users, history=history, errors=errors,
        conflicts_pending=conflicts_pending, conflicts_resolved=conflicts_resolved, tab=tab,
        role=session["role"], username=session["username"])

@app.route("/admin/sync", methods=["POST"])
@login_required
@role_required("admin")
def trigger_sync():
    threading.Thread(target=full_sync, daemon=True).start()
    return jsonify({"status": "Sync started in background."})

@app.route("/admin/users/<int:uid>/role", methods=["POST"])
@login_required
@role_required("admin")
def change_user_role(uid):
    new_role = request.form.get("role", "").strip()
    if new_role not in ("readonly", "viewedit", "admin"):
        return redirect(url_for("admin", tab="users"))
    db  = get_request_db()
    now = now_iso()
    # Prevent self-demotion
    target = db.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    if target and target["username"] == session["username"] and new_role != "admin":
        pass  # silently ignore self-demotion
    else:
        db.execute("UPDATE users SET role=?, updated_at=? WHERE id=?", (new_role, now, uid))
        db.commit()
    return redirect(url_for("admin", tab="users"))

@app.route("/admin/users/<int:uid>/delete", methods=["POST"])
@login_required
@role_required("admin")
def delete_user(uid):
    db = get_request_db()
    target = db.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()
    # Cannot delete yourself
    if target and target["username"] != session["username"]:
        db.execute("DELETE FROM users WHERE id=?", (uid,))
        db.commit()
    return redirect(url_for("admin", tab="users"))

@app.route("/admin/errors/<int:eid>/resolve", methods=["POST"])
@login_required
@role_required("admin")
def resolve_error(eid):
    db = get_request_db()
    db.execute("UPDATE sync_errors SET resolved=1 WHERE id=?", (eid,))
    db.commit()
    return redirect(url_for("admin", tab="errors"))

@app.route("/admin/writes/retry", methods=["POST"])
@login_required
@role_required("admin")
def retry_failed_writes():
    db = get_request_db()
    db.execute("UPDATE pending_writes SET status='pending' WHERE status='error'")
    db.commit()
    return redirect(url_for("admin", tab="errors"))

# ── API for search typeahead ──────────────────────────────────────────────────
@app.route("/api/search")
@login_required
def api_search():
    q        = request.args.get("q", "").strip()
    like     = f"%{q}%"
    db       = get_request_db()
    is_admin = session["role"] == "admin"
    personal = "" if is_admin else "AND (categories IS NULL OR categories NOT LIKE '%PERSONAL%')"
    rows = db.execute(f"""
        SELECT id, display_name, company, email1
        FROM contacts
        WHERE (display_name LIKE ? OR company LIKE ? OR email1 LIKE ?
               OR email2 LIKE ? OR phone_business LIKE ? OR phone_mobile LIKE ?)
        {personal}
        ORDER BY display_name COLLATE NOCASE LIMIT 10
    """, (like, like, like, like, like, like)).fetchall()
    return jsonify([dict(r) for r in rows])

# ── API for contacts list (AJAX, live search) ─────────────────────────────────
@app.route("/api/contacts")
@login_required
def api_contacts():
    db       = get_request_db()
    query    = request.args.get("q", "").strip()
    category = request.args.get("cat", "").strip()
    sort_key = request.args.get("sort", "name")
    sort_dir = request.args.get("dir", "asc")
    page     = max(1, int(request.args.get("page", 1)))
    per      = 50
    off      = (page - 1) * per
    is_admin = session["role"] == "admin"

    where_sql, params = _contact_where(query, category, is_admin)
    sort_col     = _SORT_COLS.get(sort_key, "display_name COLLATE NOCASE")
    sort_dir_sql = "DESC" if sort_dir == "desc" else "ASC"

    rows = db.execute(f"""
        SELECT id, display_name, company, job_title, email1,
               phone_business, phone_mobile, phone_home, categories,
               home_street, home_city, home_state,
               address_street, address_city, address_state,
               other_street, other_city, other_state
        FROM contacts {where_sql}
        ORDER BY {sort_col} {sort_dir_sql}
        LIMIT ? OFFSET ?
    """, params + [per, off]).fetchall()

    total = db.execute(
        f"SELECT COUNT(*) FROM contacts {where_sql}", params
    ).fetchone()[0]

    return jsonify({
        "contacts":            [dict(r) for r in rows],
        "total":               total,
        "page":                page,
        "pages":               (total + per - 1) // per,
        "sort_key":            sort_key,
        "sort_dir":            sort_dir,
        "available_categories": _available_cats(db, where_sql, params),
    })

# ── Bulk category assign ──────────────────────────────────────────────────────
@app.route("/api/contacts/bulk-assign", methods=["POST"])
@login_required
@role_required("viewedit", "admin")
def bulk_assign():
    data       = request.get_json(force=True)
    ids        = [int(i) for i in data.get("ids", [])]
    add_cat    = data.get("category", "").strip()
    if not ids or not add_cat:
        return jsonify({"error": "Missing ids or category"}), 400

    db  = get_request_db()
    now = now_iso()
    updated = 0
    for cid in ids:
        row = db.execute(
            "SELECT id, exchange_id, display_name, categories FROM contacts WHERE id=?",
            (cid,)
        ).fetchone()
        if not row:
            continue
        existing = [c.strip() for c in (row["categories"] or "").split("|") if c.strip()]
        if add_cat not in existing:
            existing.append(add_cat)
        new_cats = "|".join(existing)
        db.execute("UPDATE contacts SET categories=?, synced_at=? WHERE id=?",
                   (new_cats, now, cid))
        db.execute("""
            INSERT INTO pending_writes
              (exchange_id, field_type, field_data, requested_by, requested_at, status)
            VALUES (?, 'categories', ?, ?, ?, 'pending')
        """, (row["exchange_id"], _json.dumps({"categories": existing}),
              session["username"], now))
        _audit(db, cid, row["exchange_id"], row["display_name"],
               session["username"], "bulk_assign", "categories",
               row["categories"], new_cats)
        updated += 1
    db.commit()
    return jsonify({"updated": updated})

# ── Startup ───────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print(f"\n* Contact Sync running at http://localhost:{config.WEB_PORT}")
    print(f"  Access from other PCs at http://YOUR-PC-NAME:{config.WEB_PORT}\n")
    app.run(host="0.0.0.0", port=config.WEB_PORT, debug=False)
