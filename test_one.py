from sync_engine import get_account, contact_to_dict
from database import get_db, init_db

init_db()
account = get_account()
target = account.root / 'Top of Information Store' / 'Contacts'

# Try just the first contact
for c in target.all()[:1]:
    print("Got contact:", c.display_name)
    try:
        d = contact_to_dict(c)
        print("Dict keys:", list(d.keys()))
        print("display_name:", d["display_name"])
        print("exchange_id:", d["exchange_id"][:40])
        db = get_db()
        db.execute("""
            INSERT OR REPLACE INTO contacts (
                exchange_id, change_key, first_name, last_name, display_name,
                company, job_title, email1, email2, email3,
                phone_business, phone_mobile, phone_home,
                address_street, address_city, address_state,
                address_zip, address_country,
                notes_rtf, notes_html, notes_plain, last_modified, synced_at
            ) VALUES (
                :exchange_id, :change_key, :first_name, :last_name, :display_name,
                :company, :job_title, :email1, :email2, :email3,
                :phone_business, :phone_mobile, :phone_home,
                :address_street, :address_city, :address_state,
                :address_zip, :address_country,
                :notes_rtf, :notes_html, :notes_plain, :last_modified, :synced_at
            )
        """, d)
        db.commit()
        print("SUCCESS - contact saved to database")
        count = db.execute("SELECT COUNT(*) FROM contacts").fetchone()[0]
        print(f"Total in DB: {count}")
        db.close()
    except Exception as e:
        import traceback
        print("ERROR:", e)
        traceback.print_exc()