from sync_engine import get_account
from database import get_db

db = get_db()
row = db.execute("SELECT * FROM pending_writes WHERE status='pending'").fetchone()
print("Pending write for:", row["exchange_id"][:40])

account = get_account()
target = account.root / 'Top of Information Store' / 'Contacts'

try:
    # Find the contact by ID
    from exchangelib import EWSDateTime
    contacts = list(target.filter(id=row["exchange_id"]))
    print(f"Found {len(contacts)} contacts")
except Exception as e:
    print(f"filter by id failed: {e}")

try:
    # Alternative: fetch directly by item ID
    from exchangelib.items import Contact
    from exchangelib import ItemId
    item = account.fetch(ids=[row["exchange_id"]])[0]
    print(f"fetch() worked: {item.display_name}")
    
    # Now try writing
    rtf_bytes = row["notes_rtf"]
    item.body = item.body  # keep existing
    item.save(update_fields=["body"])
    print("Save succeeded!")
    
except Exception as e:
    import traceback
    print(f"fetch failed: {e}")
    traceback.print_exc()