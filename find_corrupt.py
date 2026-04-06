from sync_engine import get_account
from exchangelib import Contact

account = get_account()
target = account.root / 'Top of Information Store' / 'Contacts'

print("Scanning for corrupt contacts...")
for c in target.all():
    if not isinstance(c, Contact):
        print(f"FOUND CORRUPT CONTACT:")
        print(f"  Type: {type(c).__name__}")
        print(f"  ID: {getattr(c, 'id', 'unknown')}")
        print(f"  Subject: {getattr(c, 'subject', 'unknown')}")
        print(f"  Created: {getattr(c, 'datetime_created', 'unknown')}")
        print(f"  Modified: {getattr(c, 'datetime_received', 'unknown')}")