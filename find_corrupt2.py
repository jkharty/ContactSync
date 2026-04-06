from sync_engine import get_account
from exchangelib import Contact

account = get_account()
target = account.root / 'Top of Information Store' / 'Contacts'

print("Scanning contacts...")
count = 0
try:
    for c in target.all():
        count += 1
        try:
            name = getattr(c, 'display_name', None) or getattr(c, 'subject', None) or '(no name)'
            ctype = type(c).__name__
            if ctype != 'Contact':
                print(f"#{count} CORRUPT: type={ctype} name={name}")
                print(f"  id={getattr(c, 'id', 'unknown')}")
                print(f"  created={getattr(c, 'datetime_created', 'unknown')}")
                print(f"  received={getattr(c, 'datetime_received', 'unknown')}")
        except Exception as e:
            print(f"#{count} ERROR reading contact: {e}")
except Exception as e:
    print(f"Iteration stopped at contact #{count}: {e}")
    print(f"Last contact type: {type(c).__name__}")

print(f"\nDone. Scanned {count} contacts.")