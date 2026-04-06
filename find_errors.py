from sync_engine import get_account, contact_to_dict
from database import get_db

account = get_account()
target = account.root / 'Top of Information Store' / 'Contacts'
db = get_db()
errors = []

for c in target.all():
    try:
        contact_to_dict(c)
    except Exception as e:
        errors.append((c.display_name, str(e)))

for name, err in errors:
    print(f'{name}: {err}')

print(f'Total errors: {len(errors)}')