from sync_engine import get_account
a = get_account()
target = a.root / 'Top of Information Store' / 'Contacts'
print(f"Folder: {target.absolute}")
print(f"Total items: {target.total_count}")
print("First 3 items:")
for i, c in enumerate(target.all()[:3]):
    print(f"  {i+1}. {c}")
    if i >= 2:
        break