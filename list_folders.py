from sync_engine import get_account
a = get_account()
for f in a.root.walk():
    if 'contact' in str(f.absolute).lower():
        print(f.absolute)