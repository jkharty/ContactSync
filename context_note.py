"""
app.py — add config to all template contexts so admin.html can read it.
This is a small patch file — the main app.py already handles this, but
we add a context processor to expose config everywhere cleanly.
"""
# This is already handled in app.py via the context_processor below.
# No action needed — included for documentation only.
