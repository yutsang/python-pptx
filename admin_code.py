"""Today's admin code: save it in a .txt and drop that file into the app's
databook uploader to unlock the usage report and log download for that
browser session.

The code is derived from fdd_utils/config.yml -> admin.unlock_secret (or the
FDD_ADMIN_SECRET environment variable) and today's date, so it changes daily.
Run this on the SERVER's checkout (same config.yml the app reads):

    python admin_code.py
"""
from fdd_utils.usage_ledger import admin_code, admin_secret

if __name__ == "__main__":
    secret = admin_secret()
    if not secret:
        print("Admin unlock is OFF: no admin.unlock_secret in fdd_utils/config.yml.\n"
              "Add, for example:\n\n  admin:\n    unlock_secret: \"<any long private phrase>\"\n\n"
              "then run this again (no restart needed).")
        raise SystemExit(1)
    print(admin_code(secret))
