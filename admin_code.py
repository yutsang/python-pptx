"""Today's admin code. Save it as the only content of a .txt file and drop that
file into the app's databook uploader: the browser session gets a
"Usage & logs" panel with the token report and the log zip.

The code is the date run through a salt, so it changes every day and can be
worked out on any machine with this checkout -- no need to reach the server.
Yesterday's and tomorrow's codes are accepted as well (time zones).
If config.yml sets admin.unlock_secret on the server, set the same value here.

    python admin_code.py
"""
from fdd_utils.usage_ledger import admin_code, admin_secret, admin_secret_is_custom

if __name__ == "__main__":
    print(admin_code(admin_secret()))
    print("(salt: " + ("admin.unlock_secret from config.yml -- the server must have the same one)"
                       if admin_secret_is_custom() else "built-in default)"))
