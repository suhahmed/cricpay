"""
cricpay account recovery from the command line (e.g. a PythonAnywhere Bash console).

  python manage_users.py list
  python manage_users.py reset-password <username>        # asks for the new password
  python manage_users.py make-admin <username>            # also activates the account
  python manage_users.py create-admin <username> "<Full name>"

Run it from the cricket-tracker folder. No Reload needed: accounts are read on every request.
"""
import getpass
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import auth  # noqa: E402

DATA = (os.path.dirname(os.path.abspath(os.environ["TRACKER_XLSX"])) if os.environ.get("TRACKER_XLSX")
        else os.path.join(os.path.dirname(os.path.abspath(__file__)), "data"))
auth._cfg.update(users_file=os.path.join(DATA, "cricpay-users.json"),
                 settings_file=os.path.join(DATA, "cricpay-settings.json"))


def ask_password():
    while True:
        p1 = getpass.getpass("New password: ")
        if len(p1) < auth.MIN_PASSWORD:
            print(f"At least {auth.MIN_PASSWORD} characters, please.")
            continue
        if p1 == getpass.getpass("Type it again: "):
            return p1
        print("They don't match, try again.")


def main(argv):
    if len(argv) < 2 or argv[1] not in ("list", "reset-password", "make-admin", "create-admin"):
        print(__doc__)
        return 1
    d = auth._users()
    cmd = argv[1]
    if cmd == "list":
        for u in d["users"]:
            acc = "ALL (admin)" if u["role"] == "admin" else ", ".join(f"{s}: {r}" for s, r in u.get("access", {}).items()) or "-"
            print(f"{u['username']:<20} {u['role']:<7} {u['status']:<9} {u['name']:<22} {acc}")
        if not d["users"]:
            print("(no accounts yet)")
        return 0
    if len(argv) < 3:
        print("Which username?")
        return 1
    username = argv[2].strip().lower()
    u = auth._find(d, username)
    if cmd == "create-admin":
        if u:
            print("That username already exists. Use make-admin / reset-password instead.")
            return 1
        name = argv[3] if len(argv) > 3 else username
        u = {"username": username, "name": name, "role": "admin", "status": "active", "access": {},
             "created": auth._now(), "last_login": None}
        auth._set_password(u, ask_password())
        d["users"].append(u)
    elif not u:
        print("No such user. Run: python manage_users.py list")
        return 1
    elif cmd == "reset-password":
        auth._set_password(u, ask_password())
    elif cmd == "make-admin":
        u["role"], u["status"] = "admin", "active"
    auth._save_users(d)
    print(f"Done: {u['username']} is {u['role']} ({u['status']}).")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
