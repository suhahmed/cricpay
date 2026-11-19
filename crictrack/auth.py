"""
CricTrack accounts: register / login / roles / per-tournament access.

Roles
  admin    the whole website: every tournament (read + write), "All tournaments" overview,
           manages every account, full Excel download
  captain  per tournament: read + write that tournament, manages which players can see it
  player   per tournament: read only

A user record:  role = "admin" | "member",  access = {"<sheet name>": "captain" | "player"}

* Everything needs a signed-in, active user.
* New sign-ups are "pending" until an admin or a captain gives them access to a tournament.
* The very first account becomes the admin. If an edit PIN was set before
  (data/crictrack-settings.json), that PIN is required to create it.
* Users live in data/crictrack-users.json; passwords are salted PBKDF2 hashes.
"""
import datetime as dt
import hashlib
import hmac
import json
import os
import re
import secrets
import threading
import time

from flask import abort, g, jsonify, redirect, request, send_from_directory, session

TOURNAMENT_ROLES = ("captain", "player")
USERNAME_RE = re.compile(r"^[a-z0-9][a-z0-9._-]{2,29}$")
MIN_PASSWORD = 6
_lock = threading.RLock()
_fails = {}          # key -> [count, locked_until]
_cfg = {"sheets": lambda: []}


# --------------------------------------------------------------------------- #
# storage
# --------------------------------------------------------------------------- #
def _read_json(path, default):
    try:
        with open(path) as f:
            return json.load(f)
    except Exception:
        return default


def _write_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w") as f:
        json.dump(data, f, indent=2)
    os.replace(tmp, path)


def _users():
    d = _read_json(_cfg["users_file"], {"users": [], "allow_register": True})
    changed = False
    for u in d["users"]:                          # migrate accounts from the 2-role version
        if u.get("role") == "captain":
            u["role"], changed = "admin", True
        elif u.get("role") == "viewer":
            u["role"], changed = "member", True
            u.setdefault("access", {s: "player" for s in _cfg["sheets"]()})
        if "access" not in u:
            u["access"], changed = {}, True
    if changed:
        _write_json(_cfg["users_file"], d)
    return d


def _save_users(d):
    _write_json(_cfg["users_file"], d)


def _find(d, username):
    username = (username or "").strip().lower()
    return next((u for u in d["users"] if u["username"] == username), None)


def _hash(pw, salt):
    return hashlib.pbkdf2_hmac("sha256", pw.encode(), bytes.fromhex(salt), 200_000).hex()


def _set_password(u, pw):
    u["salt"] = secrets.token_hex(16)
    u["pw"] = _hash(pw, u["salt"])
    u["sv"] = secrets.token_hex(4)      # session version: changing it signs the user out everywhere


def _check_password(u, pw):
    if not u:
        _hash(pw or "x", "00" * 16)       # same work for unknown users (no timing hint)
        return False
    return hmac.compare_digest(_hash(pw or "", u["salt"]), u["pw"])


def _now():
    return dt.datetime.now().isoformat(timespec="seconds")


def public(u, only_sheets=None):
    out = {k: u.get(k) for k in ("username", "name", "role", "status", "created", "last_login", "note")}
    acc = u.get("access", {})
    out["access"] = {s: r for s, r in acc.items() if only_sheets is None or s in only_sheets}
    out["other_access"] = 0 if only_sheets is None else len([s for s in acc if s not in only_sheets])
    return out


def _admins(d):
    return [u for u in d["users"] if u["role"] == "admin" and u["status"] == "active"]


# --------------------------------------------------------------------------- #
# permission helpers used by app.py
# --------------------------------------------------------------------------- #
def current_user():
    return getattr(g, "user", None)


def is_admin(u=None):
    u = u or current_user()
    return bool(u and u.get("role") == "admin")


def access(sheet, u=None):
    """'admin' | 'captain' | 'player' | None for this user on this tournament sheet."""
    u = u or current_user()
    if not u:
        return None
    if is_admin(u):
        return "admin"
    return u.get("access", {}).get(sheet)


def can_read(sheet):
    return bool(access(sheet))


def can_write(sheet):
    return access(sheet) in ("admin", "captain")


def captain_sheets(u=None):
    u = u or current_user()
    return [s for s, r in (u or {}).get("access", {}).items() if r == "captain"]


def can_create():
    return is_admin() or bool(captain_sheets())


def can_manage():
    return is_admin() or bool(captain_sheets())


def require_admin():
    if not is_admin():
        abort(403, "Only an admin can do that.")


def require_write(sheet=None):
    if sheet is None:
        return require_admin()
    if not can_write(sheet):
        abort(403, "Only this tournament's captain (or an admin) can make changes. You have view-only access.")


def require_read(sheet):
    if not can_read(sheet):
        abort(403, "You don't have access to this tournament.")


def require_create():
    if not can_create():
        abort(403, "Only an admin or a captain can create tournaments.")


def grant(username, sheet, role):
    with _lock:
        d = _users()
        u = _find(d, username)
        if u and not is_admin(u):
            u.setdefault("access", {})[sheet] = role
            if u["status"] == "pending":
                u["status"] = "active"
            _save_users(d)


def auth_fields():
    u = current_user()
    out = {"me": public(u) if u else None, "is_admin": is_admin(), "can_create": can_create(),
           "can_manage": can_manage(), "can_edit": is_admin(), "pin_required": False}
    if can_manage():
        out["pending_users"] = sum(1 for x in _users()["users"] if x["status"] == "pending")
    return out


def _throttle(key):
    cnt, until = _fails.get(key, [0, 0])
    if until > time.time():
        abort(429, f"Too many attempts. Try again in {int(until - time.time()) + 1} seconds.")


def _fail(key):
    cnt, _ = _fails.get(key, [0, 0])
    cnt += 1
    _fails[key] = [0, time.time() + 60] if cnt >= 5 else [cnt, 0]


def _ip():
    return request.headers.get("X-Forwarded-For", request.remote_addr or "?").split(",")[0].strip()


def _body():
    return request.get_json(force=True, silent=True) or {}


def _setup_pin_ok(pin):
    """Before any account exists, a previously-set edit PIN protects admin creation."""
    st = _read_json(_cfg["settings_file"], {})
    env_pin = os.environ.get("EDIT_PIN", "").strip()
    if env_pin:
        return hmac.compare_digest(str(pin or "").strip(), env_pin)
    if not st.get("pin_hash"):
        return True
    return hmac.compare_digest(_hash(str(pin or "").strip(), st["salt"]), st["pin_hash"])


def _pin_is_set():
    return bool(os.environ.get("EDIT_PIN", "").strip() or _read_json(_cfg["settings_file"], {}).get("pin_hash"))


def _validate_new(name, username, password):
    name = re.sub(r"\s+", " ", str(name or "")).strip()[:60]
    username = str(username or "").strip().lower()
    if not name:
        abort(400, "Enter the name")
    if not USERNAME_RE.match(username):
        abort(400, "Username: 3–30 characters, letters/numbers and . _ - only (e.g. suhail or suhail.a)")
    if len(str(password or "")) < MIN_PASSWORD:
        abort(400, f"Password must be at least {MIN_PASSWORD} characters")
    return name, username, str(password)


def _clean_access(raw, allowed_sheets, allowed_roles):
    out = {}
    for s, r in (raw or {}).items():
        if s not in allowed_sheets:
            abort(403, f"You can't give access to '{s}'.")
        if r is None or r == "":
            out[s] = None
        elif r in allowed_roles:
            out[s] = r
        elif r in TOURNAMENT_ROLES:
            abort(403, "Captains can add players. Only an admin can make someone a captain.")
        else:
            abort(400, "Bad access level")
    return out


# --------------------------------------------------------------------------- #
# wiring
# --------------------------------------------------------------------------- #
def init(app, data_dir, static_dir, sheets_fn):
    _cfg.update(users_file=os.path.join(data_dir, "crictrack-users.json"),
                settings_file=os.path.join(data_dir, "crictrack-settings.json"), sheets=sheets_fn)
    st = _read_json(_cfg["settings_file"], {})
    if not st.get("secret_key"):
        st["secret_key"] = secrets.token_hex(32)
        _write_json(_cfg["settings_file"], st)
    app.secret_key = st["secret_key"]
    app.config.update(SESSION_COOKIE_NAME="crictrack", SESSION_COOKIE_HTTPONLY=True,
                      SESSION_COOKIE_SAMESITE="Lax",
                      SESSION_COOKIE_SECURE=os.environ.get("SECURE_COOKIES", "") == "1",
                      PERMANENT_SESSION_LIFETIME=dt.timedelta(days=30))

    @app.before_request
    def _guard():
        p = request.path
        if request.method in ("POST", "PUT", "PATCH", "DELETE") and request.headers.get("X-CricTrack") != "1":
            abort(403, "Request blocked (missing app header).")       # CSRF protection
        g.user = None
        uname = session.get("u")
        if uname:
            u = _find(_users(), uname)
            if u and u["status"] == "active" and session.get("sv") == u.get("sv"):
                g.user = u
            else:
                session.clear()
        if p in ("/login", "/favicon.ico", "/icon.svg", "/apple-touch-icon.png") or p.startswith("/api/auth/"):
            return
        if not g.user:
            if p.startswith("/api/") or p == "/download":
                return jsonify({"ok": False, "error": "Please sign in", "login": True}), 401
            return redirect("/login")

    @app.get("/login")
    def login_page():
        if current_user():
            return redirect("/")
        return send_from_directory(static_dir, "login.html")

    @app.get("/api/auth/state")
    def auth_state():
        d = _users()
        return jsonify({"setup": not d["users"], "pin_set": _pin_is_set(),
                        "allow_register": d.get("allow_register", True),
                        "me": public(current_user()) if current_user() else None})

    @app.post("/api/auth/login")
    def auth_login():
        b = _body()
        username = str(b.get("username") or "").strip().lower()
        key = f"{_ip()}|{username}"
        _throttle(key)
        with _lock:
            d = _users()
            u = _find(d, username)
            if not _check_password(u, b.get("password")):
                _fail(key)
                abort(401, "Wrong username or password")
            if u["status"] == "pending":
                abort(403, "Your account is waiting for approval from your captain or the admin.")
            if u["status"] != "active":
                abort(403, "Your account has been disabled. Please contact your captain.")
            _fails.pop(key, None)
            u["last_login"] = _now()
            _save_users(d)
        session.clear()
        session.permanent = True
        session["u"], session["sv"] = u["username"], u["sv"]
        return jsonify({"ok": True, "me": public(u)})

    @app.post("/api/auth/register")
    def auth_register():
        b = _body()
        _throttle(f"reg|{_ip()}")
        name, username, password = _validate_new(b.get("name"), b.get("username"), b.get("password"))
        with _lock:
            d = _users()
            first = not d["users"]
            if first:
                if not _setup_pin_ok(b.get("setup_pin")):
                    _fail(f"reg|{_ip()}")
                    abort(403, "Enter the current edit PIN to create the admin account.")
            elif not d.get("allow_register", True):
                abort(403, "Sign-ups are closed. Ask your captain to create an account for you.")
            if _find(d, username):
                abort(409, "That username is taken. Try another one.")
            u = {"username": username, "name": name, "role": "admin" if first else "member",
                 "status": "active" if first else "pending", "access": {},
                 "note": str(b.get("note") or "").strip()[:80], "created": _now(), "last_login": None}
            _set_password(u, password)
            d["users"].append(u)
            _save_users(d)
        if first:
            session.clear()
            session.permanent = True
            session["u"], session["sv"] = u["username"], u["sv"]
            return jsonify({"ok": True, "me": public(u), "message": "Admin account created."})
        return jsonify({"ok": True, "pending": True,
                        "message": "Thanks! Your captain will add you to your tournament. You'll be able to sign in once they do."})

    @app.post("/api/auth/logout")
    def auth_logout():
        session.clear()
        return jsonify({"ok": True})

    @app.post("/api/auth/password")
    def auth_password():
        u = current_user()
        if not u:
            abort(401, "Please sign in")
        b = _body()
        if len(str(b.get("new") or "")) < MIN_PASSWORD:
            abort(400, f"New password must be at least {MIN_PASSWORD} characters")
        with _lock:
            d = _users()
            me = _find(d, u["username"])
            if not _check_password(me, b.get("current")):
                abort(400, "Current password is wrong")
            _set_password(me, str(b["new"]))
            _save_users(d)
        session["sv"] = me["sv"]
        return jsonify({"ok": True, "message": "Password changed. Other devices have been signed out."})

    # ---------------- access management: admins (everything) and captains (their tournaments) ----------------
    def _manage_ctx():
        if not can_manage():
            abort(403, "Only an admin or a captain can manage access.")
        all_sheets = list(_cfg["sheets"]())
        return (all_sheets, None) if is_admin() else (captain_sheets(), captain_sheets())

    def _captain_may_touch(target, my_sheets):
        """Captains may manage plain players (not admins, not captains), and only via their own tournaments."""
        if is_admin(target):
            abort(403, "Only an admin can change an admin's access.")
        if any(r == "captain" for r in target.get("access", {}).values()):
            abort(403, "Only an admin can change another captain's access.")

    @app.get("/api/users")
    def users_list():
        sheets, only = _manage_ctx()
        d = _users()
        order = {"pending": 0, "active": 1, "disabled": 2}
        users = sorted(d["users"], key=lambda u: (order.get(u["status"], 3), u["role"] != "admin", u["name"].lower()))
        if not is_admin():
            users = [u for u in users if u["status"] != "disabled"]
        return jsonify({"ok": True, "scope": "admin" if is_admin() else "captain", "sheets": sheets,
                        "users": [public(u, only) for u in users],
                        "allow_register": d.get("allow_register", True), "me": public(current_user())})

    @app.post("/api/users")
    def users_create():
        sheets, only = _manage_ctx()
        b = _body()
        name, username, password = _validate_new(b.get("name"), b.get("username"), b.get("password"))
        if is_admin():
            role = "admin" if b.get("role") == "admin" else "member"
            acc = _clean_access(b.get("access"), sheets, TOURNAMENT_ROLES)
        else:
            role = "member"
            acc = _clean_access(b.get("access"), sheets, ("player",))
            if not any(acc.values()):
                abort(400, "Pick the tournament to add them to.")
        with _lock:
            d = _users()
            if _find(d, username):
                abort(409, "That username is taken")
            u = {"username": username, "name": name, "role": role, "status": "active",
                 "access": {s: r for s, r in acc.items() if r}, "created": _now(), "last_login": None}
            _set_password(u, password)
            d["users"].append(u)
            _save_users(d)
        what = "admin" if role == "admin" else ", ".join(f"{r} of {s}" for s, r in u["access"].items()) or "no tournaments yet"
        return jsonify({"ok": True, "message": f"Account created for {name} ({what})."})

    @app.patch("/api/users/<username>")
    def users_update(username):
        sheets, only = _manage_ctx()
        b = _body()
        with _lock:
            d = _users()
            u = _find(d, username)
            if not u:
                abort(404, "User not found")
            me = current_user()["username"]
            if is_admin():
                role = b.get("role", u["role"])
                status = b.get("status", u["status"])
                if role not in ("admin", "member") or status not in ("active", "pending", "disabled"):
                    abort(400, "Bad role or status")
                if u["username"] == me and (role != "admin" or status != "active"):
                    abort(400, "You can't remove your own admin access. Ask another admin to do it.")
                if is_admin(u) and u["status"] == "active" and (role != "admin" or status != "active") and len(_admins(d)) <= 1:
                    abort(400, "There must always be at least one active admin.")
                acc = _clean_access(b.get("access"), sheets, TOURNAMENT_ROLES)
                if b.get("name"):
                    u["name"] = re.sub(r"\s+", " ", str(b["name"])).strip()[:60]
                if (role, status) != (u["role"], u["status"]) and status != "active":
                    u["sv"] = secrets.token_hex(4)
                u["role"], u["status"] = role, status
            else:
                if any(k in b for k in ("role", "status", "name")):
                    abort(403, "Only an admin can change roles or account status.")
                _captain_may_touch(u, sheets)
                acc = _clean_access(b.get("access"), sheets, ("player",))
            for s, r in acc.items():
                if r:
                    u.setdefault("access", {})[s] = r
                else:
                    u.setdefault("access", {}).pop(s, None)
            if acc and any(acc.values()) and u["status"] == "pending":
                u["status"] = "active"                      # giving access approves a sign-up
            _save_users(d)
        if u["status"] == "pending":
            msg = f"{u['name']} is waiting for approval."
        elif u["status"] == "disabled":
            msg = f"{u['name']} is disabled and can't sign in."
        elif is_admin(u):
            msg = f"{u['name']} is an admin (all tournaments)."
        else:
            vis = {s: r for s, r in u.get("access", {}).items() if only is None or s in only}
            msg = f"{u['name']}: " + (", ".join(f"{r} of {s}" for s, r in vis.items()) or "no access to " +
                                      ("any tournament" if only is None else "your tournaments"))
        return jsonify({"ok": True, "message": msg})

    @app.post("/api/users/<username>/password")
    def users_reset_password(username):
        sheets, only = _manage_ctx()
        b = _body()
        if len(str(b.get("password") or "")) < MIN_PASSWORD:
            abort(400, f"Password must be at least {MIN_PASSWORD} characters")
        with _lock:
            d = _users()
            u = _find(d, username)
            if not u:
                abort(404, "User not found")
            if not is_admin():
                _captain_may_touch(u, sheets)
                if any(s not in sheets for s in u.get("access", {})):
                    abort(403, "This player is also in other tournaments. Ask the admin to reset their password.")
            _set_password(u, str(b["password"]))
            _save_users(d)
        if u["username"] == current_user()["username"]:
            session["sv"] = u["sv"]
        return jsonify({"ok": True, "message": f"Password reset for {u['name']}. Share it with them privately."})

    @app.delete("/api/users/<username>")
    def users_delete(username):
        sheets, only = _manage_ctx()
        with _lock:
            d = _users()
            u = _find(d, username)
            if not u:
                abort(404, "User not found")
            if u["username"] == current_user()["username"]:
                abort(400, "You can't delete your own account.")
            if not is_admin() and not (u["status"] == "pending" and not u.get("access")):
                abort(403, "Captains can only reject sign-up requests. Remove their tournament access instead.")
            if is_admin(u) and u["status"] == "active" and len(_admins(d)) <= 1:
                abort(400, "There must always be at least one active admin.")
            d["users"].remove(u)
            _save_users(d)
        return jsonify({"ok": True, "message": f"Removed {u['name']}'s account."})

    @app.post("/api/users/settings")
    def users_settings():
        require_admin()
        b = _body()
        with _lock:
            d = _users()
            d["allow_register"] = bool(b.get("allow_register"))
            _save_users(d)
        return jsonify({"ok": True, "allow_register": d["allow_register"],
                        "message": "Sign-ups are open (new accounts wait for access)." if d["allow_register"]
                        else "Sign-ups are closed. Only admins and captains can create accounts."})
