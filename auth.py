"""
auth.py
=======
Central authentication and authorisation layer for the RDC-OP multi-module app.

Roles (most to least privileged):
  SUPER_ADMIN     - unrestricted access to every feature in every module
  HO_VIEWER       - read + export across all plants; no write/upload/calculate/settings
  FINANCE_VIEWER  - same as HO_VIEWER
  REGIONAL_USER   - read + export for assigned plants only
  PLANT_USER      - read + export for own plant(s) only
                    additionally: ECMD data entry for own plant only

Plant matching is done by plant_name (case-insensitive).
"""

from __future__ import annotations

import json
import os
from datetime import datetime, timedelta
from functools import wraps
from typing import Optional

import pandas as pd
from flask import (Flask, abort, flash, g, redirect, render_template,
                   request, session, url_for)
from werkzeug.security import check_password_hash, generate_password_hash

# ── Role constants ─────────────────────────────────────────────────────────────
SUPER_ADMIN     = "SUPER_ADMIN"
HO_VIEWER       = "HO_VIEWER"
FINANCE_VIEWER  = "FINANCE_VIEWER"
REGIONAL_USER   = "REGIONAL_USER"
PLANT_USER      = "PLANT_USER"

ALLOWED_ROLES = [SUPER_ADMIN, HO_VIEWER, FINANCE_VIEWER, REGIONAL_USER, PLANT_USER]

# Roles that may write / mutate any data
WRITE_ROLES = [SUPER_ADMIN]

# Roles that can see all plants without a filter
GLOBAL_VIEW_ROLES = [SUPER_ADMIN, HO_VIEWER, FINANCE_VIEWER]

# Roles whose plant access is restricted to user_plant_access rows
RESTRICTED_ROLES = [REGIONAL_USER, PLANT_USER]

# Roles that can enter ECMD readings for their own plant
ECMD_ENTRY_ROLES = [SUPER_ADMIN, PLANT_USER]

SESSION_TIMEOUT_MINUTES = 20


# ── Bootstrap ─────────────────────────────────────────────────────────────────

def bootstrap_admin(app: Flask) -> None:
    """
    Called once at startup. If the users table is empty, create the first
    SUPER_ADMIN from environment variables ADMIN_USERNAME / ADMIN_PASSWORD /
    ADMIN_EMAIL. Prints clear console messages if env vars are missing.
    """
    import database as db

    with app.app_context():
        existing = db.count_users()
        if existing > 0:
            return  # users already exist — nothing to do

        username = os.environ.get("ADMIN_USERNAME", "").strip()
        password = os.environ.get("ADMIN_PASSWORD", "").strip()
        email    = os.environ.get("ADMIN_EMAIL",    "").strip()

        if not username or not password:
            print("=" * 65)
            print("  [!] NO USERS FOUND IN DATABASE")
            print("  Create a .env file with these variables, then restart:")
            print("     ADMIN_USERNAME=your_admin_username")
            print("     ADMIN_PASSWORD=your_secure_password")
            print("     ADMIN_EMAIL=your_email@example.com")
            print("=" * 65)
            return

        pw_hash = generate_password_hash(password)
        db.create_user(
            full_name="System Administrator",
            email=email or f"{username}@rdc.local",
            username=username,
            password_hash=pw_hash,
            role=SUPER_ADMIN,
            is_active=True,
            must_change_password=False,
        )
        print(f"  [OK] Initial SUPER_ADMIN '{username}' created from .env")


# ── Session helpers ───────────────────────────────────────────────────────────

def get_current_user() -> Optional[dict]:
    """
    Return the logged-in user dict (from DB) or None.
    Also enforces the 20-minute idle timeout.
    """
    if "user_id" not in session:
        return None

    # Idle timeout check
    last_active = session.get("last_active")
    if last_active:
        try:
            la = datetime.fromisoformat(last_active)
            if datetime.utcnow() - la > timedelta(minutes=SESSION_TIMEOUT_MINUTES):
                session.clear()
                return None
        except Exception:
            session.clear()
            return None

    session["last_active"] = datetime.utcnow().isoformat()

    import database as db
    user = db.get_user_by_id(session["user_id"])
    if user is None or not user.get("is_active"):
        session.clear()
        return None
    return user


def login_user(user: dict) -> None:
    """Store user id in the Flask session."""
    session.clear()
    session["user_id"]    = user["id"]
    session["last_active"] = datetime.utcnow().isoformat()
    session.permanent = False  # timeout on browser close + idle timeout


def logout_user() -> None:
    """Clear the Flask session."""
    session.clear()


# ── Plant-access helpers ───────────────────────────────────────────────────────

def get_user_allowed_plants(user: dict) -> list[str]:
    """
    Return list of plant_names this user is allowed to see.
    SUPER_ADMIN / HO_VIEWER / FINANCE_VIEWER → [] means "all" (no filter).
    REGIONAL_USER / PLANT_USER → their assigned plant_names.
    """
    if user["role"] in GLOBAL_VIEW_ROLES:
        return []  # empty = no restriction

    import database as db
    rows = db.get_user_plant_access(user["id"])
    return [r["plant_name"] for r in rows]


def user_can_access_plant(user: dict, plant_name: str) -> bool:
    """True if the user is allowed to see the given plant_name."""
    if user["role"] in GLOBAL_VIEW_ROLES:
        return True
    allowed = get_user_allowed_plants(user)
    return plant_name.strip().lower() in [p.lower() for p in allowed]


def apply_plant_filter_df(df: pd.DataFrame, user: dict,
                          plant_name_col: str = "plant_name") -> pd.DataFrame:
    """
    Filter a DataFrame to only rows the user can see.
    Works on any DataFrame that has a plant_name column.
    SUPER_ADMIN / HO / FINANCE → return full df unchanged.
    REGIONAL / PLANT → keep only rows whose plant_name is in allowed list.
    """
    if df is None or df.empty:
        return df
    if user["role"] in GLOBAL_VIEW_ROLES:
        return df
    allowed = get_user_allowed_plants(user)
    if not allowed:
        return df.iloc[0:0]  # empty — no plants assigned
    allowed_lower = [p.lower() for p in allowed]
    if plant_name_col not in df.columns:
        return df
    return df[df[plant_name_col].str.lower().isin(allowed_lower)]


def apply_plant_filter_rows(rows: list, user: dict,
                             plant_name_col: str = "plant_name") -> list:
    """Same as apply_plant_filter_df but for a list of dicts."""
    if user["role"] in GLOBAL_VIEW_ROLES:
        return rows
    allowed = get_user_allowed_plants(user)
    if not allowed:
        return []
    allowed_lower = [p.lower() for p in allowed]
    return [r for r in rows if str(r.get(plant_name_col, "")).lower() in allowed_lower]


def apply_plant_filter_rows_by_code(rows: list, user: dict,
                                    plant_code_col: str = "plant_code") -> list:
    """
    Filter rows by plant_code when plant_name is not available.
    Resolves plant codes → plant names via tp_plant_data.
    Falls back to plant_code direct match if mapping unavailable.
    """
    if user["role"] in GLOBAL_VIEW_ROLES:
        return rows
    allowed_names = get_user_allowed_plants(user)
    if not allowed_names:
        return []

    import database as db
    plant_rows = db.get_tp_plants()
    # Build code→name mapping
    code_to_name = {p["plant_code"].lower(): p["plant_name"].lower()
                    for p in plant_rows}
    allowed_lower = [n.lower() for n in allowed_names]

    def _allowed(row):
        code = str(row.get(plant_code_col, "")).lower()
        name_via_code = code_to_name.get(code, "")
        return name_via_code in allowed_lower or code in allowed_lower

    return [r for r in rows if _allowed(r)]


# ── Audit logging ─────────────────────────────────────────────────────────────

def log_activity(user: Optional[dict], action: str, module_name: str = "",
                 details: dict | None = None) -> None:
    """Write one row to user_activity_log. Safe — never raises."""
    try:
        import database as db
        db.log_user_activity(
            user_id=user["id"] if user else None,
            action=action,
            module_name=module_name,
            details_json=json.dumps(details or {}),
            ip_address=request.remote_addr or "",
        )
    except Exception:
        pass


def log_login(user_id: Optional[int], status: str,
              failure_reason: str = "") -> None:
    """Write one row to login_audit_log."""
    try:
        import database as db
        db.log_login_attempt(
            user_id=user_id,
            ip_address=request.remote_addr or "",
            user_agent=request.headers.get("User-Agent", "")[:255],
            status=status,
            failure_reason=failure_reason,
        )
    except Exception:
        pass


# ── Decorators ────────────────────────────────────────────────────────────────

def login_required(f):
    """Redirect to /login if no valid session."""
    @wraps(f)
    def decorated(*args, **kwargs):
        user = get_current_user()
        if user is None:
            flash("Please log in to continue.", "warning")
            return redirect(url_for("page_login", next=request.path))
        g.current_user = user
        return f(*args, **kwargs)
    return decorated


def role_required(*allowed_roles):
    """
    Decorator: user must be logged in AND have one of the allowed roles.
    Usage:
        @role_required(SUPER_ADMIN)
        @role_required(SUPER_ADMIN, HO_VIEWER)
    """
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            user = get_current_user()
            if user is None:
                flash("Please log in to continue.", "warning")
                return redirect(url_for("page_login", next=request.path))
            if user["role"] not in allowed_roles:
                log_activity(user, "ACCESS_DENIED",
                             details={"path": request.path, "method": request.method})
                return render_template("access_denied.html",
                                       current_user=user,
                                       required_roles=list(allowed_roles)), 403
            g.current_user = user
            return f(*args, **kwargs)
        return decorated
    return decorator


def admin_required(f):
    """Shorthand: only SUPER_ADMIN."""
    return role_required(SUPER_ADMIN)(f)


def viewer_or_above(f):
    """Any logged-in user — but must be logged in."""
    return login_required(f)


def can_write(f):
    """Only roles that may mutate data (SUPER_ADMIN)."""
    return role_required(SUPER_ADMIN)(f)


# ── Permission check helpers ───────────────────────────────────────────────────

def check_page_permission(user: dict, page_name: str) -> bool:
    """
    Return True if the user's role allows the named page.
    page_name is a string like 'settings', 'data_uploader', 'calculate', etc.
    """
    role = user["role"]
    write_pages = {
        "settings", "data_uploader", "data_entry_write",
        "calculate", "sync", "oracle_fetch", "user_management",
        "audit_log", "validation_download",
    }
    if role == SUPER_ADMIN:
        return True
    if page_name in write_pages:
        return False
    return True  # read pages open to all authenticated roles


def can_do_ecmd_entry(user: dict, plant_name: str) -> bool:
    """True if this user can save/view ECMD readings for the given plant."""
    if user["role"] == SUPER_ADMIN:
        return True
    if user["role"] == PLANT_USER:
        return user_can_access_plant(user, plant_name)
    return False


# ── Login / Logout views (registered in app.py) ───────────────────────────────

def login_view():
    """Handle GET /login and POST /login."""
    import database as db

    if get_current_user() is not None:
        next_url = request.args.get("next", "")
        if next_url and next_url.startswith("/") and not next_url.startswith("//"):
            return redirect(next_url)
        return redirect(url_for("page_home"))

    if request.method == "GET":
        return render_template("login.html",
                               next=request.args.get("next", ""))

    email = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    next_url = request.form.get("next", "").strip() or url_for("page_home")

    if not email or not password:
        flash("Email and password are required.", "error")
        return render_template("login.html", next=next_url), 400

    user = db.get_user_by_email(email)

    if user is None:
        log_login(None, "FAILED", f"Unknown email: {email}")
        flash("Invalid email or password.", "error")
        return render_template("login.html", next=next_url), 401

    if not user.get("is_active"):
        log_login(user["id"], "FAILED", "Account inactive")
        flash("Your account is inactive. Contact the administrator.", "error")
        return render_template("login.html", next=next_url), 401

    if not check_password_hash(user["password_hash"], password):
        log_login(user["id"], "FAILED", "Wrong password")
        flash("Invalid username or password.", "error")
        return render_template("login.html", next=next_url), 401

    # Success
    login_user(user)
    db.update_last_login(user["id"])
    log_login(user["id"], "SUCCESS")
    log_activity(user, "LOGIN", details={"ip": request.remote_addr})

    # Redirect to change-password page if flagged
    if user.get("must_change_password"):
        flash("Please change your password before continuing.", "warning")
        return redirect(url_for("page_change_password"))

    # Safe redirect — only allow relative paths
    if next_url and next_url.startswith("/") and not next_url.startswith("//"):
        return redirect(next_url)
    return redirect(url_for("page_home"))


def logout_view():
    """Handle POST /logout."""
    user = get_current_user()
    if user:
        log_activity(user, "LOGOUT")
    logout_user()
    flash("You have been logged out.", "info")
    return redirect(url_for("page_login"))


def change_password_view():
    """Handle GET/POST /change-password."""
    import database as db

    user = get_current_user()
    if user is None:
        return redirect(url_for("page_login"))

    if request.method == "GET":
        return render_template("change_password.html", current_user=user)

    current_pw  = request.form.get("current_password", "")
    new_pw      = request.form.get("new_password", "")
    confirm_pw  = request.form.get("confirm_password", "")

    if not check_password_hash(user["password_hash"], current_pw):
        flash("Current password is incorrect.", "error")
        return render_template("change_password.html", current_user=user), 400

    if len(new_pw) < 8:
        flash("New password must be at least 8 characters.", "error")
        return render_template("change_password.html", current_user=user), 400

    if new_pw != confirm_pw:
        flash("New passwords do not match.", "error")
        return render_template("change_password.html", current_user=user), 400

    db.update_user_password(user["id"], generate_password_hash(new_pw),
                            must_change_password=False)
    log_activity(user, "PASSWORD_CHANGED")
    flash("Password changed successfully.", "success")
    return redirect(url_for("page_home"))


# ── Context processor (inject current_user into every template) ───────────────

def inject_auth_context():
    """Registered as a context_processor in app.py."""
    user = get_current_user()
    allowed_plants = get_user_allowed_plants(user) if user else []
    return dict(
        current_user=user,
        allowed_plants=allowed_plants,
        SUPER_ADMIN=SUPER_ADMIN,
        HO_VIEWER=HO_VIEWER,
        FINANCE_VIEWER=FINANCE_VIEWER,
        REGIONAL_USER=REGIONAL_USER,
        PLANT_USER=PLANT_USER,
        GLOBAL_VIEW_ROLES=GLOBAL_VIEW_ROLES,
        RESTRICTED_ROLES=RESTRICTED_ROLES,
        WRITE_ROLES=WRITE_ROLES,
    )
