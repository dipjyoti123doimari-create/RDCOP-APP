"""
cache_helpers.py
================
Phase 13 — keep the app fast and tidy.

This file provides two simple, SAFE maintenance jobs that the Settings page
uses behind the "Clear Cache" button:

  1. clear_streamlit_cache()
     Streamlit can keep results in memory to make pages load faster. If those
     in-memory results ever go stale, clearing them forces the app to read
     fresh data from the database on the next click. This NEVER deletes any of
     your saved data — it only empties Streamlit's temporary memory.

  2. compact_database()
     Over time, after many uploads and re-calculations, the SQLite database
     file (data/app.db) can hold "empty space" left behind by deleted rows.
     SQLite's VACUUM command rewrites the file without that empty space, so the
     file gets smaller and reads stay quick. Your data is fully preserved.

The main helper the button calls is clear_cache(), which runs BOTH jobs and
returns a short summary (so the page can tell you exactly what happened,
including how many megabytes were freed).

There is also a tiny file_hash() helper: it makes a short "fingerprint" of an
uploaded file's contents. This can be used later to notice when an uploaded
file has not actually changed, so we can skip redoing work.

IMPORTANT: Nothing in this file deletes Master Data, Backend Data, calculation
results or settings. It is safe to run any time.
"""

import hashlib
import os

import database


def file_hash(file_bytes: bytes) -> str:
    """
    Return a short, stable "fingerprint" (SHA-256, first 16 characters) for the
    given file contents. The same bytes always give the same fingerprint, so we
    can tell when a re-uploaded file is actually identical to the last one.
    """
    if not file_bytes:
        return ""
    return hashlib.sha256(file_bytes).hexdigest()[:16]


def get_db_size_mb() -> float:
    """Return the size of the database file in megabytes (0.0 if it is missing)."""
    if not os.path.exists(database.DB_PATH):
        return 0.0
    size_bytes = os.path.getsize(database.DB_PATH)
    return round(size_bytes / (1024 * 1024), 2)


def clear_streamlit_cache() -> bool:
    """
    Empty Streamlit's in-memory caches (both kinds). Returns True if it ran.

    We import streamlit INSIDE the function so this file can still be used in
    plain scripts/tests that do not run inside Streamlit.
    """
    try:
        import streamlit as st
        # cache_data = cached values (dataframes, numbers, etc.)
        st.cache_data.clear()
        # cache_resource = cached connections/objects
        st.cache_resource.clear()
        return True
    except Exception:
        # If Streamlit caching is unavailable for any reason, fail quietly —
        # there is simply nothing in memory to clear.
        return False


def compact_database() -> dict:
    """
    Run SQLite's VACUUM to shrink the database file by removing unused space.

    Returns a small summary:
        {"before_mb": float, "after_mb": float, "freed_mb": float, "error": str|None}

    NOTE: VACUUM cannot run inside an open transaction, so we set
    isolation_level = None (autocommit) on a fresh connection first.
    """
    before = get_db_size_mb()
    result = {"before_mb": before, "after_mb": before, "freed_mb": 0.0, "error": None}
    try:
        conn = database.get_connection()
        try:
            conn.isolation_level = None      # autocommit, required for VACUUM
            conn.execute("VACUUM")
        finally:
            conn.close()
        after = get_db_size_mb()
        result["after_mb"] = after
        result["freed_mb"] = round(max(0.0, before - after), 2)
    except Exception as exc:  # noqa: BLE001 - report any failure to the page
        result["error"] = str(exc)
    return result


def clear_cache() -> dict:
    """
    The single helper the Settings "Clear Cache" button calls.

    It does the two safe maintenance jobs and returns a summary the page can
    show to the user:

        {
            "streamlit_cache_cleared": bool,
            "before_mb": float,
            "after_mb":  float,
            "freed_mb":  float,
            "error":     str | None,   # set only if compacting failed
        }
    """
    cleared = clear_streamlit_cache()
    vac = compact_database()
    return {
        "streamlit_cache_cleared": cleared,
        "before_mb": vac["before_mb"],
        "after_mb":  vac["after_mb"],
        "freed_mb":  vac["freed_mb"],
        "error":     vac["error"],
    }
