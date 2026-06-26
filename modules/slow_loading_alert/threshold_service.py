"""
threshold_service.py
====================
Calculates the allowed loading time for a given mixer capacity and batched quantity.

Reference table (base time for 6 m³):
  30→16 min, 45→12, 56→10, 60→10, 65→10, 70→10, 75→10,
  80→10, 86→10, 90→8, 100→8, 110→8, 120→8

Dynamic rule:
  Allowed Time = base_time × (batched_quantity / 6)  → rounded to nearest minute

Nearest-capacity fallback:
  If exact capacity not found, use nearest lower; if none, nearest higher.
  Add remark to explain.
"""

import math
import database


def _load_thresholds() -> list:
    """Load active thresholds sorted by mixer_capacity ascending."""
    return database.sla_get_thresholds()


def get_allowed_time(mixer_capacity, batched_quantity) -> tuple[float, str]:
    """
    Return (allowed_minutes: float, remark: str).

    allowed_minutes is rounded to the nearest whole minute.
    remark is empty string when an exact match is found.
    """
    remarks = []

    # Coerce inputs
    try:
        cap = float(mixer_capacity) if mixer_capacity is not None else None
    except (TypeError, ValueError):
        cap = None

    try:
        qty = float(batched_quantity) if batched_quantity is not None else 6.0
        if qty <= 0:
            qty = 6.0
    except (TypeError, ValueError):
        qty = 6.0

    thresholds = _load_thresholds()
    if not thresholds:
        # Fallback: hardcoded defaults if DB is empty
        thresholds = [
            {"mixer_capacity": c, "reference_quantity": 6.0, "base_allowed_minutes": b}
            for c, b in [
                (30,16),(45,12),(56,10),(60,10),(65,10),(70,10),(75,10),
                (80,10),(86,10),(90,8),(100,8),(110,8),(120,8)
            ]
        ]

    if cap is None:
        base = 10.0
        remarks.append("Mixer capacity missing, default threshold used.")
    else:
        capacities = [t["mixer_capacity"] for t in thresholds]
        # Exact match
        exact = [t for t in thresholds if t["mixer_capacity"] == cap]
        if exact:
            base = float(exact[0]["base_allowed_minutes"])
            ref_qty = float(exact[0]["reference_quantity"])
        else:
            # Nearest lower
            lower = [t for t in thresholds if t["mixer_capacity"] < cap]
            if lower:
                best = max(lower, key=lambda t: t["mixer_capacity"])
                base = float(best["base_allowed_minutes"])
                ref_qty = float(best["reference_quantity"])
                remarks.append(
                    f"Threshold derived using nearest mixer capacity "
                    f"({best['mixer_capacity']} m³ used for {cap} m³)."
                )
            else:
                # Nearest higher
                higher = [t for t in thresholds if t["mixer_capacity"] > cap]
                best = min(higher, key=lambda t: t["mixer_capacity"])
                base = float(best["base_allowed_minutes"])
                ref_qty = float(best["reference_quantity"])
                remarks.append(
                    f"Threshold derived using nearest mixer capacity "
                    f"({best['mixer_capacity']} m³ used for {cap} m³)."
                )

    # Scale by batched quantity relative to reference quantity
    ref_qty_val = 6.0
    if cap is not None:
        exact = [t for t in thresholds if t["mixer_capacity"] == cap]
        if exact:
            ref_qty_val = float(exact[0]["reference_quantity"])

    allowed = base * (qty / ref_qty_val)
    allowed = round(allowed)  # nearest whole minute

    return float(allowed), "; ".join(remarks)


def classify_severity(delay_minutes: float) -> str:
    """Return 'RED', 'AMBER', or '' based on delay."""
    if delay_minutes > 3:
        return "RED"
    if delay_minutes >= 1:
        return "AMBER"
    return ""
