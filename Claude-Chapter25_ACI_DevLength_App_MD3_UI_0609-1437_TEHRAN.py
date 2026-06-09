"""
Chapter25 ACI Development Length App - Professional Clean Build
Based on: Chapter25-V283_LENGTHS_STD_TABLE_WITH_DIFFS_0218-1312_TEHRAN.py
Purpose: Streamlit app for ACI 318-2019 / 2025 rebar development lengths

This file is a clean distributable version prepared for maintenance and deployment.
"""

# v83: Fix Excel export robustness + stable download + persist append-diff checkbox state (1227-1240_TEHRAN)
# acirebar_app_v61_UIpro.py
# Streamlit app for ACI 318-2019 / 2025 rebar development lengths
# v61: Same technical core as v60, improved UI/UX only.
# MIT License

import math
import io
import re
from typing import List, Dict, Any, Optional, Tuple

import streamlit as st
import streamlit.components.v1 as components

# ---------------------------------------------------------------------------
# CONTROLLED ROADMAP IMPLEMENTATION
# Phase 0/1: Baseline lock + safe runtime defaults
# Generated: 0518-1636_TEHRAN
# Notes:
# - Calculation formulas are not modified in this phase.
# - UI shell upgrade will be added incrementally after this stable baseline.
# ---------------------------------------------------------------------------
ENGINE_VERSION = "FIX14_GEOMETRY_PANEL_COMPACT_PRO_LAYOUT_0525-1905_TEHRAN"
UI_ROADMAP_PHASE = "Phase 8 - Compact Professional Geometry Panel Layout"


# ======================
# Release / Version Info
# ======================
from pathlib import Path
import sys
from datetime import datetime, timedelta, timezone


def _fmt_used_params_kv(pairs):
    """Format ACI parameters as a compact, comma-separated string."""
    out = []
    for sym, val, unit in pairs:
        if val is None:
            continue
        try:
            sval = f"{float(val):.2f}".rstrip("0").rstrip(".")
        except Exception:
            sval = str(val)
        out.append(f"{sym} = {sval}{(' ' + unit) if unit else ''}".rstrip())
    return ", ".join(out)

APP_NAME = "ACI 318 Chapter 25 – Rebar Development & Hook Length Calculator"
APP_VERSION = "0.9.9"   # Removed Apply&Recalc, unified Calculate/Update

_TEHRAN_OFFSET = timedelta(hours=3, minutes=30)

def now_tehran() -> datetime:
    """Return current datetime in Tehran time (UTC+3:30, no DST)."""
    return datetime.now(timezone.utc) + _TEHRAN_OFFSET

BUILD_ID = now_tehran().strftime("%m%d-%H%M_TEHRAN")
BUILD_TS = now_tehran().strftime("%Y-%m-%d %H:%M Tehran")

def resource_path(relative_path: str) -> str:
    """Resolve resource paths for both normal execution and PyInstaller bundles."""
    base_path = getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)
    return str(Path(base_path) / relative_path)

OUTPUT_DIR = Path.cwd() / "output"

def make_artifact_filename(kind: str, ext: str, inputs: Dict[str, Any], run_id: str) -> str:
    """Generate traceable artifact filenames for downloads and saved outputs."""
    edition = str(inputs.get("edition", "ACI")).replace(" ", "").replace("/", "-")
    units = str(inputs.get("units", "SI")).upper()
    mode = str(inputs.get("user_mode", "")).strip()
    if mode.startswith("🎓"):
        mode_tag = "EDU"
    elif mode.startswith("🧠"):
        mode_tag = "INT"
    else:
        mode_tag = "PRO"
    return f"ACI_Ch25_{kind}_{mode_tag}_{edition}_{units}_v{APP_VERSION}_{run_id}.{ext}"




# ================================
# Input Schema & Normalization
# ================================
# This schema guarantees that calc_inputs contains every key used by reporting/export
# so that Report/Excel cannot crash with KeyError/UnboundLocalError.
# (UI may conditionally show/hide fields, but snapshot must remain complete.)
INPUT_SCHEMA: Dict[str, Dict[str, Any]] = {
    # Core identifiers
    "edition": {"default": "ACI318-2019", "required": False},
    "units": {"default": "SI", "required": True},  # "SI" or "US"

    # Material properties (display units)
    "fc_d": {"default": None, "required": True},
    "fy_d": {"default": None, "required": True},

    # Options / flags
    "epoxy": {"default": False, "required": False},
    "special_seismic_region": {"default": False, "required": False},

    # Geometry (display units; may be None if not applicable)
    "cover_d": {"default": None, "required": False},
    "spacing_d": {"default": None, "required": False},
    "so_d": {"default": None, "required": False},

    # Geometry (internal SI units; kept for report/exports)
    "cover_mm": {"default": None, "required": False},
    "spacing_mm": {"default": None, "required": False},
    "so_mm": {"default": None, "required": False},
    "beam_b_mm": {"default": 400.0, "required": False},
    "beam_h_mm": {"default": 600.0, "required": False},
    "stirrup_db_mm": {"default": 8.0, "required": False},

    # Material properties (internal SI units)
    "fc_mpa": {"default": None, "required": False},
    "fy_mpa": {"default": None, "required": False},

    # Bar data (internal SI units)
    "db_mm": {"default": None, "required": False},
    "top_bar": {"default": False, "required": False},
    "psi_t_T": {"default": 1.0, "required": False},

    # Confinement input (optional; for Ktr derivation)
    "Atr_mm2": {"default": 0.0, "required": False},
    "s_tr_mm": {"default": 0.0, "required": False},
    "n_long": {"default": 1, "required": False},

    # Hook confinement (input hygiene)
    "s_hook_d": {"default": 0.0, "required": False},
    "s_hook_mm": {"default": 0.0, "required": False},
    "hook_conf_mode": {"default": "Manual", "required": False},  # "Manual" | "Auto (Needs verification)"
    "hook_conf_manual": {"default": False, "required": False},
    "psi_r_used": {"default": 1.0, "required": False},
    "psi_r_source": {"default": "conservative_default", "required": False},

    # Derived / intermediate values used in reports
    "Ktr_base": {"default": 0.0, "required": False},
    "lam": {"default": 1.0, "required": False},
    "class_val": {"default": "Auto", "required": False},

    "cb_mm": {"default": None, "required": False},
    "cb_gov": {"default": None, "required": False},

    "alpha_used": {"default": None, "required": False},
    "beta_raw": {"default": None, "required": False},
    "beta_eff": {"default": None, "required": False},
    "beta_is_capped": {"default": None, "required": False},

    # Development length method / Both outputs
    "ld_method": {"default": "General method (ACI Eq. 25.4.2.4a)", "required": False},
    "ld_detailed": {"default": None, "required": False},
    "ld_simplified": {"default": None, "required": False},
    "ld_final": {"default": None, "required": False},

    # Headed-bar factors (used by some report branches; safe defaults)
    "psi_p_h": {"default": 1.0, "required": False},
    "psi_o_h": {"default": 1.0, "required": False},
    "psi_c_h": {"default": 1.0, "required": False},

    # Hook / confinement keys referenced by report (safe defaults)
    "hook_angle": {"default": 90, "required": False},
    "has_conf": {"default": False, "required": False},
    "cover_ok": {"default": False, "required": False},

    # Comparison flags (report/UI)
    "compare_mode": {"default": False, "required": False},
    "append_diff": {"default": False, "required": False},
    "show_diff_only": {"default": False, "required": False},

    # Warning aggregation
    "warnings": {"default": [], "required": False},
}


def normalize_calc_inputs(raw: Dict[str, Any], *, enforce_required: bool = False) -> Dict[str, Any]:
    """Return a normalized calc_inputs dict with a complete key-set.

    - Fills missing keys with schema defaults.
    - If enforce_required=True, required keys must be present and non-None.
    - Keeps user-provided values as-is unless missing.
    """
    raw = dict(raw or {})
    out: Dict[str, Any] = {}
    for k, meta in INPUT_SCHEMA.items():
        if k in raw and raw[k] is not None:
            out[k] = raw[k]
        else:
            # Use a fresh list for list defaults to avoid shared state
            default = meta.get("default")
            out[k] = list(default) if isinstance(default, list) else default

    # Bring forward any extra keys the app stores (do not drop them)
    for k, v in raw.items():
        if k not in out:
            out[k] = v

    if enforce_required:
        missing = []
        for k, meta in INPUT_SCHEMA.items():
            if meta.get("required") and out.get(k) is None:
                missing.append(k)
        if missing:
            raise KeyError(f"Missing required calc_inputs keys: {missing}")

    return out



def derive_inputs(ci: Dict[str, Any]) -> Dict[str, Any]:
    """Derive secondary geometric/confinement terms from normalized snapshot.

    Conservative by default. Does not alter ACI equations; only organizes inputs.
    """
    ci = dict(ci or {})

    # --- c_b
    try:
        if ci.get("cb_mm") is None:
            cover = ci.get("cover_mm")
            spacing = ci.get("spacing_mm")
            so = ci.get("so_mm")
            db = ci.get("db_mm")
            if all(v is not None for v in (cover, spacing, db)) and float(db) > 0.0:
                cb_val, cb_gov = cb_from_geom(float(cover), float(spacing), float(so or 0.0), float(db))
                ci["cb_mm"] = cb_val
                ci["cb_gov"] = cb_gov
    except Exception:
        # keep existing values
        pass

    # --- Ktr (only if not already computed)
    try:
        if ci.get("Ktr_base") in (None, 0.0):
            Atr = float(ci.get("Atr_mm2") or 0.0)
            s_tr = float(ci.get("s_tr_mm") or 0.0)
            n_long = int(ci.get("n_long") or 0)
            if Atr > 0.0 and s_tr > 0.0 and n_long > 0:
                ci["Ktr_base"] = ktr_aci(Atr, s_tr, n_long)
    except Exception:
        pass

    # --- beta
    try:
        db = float(ci.get("db_mm") or 0.0)
        cb = ci.get("cb_mm")
        Ktr = float(ci.get("Ktr_base") or 0.0)
        if db > 0.0 and cb is not None:
            beta_raw = (float(cb) + Ktr) / db
            beta_eff = min(beta_raw, 2.5)
            ci["beta_raw"] = beta_raw
            ci["beta_eff"] = beta_eff
            ci["beta_is_capped"] = beta_raw > 2.5
    except Exception:
        pass

    # --- cover_ok (keep conservative unless explicitly set elsewhere)
    if ci.get("cover_ok") is None:
        ci["cover_ok"] = False


    # --- ψg (reinforcement grade factor) per ACI Table 25.4.2.5 (SI only)
    try:
        if str(ci.get("units") or "SI").upper() == "SI":
            # Prefer already-normalized fy_mpa; fall back to display fy_d (SI display is MPa here)
            fy_mpa = ci.get("fy_mpa")
            if fy_mpa is None and ci.get("fy_d") is not None:
                fy_mpa = float(ci.get("fy_d"))
            if fy_mpa is not None:
                psi_g, grade_lbl, note = psi_g_from_fy_si(float(fy_mpa), str(ci.get("edition") or "ACI318-2019"))
                ci["psi_g_val"] = psi_g
                ci["psi_g_grade"] = grade_lbl
                ci["psi_g_source"] = note
    except Exception:
        # keep existing values
        pass

    # --- Hook confinement → psi_r bookkeeping (Needs verification for auto rule)
    try:
        db = float(ci.get("db_mm") or 0.0)
        s_hook = float(ci.get("s_hook_mm") or 0.0)
        mode = str(ci.get("hook_conf_mode") or "Manual")
        manual = bool(ci.get("hook_conf_manual"))

        if mode.startswith("Auto") and db > 0.0 and s_hook > 0.0:
            if s_hook <= 3.0 * db:
                ci["has_conf"] = True
                ci["psi_r_used"] = 0.8
                ci["psi_r_source"] = "auto_user_rule_s_hook_le_3db_NEEDS_VERIFICATION"
            else:
                ci["has_conf"] = False
                ci["psi_r_used"] = 1.0
                ci["psi_r_source"] = "auto_user_rule_not_met"
        else:
            ci["has_conf"] = manual
            ci["psi_r_used"] = 0.8 if manual else 1.0
            ci["psi_r_source"] = "manual" if manual else "conservative_default"
    except Exception:
        pass

    return ci

import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime as _dt

# ---------------------------------------------------------------------------
# FIX12E — Arrow-safe display tables
# UI/reporting layer only. No ACI calculation value is recalculated or changed.
# Streamlit/PyArrow can warn when one display column contains both numbers
# and text such as "Spacing/2"; this helper casts mixed object columns to str.
# ---------------------------------------------------------------------------

def _arrow_safe_df(df: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
    """Return a dataframe safe for st.dataframe Arrow serialization."""
    if df is None:
        return None
    try:
        out = df.copy()
        for col in out.columns:
            try:
                if out[col].dtype == "object":
                    out[col] = out[col].map(lambda v: "" if v is None else str(v))
            except Exception:
                out[col] = out[col].astype(str)
        return out
    except Exception:
        try:
            return pd.DataFrame(df).astype(str)
        except Exception:
            return df



# ======================================================================================
# V68 ADDITION (SAFE): ACI 318 Chapter 25 Data Tables (2019 & 2025)
    # --------------------------------------------------------------------------------------
# Data dictionaries for Chapter 25 modification-factor tables used in development length
# calculations (straight bars, hooks, headed bars, compression).
#
# Hooked bars (ACI 318M-25 Table 25.4.3.2) — per user screenshots:
#   ψs: No.29 and smaller=1.0; No.32 & No.36=1.15; No.43=1.30; No.57=1.50
#   ψcc: meets cover=0.7 else 1.0 ; ψr: meets confinement=0.8 else 1.0
# ======================================================================================

from typing import Dict, Any

ACI_318_2019_TABLES: Dict[str, Dict[str, Any]] = {
    "25.4.2.5": {
        "lambda": {"normal": 1.0, "lightweight": 0.75},
        "psi_g": {"grade_280": 1.0, "grade_420": 1.0, "grade_550": 1.15, "grade_690": 1.30},
        "psi_e": {"uncoated_or_zinc": 1.0, "epoxy_other": 1.2, "epoxy_low_cover_or_spacing": 1.5},
        "psi_s": {"no22_and_larger": 1.0, "no19_and_smaller_or_wire": 0.8},
        "psi_t": {"top_cast": 1.3, "other": 1.0},
        "psi_te_cap": {"value": 1.7},
    },
    "25.4.3.2": {
        "lambda": {"normal": 1.0, "lightweight": 0.75},
        "psi_e": {"uncoated_or_zinc": 1.0, "epoxy_or_dual": 1.2},
        "psi_r": {"meets_confinement": 1.0, "other": 1.6},
        "psi_o": {"meets_location": 1.0, "other": 1.25},
        "psi_c_rule": {"fc_lt_42": "psi_c = fc/105 + 0.6", "fc_ge_42": "psi_c = 1.0"},
    },
    "25.4.4.3": {
        "psi_e": {"uncoated_or_zinc": 1.0, "epoxy_or_dual": 1.2},
        "psi_p_rule": {"meets_parallel_ties": 1.0, "other": "psi_p = min(2 - s/(8*db), 1.6)"},
        "psi_o": {"meets_location": 1.0, "other": 1.25},
        "psi_c_rule": {"fc_lt_42": "psi_c = fc/105 + 0.6", "fc_ge_42": "psi_c = 1.0"},
    },
    "25.4.9.3": {
        "lambda": {"normal": 1.0, "lightweight": 0.75},
        "psi_r": {"spiral": 0.75, "other": 1.0},
    },
}

ACI_318_2025_TABLES: Dict[str, Dict[str, Any]] = {
    "25.4.2.5": {
        "lambda": {"normal": 1.0, "lightweight": 0.75},
        "psi_g": {"grade_280": 1.0, "grade_420": 1.0, "grade_550": 1.15, "grade_690": 1.30},
        "psi_e": {"uncoated_or_zinc": 1.0, "epoxy_other": 1.2, "epoxy_low_cover_or_spacing": 1.5},
        "psi_s": {"no22_and_larger": 1.0, "no19_and_smaller_or_wire": 0.8},
        "psi_t": {"top_cast": 1.3, "other": 1.0},
        "psi_te_cap": {"value": 1.7},
    },
    "25.4.3.2": {
        "lambda": {"normal": 1.0, "lightweight": 0.75},
        "psi_e": {"uncoated_or_zinc": 1.0, "epoxy_or_dual": 1.2},
        "psi_s": {"no29_and_smaller": 1.0, "no32_no36": 1.15, "no43": 1.30, "no57": 1.50},
        "psi_cc": {"meets_cover": 0.7, "other": 1.0},
        "psi_r": {"meets_tie_confinement": 0.8, "other": 1.0},
    },
    "25.4.4.3": {
        "psi_e": {"uncoated_or_zinc": 1.0, "epoxy_or_dual": 1.2},
        "psi_p_rule": {"meets_parallel_ties": 1.0, "other": "psi_p = min(2 - s/(8*db), 1.6)"},
        "psi_o": {"meets_location": 1.0, "other": 1.25},
        "psi_c_rule": {"fc_lt_42": "psi_c = fc/105 + 0.6", "fc_ge_42": "psi_c = 1.0"},
    },
    "25.4.9.3": {
        "lambda": {"normal": 1.0, "lightweight": 0.75},
        "psi_r": {"spiral": 0.75, "other": 1.0},
    },
}



def psi_g_from_fy_si(fy_mpa: float, edition: str) -> tuple[float, str, str]:
    """Return (ψg, grade_label, note) per ACI Table 25.4.2.5 (SI).

    Policy (user-selected for Iran practice):
    - If fy matches a listed nominal Grade (280/420/550/690 MPa) within tolerance -> use that row (exact).
    - Otherwise -> use the *nearest* listed Grade row (no interpolation).
      If two grades are equally near, the higher grade is selected (conservative for ψg).

    Notes are designed for display in the detailed report, to clearly flag "nearest grade" usage.
    """
    try:
        fy = float(fy_mpa)
    except Exception:
        return 1.0, "unknown", "ψg defaulted (fy not numeric)"

    tables = ACI_318_2025_TABLES if str(edition).startswith("ACI318-2025") else ACI_318_2019_TABLES
    tab = (tables.get("25.4.2.5") or {}).get("psi_g") or {}

    grade_map = {280.0: "grade_280", 420.0: "grade_420", 550.0: "grade_550", 690.0: "grade_690"}

    # 1) Exact match to nominal Grade (within tolerance)
    tol = 1.0  # MPa tolerance for matching nominal Grade
    for g_nom, key in grade_map.items():
        if abs(fy - g_nom) <= tol and key in tab:
            return float(tab[key]), f"Grade {int(g_nom)}", f"ACI Table 25.4.2.5 (exact Grade {int(g_nom)})"

    # 2) Nearest nominal Grade (no interpolation)
    available = [(g_nom, key) for g_nom, key in grade_map.items() if key in tab]
    if not available:
        return 1.0, "unknown", "ψg defaulted (Table 25.4.2.5 not found)"

    # Nearest by absolute difference; if tie, pick higher grade
    available.sort(key=lambda t: (abs(fy - t[0]), -t[0]))
    g_near, key_near = available[0]
    psi = float(tab[key_near])
    return psi, f"Grade {int(g_near)}", f"ACI Table 25.4.2.5 (nearest Grade {int(g_near)})"


def show_used_values_summary(title: str, rows: list[tuple[str, object]]) -> None:
    """
    UI-only helper: Show a compact 'used values' summary in the main (non-sidebar) area.
    Display is gated to AFTER calculation only.
    """
    if not st.session_state.get("data_calculated", False):
        return

    st.markdown(f'<div class="aci-section-title">✅ Used values summary — {title}</div>', unsafe_allow_html=True)
    with st.container(border=True):
        if not rows:
            st.write("—")
            return
        left, right = st.columns(2)
        half = (len(rows) + 1) // 2
        for i, (k, v) in enumerate(rows):
            tgt = left if i < half else right
            with tgt:
                st.caption(str(k))
                st.write("—" if v is None else v)

def get_ch25_tables_for_edition(edition: str) -> Dict[str, Dict[str, Any]]:
    return ACI_318_2025_TABLES if str(edition).startswith("ACI318-2025") else ACI_318_2019_TABLES

def hook_psi_s_group_2025_by_db_mm(db_mm: float) -> str:
    if db_mm <= 29.0:
        return "no29_and_smaller"
    if db_mm <= 36.0:
        return "no32_no36"
    if db_mm <= 43.0:
        return "no43"
    return "no57"


# ==========================================
# 1. CONSTANTS & MATH UTILS (UNCHANGED CORE)
# ==========================================
MPA_PER_KGFCM2 = 0.0980665
SQRT_FC_CAP_MPA = 8.3  # cap on sqrt(fc') for numerical stability

def _compare_diff_rows(label: str, v19, v25):
    """Return (label, 2019, 2025) if different, else None."""
    if str(v19) == str(v25):
        return None
    return (label, v19, v25)



def _compare_row_all(label: str, v19, v25):
    """Return (label, 2019, 2025) always (even if equal)."""
    return (label, "" if v19 is None else str(v19), "" if v25 is None else str(v25))
def _compare_row(label: str, v19, v25):
    """Always return a comparison row (used for side-by-side mode)."""
    return (label, v19, v25)

def render_diff_only_table(title: str, rows, *, silent_if_empty: bool = True) -> bool:
    """Render a compact table showing ONLY rows with differences.

    Returns True if something was rendered (i.e., at least one differing row exists).
    """
    rows = [r for r in rows if r is not None]
    if not rows:
        if not silent_if_empty:
            st.success("No differences for the selected inputs (this item governs the same in both editions).")
        return False
    st.markdown(f"**{title} — Differences only (2019 vs 2025):**")
    df = pd.DataFrame(rows, columns=["Item", "ACI 318-2019", "ACI 318-2025"])
    st.dataframe(df.astype(str), width='stretch', hide_index=True)
    return True


def render_compare_table(title: str, rows):
    """Render a side-by-side table (including equal items)."""
    rows = [r for r in rows if r is not None]
    st.markdown(f"**{title} — Side-by-side (2019 vs 2025):**")
    if not rows:
        st.info('No items to display.')
        return
    df = pd.DataFrame(rows, columns=['Item', 'ACI 318-2019', 'ACI 318-2025'])
    st.dataframe(df.astype(str), width='stretch', hide_index=True)



LD_MIN_MM = 300.0
LDC_MIN_MM = 200.0
LDH_MIN_MM_2019 = 150.0
LDH_MIN_MM_2025 = 152.0  # Corrected per ACI official confirmation (6 inches = 152.4 mm → 152 mm)
KTR_FACTOR = 40.0      # ACI Eq. 25.4.2.4b (SI)
LDT_C_2019 = 31.0      # ACI 318-19 §25.4.4.2 (SI)
LDT_C_2025 = 38.0      # ACI 318-25 §25.4.4.2 (SI)


# ---------------- Conversions ----------------
def to_si_stress_from_display(x: float, u: str) -> float:
    """Convert user display stress (MPa or kgf/cm2) to MPa."""
    return x if u == "SI" else x * MPA_PER_KGFCM2


def to_si_len_from_display(x: float, u: str) -> float:
    """Convert user display length (mm or cm) to mm."""
    return x if u == "SI" else x * 10.0


def from_si_len_to_display(mm: float, u: str) -> float:
    """Convert length in mm to display (mm or cm)."""
    return mm if u == "SI" else mm / 10.0


def sqrt_fc_cap(fcm: float) -> float:
    """Return sqrt(fc') with upper cap for numerical stability."""
    return min(math.sqrt(max(fcm, 0.0)), SQRT_FC_CAP_MPA)


def clamp(x: float, a: float, b: float) -> float:
    return max(a, min(b, x))


def parse_bars_mm(s: str) -> List[float]:
    """Parse comma-separated list of bar diameters in mm."""
    out: List[float] = []
    if not s:
        return out
    for x in s.split(","):
        try:
            out.append(float(x.strip()))
        except Exception:
            pass
    return out


# ---------------- Geometry & Ktr ----------------
def cb_from_geom(cover: float, spacing: float, so: float, db: float):
    """
    Compute c_b per ACI 25.4.2.3 as minimum of:
      - clear cover to tension face,
      - half clear spacing between bars,
      - half distance to edge or next bar layer (s_o / 2),
    with a lower bound of 0.5 db.
    """
    val_cover = cover
    val_spacing = spacing / 2.0
    val_so = so / 2.0 if (so and so > 0) else 999999.0
    cb = min(val_cover, val_spacing, val_so)
    cb = max(cb, 0.5 * db)

    if cb == val_cover:
        gov = "Cover"
    elif cb == val_spacing:
        gov = "Spacing/2"
    else:
        gov = "s_o/2"
    return cb, gov


def ktr_aci(Atr_mm2: float, s_mm: float, n_long: int) -> float:
    """
    ACI 318-19 Eq. 25.4.2.4b (SI):
        Ktr = 40 * Atr / (s * n)
    """
    if not (Atr_mm2 > 0 and s_mm > 0 and n_long > 0):
        return 0.0
    return KTR_FACTOR * Atr_mm2 / (s_mm * n_long)


# ---------------- ψ factors ----------------
def psi_s_tension(db_mm: float) -> float:
    """
    Size factor ψs for tension development (Table 25.4.2.4, SI).
    For db <= 19 mm, 0.8; otherwise 1.0.
    """
    return 0.8 if db_mm <= 19.0 else 1.0


def psi_e_tension(epoxy: bool, cover_mm: float, spacing_mm: float, db_mm: float) -> float:
    """
    Epoxy factor ψe (Table 25.4.2.4):
      - 1.0 for uncoated
      - 1.2 if epoxy, with good cover/spacing
      - 1.5 if epoxy, with poor cover/spacing
    """
    if not epoxy:
        return 1.0
    good = (cover_mm >= 3.0 * db_mm) and (spacing_mm >= 6.0 * db_mm)
    return 1.2 if good else 1.5


def cap_psi_t_times_psi_e(psi_t: float, psi_e: float) -> float:
    """ACI limit: ψt * ψe <= 1.7."""
    return (1.7 / psi_t) if psi_t * psi_e > 1.7 else psi_e


# ==========================================
# 2. CORE DEVELOPMENT FUNCTIONS (ACI-aligned)
# ==========================================
def ld_general_mm(
    fc: float,
    fy: float,
    db: float,
    lam: float,
    psi_t: float,
    psi_e: float,
    psi_s: float,
    cb: float,
    Ktr: float,
    psi_g: float,
) -> Dict[str, float | str]:
    """
    General tension development length ℓd – ACI 318-19 Eq. 25.4.2.4(a) (SI).

    ℓd = [ fy / (1.1 λ √f'c) · (ψt ψe ψs ψg) / ((c_b + Ktr)/db) ] db
    with β = (c_b + Ktr)/db ≤ 2.5  and ℓd ≥ 300 mm.
    """
    psi_e_eff = cap_psi_t_times_psi_e(psi_t, psi_e)
    rt = sqrt_fc_cap(fc)
    mods = psi_t * psi_e_eff * psi_s * psi_g
    confinement_term = (cb + Ktr) / db
    beta = clamp(confinement_term, 0.0, 2.5)
    if beta <= 0.0:
        raise ValueError("Invalid ACI tension development confinement term: (cb + Ktr) / db must be greater than zero.")
    denom = 1.1 * lam * rt
    eq_a = (fy / denom) * (mods / beta) * db
    ld = max(eq_a, LD_MIN_MM)

    return {
        "ld": ld,
        "raw": eq_a,
        "beta": beta,
        "confinement_raw": confinement_term,  # previous
        "beta_raw": confinement_term,         # for new report
        "is_capped": (confinement_term >= 2.5),
        "denom": denom,
        "mods": mods,
        "rt": rt,
        "psi_e_eff": psi_e_eff,   # for Detailed Report
        "psi_t": psi_t,
        "psi_e": psi_e,
        "psi_s": psi_s,
        "psi_g": psi_g,
        "cb": cb,
        "Ktr": Ktr,
    }


def ld_simplified_mm(
    fc: float,
    fy: float,
    db: float,
    lam: float,
    psi_t: float,
    psi_e: float,
    psi_g: float,
    db_small: bool,
    good_geom: bool,
) -> Dict[str, float | str]:
    """
    Simplified tension development ℓd – ACI 318-19/25 Table 25.4.2.3 (SI).

    Implemented via Table 25.4.2.3 coefficients (SI):
      - Denominator factor α = 2.1, 1.7, 1.4, or 1.1 depending on bar size and spacing/cover case.
      - Numerator multiplier = 1.0 for 'good bond/spacing & cover' rows, and 1.5 for 'Other cases' row.
    """
    psi_e_eff = cap_psi_t_times_psi_e(psi_t, psi_e)
    rt = sqrt_fc_cap(fc)

    # Select α per table
    if good_geom:
        denom_factor = 2.1 if db_small else 1.7
    else:
        denom_factor = 1.4 if db_small else 1.1

    denom_val = denom_factor * lam * rt
    psi_product = psi_t * psi_e_eff * psi_g
    num_factor = 1.0 if good_geom else 1.5  # Table 25.4.2.3 'Other cases'
    raw = (num_factor * fy / denom_val) * psi_product * db
    ld_val = max(raw, LD_MIN_MM)

    return {
        "ld": ld_val,
        "raw": raw,
        "denom_factor": denom_factor,
        "denom": denom_val,
        "num_factor": num_factor,
        "psi_t": psi_t,
        "psi_e": psi_e,
        "psi_e_eff": psi_e_eff,
        "psi_g": psi_g,
        "rt": rt,
        "db_small": db_small,
        "good_geom": good_geom,
        "mods": psi_product,
    }



def ldc_compression_mm(
    fc: float,
    fy: float,
    db: float,
    lam: float,
    psi_r: float,
    edition: str = "ACI318-2019",
) -> Dict[str, float | str]:
    """
    Compression development length ℓdc for deformed bars/wires in compression (SI).

    ACI 318-19 §25.4.9.2(a):
        ℓdc = max[
            (0.24 * fy * ψr / (λ * √f'c)) * db ,
            0.043 * fy * ψr * db ,
            200 mm
        ]

    ACI 318-25 §25.4.9.2(a):
        ℓdc = max[
            (fy * ψr / (4.2 * λ * √f'c)) * db ,
            0.043 * fy * ψr * db ,
            200 mm
        ]

    Notes:
    - Both editions require ℓdc ≥ 200 mm per §25.4.9.1(b).
    - Only the coefficient in term (a) differs (0.24 vs 1/4.2).
    """
    rt = sqrt_fc_cap(fc)

    if edition and edition.startswith("ACI318-2025"):
        c_a = 1.0 / 4.2  # ACI 318-25 §25.4.9.2(a)
    else:
        c_a = 0.24       # ACI 318-19 §25.4.9.2(a)

    term_a = c_a * fy * psi_r / (lam * rt) * db
    term_b = 0.043 * fy * psi_r * db

    base = max(term_a, term_b)
    ldc = max(base, LDC_MIN_MM)

    return {
        "ldc": ldc,
        "term_a": term_a,
        "term_b": term_b,
        "c_a": c_a,
    }



def ldh_hook_mm(
    fc: float,
    fy: float,
    db: float,
    lam: float,
    epoxy: bool,
    edition: str,
    has_conf: bool,
    cover_ok: bool,
    cover_top_ok_for_90: bool,  # reserved, not used explicitly
    hook_angle: int,
) -> Dict[str, float | str]:
    """
    Standard hooked bar development length ℓdh in tension.

    ACI 318-19 §25.4.3.1(a) (SI):
        ℓdh = max[ (fy ψe ψr ψo ψc / (23 λ √fc')) db^1.5 , 8 db , 150 mm ]

    ACI 318-25 §25.4.3.1(a) (SI) — corrected per ACI official confirmation (Jan 2026):
        ℓdh = max[ (fy ψe ψs ψcc ψr / (4.16 λ √fc')) db , 8 db , 152 mm ]

    Note: The official ACI 318M-25 printed version contains an error (C=21, db^1.5).
          The corrected values (C=4.16, db^1.0, 152 mm) are implemented here as confirmed by ACI.
    """
    rt = sqrt_fc_cap(fc)

    psi_e = 1.2 if epoxy else 1.0
    psi_r = 0.8 if has_conf else 1.0
    psi_s = 1.0  # default; used in 2025 (and reported for 2019 as 1.0)
    psi_cc = 1.0  # default; used in 2025 (and reported for 2019 as psi_c)

    if edition and edition.startswith("ACI318-2019"):
        # ACI 318-19: ψe, ψr, ψo, ψc; C = 23
        C = 23.0
        psi_o = 1.0
        psi_c = 0.7 if (hook_angle == 90 and cover_ok) else 1.0
        psi_cc = psi_c
        psi_prod = psi_e * psi_r * psi_o * psi_c
        power_db = 1.5
        min_abs = LDH_MIN_MM_2019  # 150 mm
    else:
        # ACI 318-2025 — corrected per ACI confirmation
        C = 4.16
        psi_s = 1.0
        psi_cc = 0.7 if (hook_angle == 90 and cover_ok) else 1.0

        # V68: Educational overrides for ACI 318-25 (SI) hooked-bar factors (optional)
        edu = st.session_state.get("edu_modifiers", {}) if "st" in globals() else {}
        if str(edition).startswith("ACI318-2025") and isinstance(edu, dict) and edu:
            try:
                tbl = get_ch25_tables_for_edition(edition)["25.4.3.2"]

                grp = edu.get("hook_psi_s_group", "Auto (from db)")
                if isinstance(grp, str) and grp.startswith("Auto"):
                    grp_key = hook_psi_s_group_2025_by_db_mm(db)
                elif "No.29" in grp:
                    grp_key = "no29_and_smaller"
                elif "No.32" in grp:
                    grp_key = "no32_no36"
                elif "No.43" in grp:
                    grp_key = "no43"
                elif "No.57" in grp:
                    grp_key = "no57"
                else:
                    grp_key = hook_psi_s_group_2025_by_db_mm(db)

                psi_s = float(tbl["psi_s"][grp_key])
                psi_cc = float(tbl["psi_cc"][edu.get("hook_psi_cc_key", "other")])
                psi_r = float(tbl["psi_r"][edu.get("hook_psi_r_key", "other")])
            except Exception:
                pass  # fail-safe

        psi_prod = psi_e * psi_s * psi_cc * psi_r
        psi_o = 1.0
        psi_c = psi_cc
        power_db = 1.0
        min_abs = LDH_MIN_MM_2025  # 152 mm

    # Bar diameter exponent differs by edition (per printed equations)
    base = (fy * psi_prod) / (C * lam * rt) * (db ** power_db)

    # Minimum per §25.4.3.1(b): 8db (all standard hook angles)
    angle_min = 8.0 * db
    ldh = max(base, angle_min, min_abs)

    if ldh == base:
        governing = "ACI main term"
    elif ldh == angle_min:
        governing = "8 db term"
    else:
        governing = f"{min_abs} mm minimum"
    return {
        "ldh": ldh,
        "base": base,
        "psi_e": psi_e,
        "psi_s": psi_s,
        "psi_cc": psi_cc,
        "psi_r": psi_r,
        "psi_o": psi_o,
        "psi_c": psi_c,
        "C": C,
        "C_used": C,
        "db_exponent": power_db,
        "rt": rt,
        "angle_min": angle_min,
        "governing": governing,
    }

def ldt_headed_mm(
    fc: float,
    fy: float,
    db: float,
    epoxy: bool,
    edition: str,
    psi_p: float,
    psi_o: float,
    psi_c: float,
) -> Dict[str, float | str]:
    """
    Headed bar development length ℓdt in tension.

    ACI 318-19 §25.4.4.2 (SI):
        ℓdt = max[ (fy ψe ψp ψo ψc / (31 √fc')) db^1.5 , 8 db , 150 mm ]

    ACI 318-25 §25.4.4.2 (SI):
        ℓdt = max[ (fy ψe ψp ψo ψc / (38 √fc')) db^1.5 , 8 db , 150 mm ]

    Per §25.4.4.1 headed bars are limited to normalweight concrete,
    so λ does not appear in the formula.
    """
    rt = sqrt_fc_cap(fc)
    psi_e = 1.2 if epoxy else 1.0

    if edition and edition.startswith("ACI318-2019"):
        C = LDT_C_2019
    else:
        C = LDT_C_2025

    base = (fy * psi_e * psi_p * psi_o * psi_c) / (C * rt) * (db ** 1.5)
    ldt = max(base, 8.0 * db, LDH_MIN_MM_2019)

    return {
        "ldt": ldt,
        "base": base,
        "psi_e": psi_e,
        "psi_p": psi_p,
        "psi_o": psi_o,
        "psi_c": psi_c,
        "C_used": C,
    }


def splice_class_auto(as_ratio: float, pct_spliced: float):
    """
    Auto classification of splice per ACI Table 25.5.2.1:
      - Class A if As,prov / As,req >= 2.0 and ≤ 50% bars spliced,
      - otherwise Class B.
    """
    try:
        as_ratio = float(as_ratio)
        pct_spliced = float(pct_spliced)
    except (TypeError, ValueError) as exc:
        raise ValueError("Invalid splice-class inputs: As ratio and percent spliced must be numeric.") from exc
    if not (math.isfinite(as_ratio) and math.isfinite(pct_spliced)):
        raise ValueError("Invalid splice-class inputs: values must be finite.")
    if as_ratio < 0.0 or pct_spliced < 0.0 or pct_spliced > 100.0:
        raise ValueError("Invalid splice-class inputs: As ratio must be nonnegative and percent spliced must be between 0 and 100.")
    if as_ratio >= 2.0 and pct_spliced <= 50.0:
        return "Class A", 1.0
    return "Class B", 1.3



def lap_splice_length_compression_mm(
    fc: float,
    fy: float,
    db: float,
    lst_tension: float,
    edition: str = "ACI318-2019",
) -> Dict[str, float | str]:
    """
    Compression lap splice length ℓsc for deformed bars in compression (SI).

    ACI 318-19 §25.5.5.1:
      (a) For fy ≤ 420 MPa:  ℓsc = longer of (0.071 fy db) and 300 mm
      (b) For 420 < fy ≤ 550 MPa: ℓsc = longer of ((0.13 fy − 24) db) and 300 mm
      (c) For fy > 550 MPa: ℓsc = longer of ((0.13 fy − 24) db) and ℓst per §25.5.2.1

    ACI 318-19 §25.5.5.1 (last paragraph):
      For fc' < 21 MPa, increase lap length by one-third.

    ACI 318-25 §25.5.5.1 contains the same piecewise expressions and fc' < 21 MPa
    increase-by-one-third requirement (verified from the official PDF in this project).

    Notes / limits:
      - §25.5.5.2: compression lap splices are not used for bars larger than No. 36,
        except as permitted in §25.5.5.3 and sized per §25.5.5.4 (different bar sizes).
      - In SI, No. 36 corresponds to nominal diameter 36 mm. This function therefore
        enforces db > 36 mm as "NOT PERMITTED" unless a different-size splice per §25.5.5.4
        is explicitly evaluated elsewhere.
    """
    # Bar size limitation (§25.5.5.2–§25.5.5.4)
    if db > 36.0:
        return {
            "lsc": "NOT PERMITTED (ACI §25.5.5.2–§25.5.5.4)",
            "governing": "bar_size_limit",
        }

    # Piecewise per §25.5.5.1(a)–(c)
    if fy <= 420.0:
        term = 0.071 * fy * db  # mm
        lsc_base = max(term, 300.0)
        branch = "a"
    elif fy <= 550.0:
        term = (0.13 * fy - 24.0) * db  # mm
        lsc_base = max(term, 300.0)
        branch = "b"
    else:
        term = (0.13 * fy - 24.0) * db  # mm
        lsc_base = max(term, float(lst_tension))
        branch = "c"

    # fc' < 21 MPa → increase by one-third (§25.5.5.1)
    if fc < 21.0:
        lsc = (4.0 / 3.0) * lsc_base
        fc_adj = True
    else:
        lsc = lsc_base
        fc_adj = False

    return {
        "lsc": lsc,
        "term": term,
        "min_300": 300.0,
        "lst_tension": float(lst_tension),
        "branch": branch,
        "fc_lt_21": fc_adj,
        "governing": "calculated",
    }

def calculate_exact_hook_geometry(db: float, usage: str):
    """
    Geometric hook dimensions (D, extension, total length) based on ACI tables
    for main bars vs stirrups/ties (Table 25.3.1 and 25.3.2, SI).
    """
    u = usage.lower()
    is_stirrup = (u in ["stirrup", "tie", "hoop"])

    # Bend diameter factor k_d from Tables 25.3.1 and 25.3.2
    if is_stirrup:
        # Stirrups / ties: group by bar size (No.10–16 vs No.19–25)
        if db <= 16.0:
            bend_factor = 4.0
        elif db <= 25.0:
            bend_factor = 6.0
        else:
            bend_factor = 10.0
    else:
        # Main bars
        if db <= 25.0:
            bend_factor = 6.0
        elif db <= 36.0:
            bend_factor = 8.0
        else:
            bend_factor = 10.0

    D = bend_factor * db

    # Hook extensions
    if is_stirrup:
        # Table 25.3.2 – stirrups / ties
        if db <= 16.0:
            # ≈ No.10–16: max(6db, 75 mm)
            ext_90 = max(6.0 * db, 75.0)
        else:
            # ≈ No.19–25: 12 db
            ext_90 = 12.0 * db
    else:
        # Main tension bars – Table 25.3.1
        ext_90 = 12.0 * db

    ext_135 = max(6.0 * db, 75.0)
    ext_180 = max(4.0 * db, 65.0)

    L1_total = (D / 2.0) + db + ext_90
    L2_total = (D / 2.0) + db + ext_135
    L3_total = (D / 2.0) + db + ext_180

    return {
        "D_mm": D,
        "bend_factor": bend_factor,
        "L1_ext": ext_90,
        "L1_total": L1_total,
        "L2_ext": ext_135,
        "L2_total": L2_total,
        "L3_ext": ext_180,
        "L3_total": L3_total,
    }


# ==========================================
# 3. REPORT & EXCEL (CACHED)
# ==========================================
@st.cache_data
def generate_excel_blob(df_len, df_geo_m, df_geo_s):
    """Create an in-memory Excel workbook as bytes.

    Notes:
      - Tries xlsxwriter first (preferred), then falls back to openpyxl.
      - Raises the last exception if all engines fail.
    """
    engines = ["xlsxwriter", "openpyxl"]
    last_err = None

    for eng in engines:
        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine=eng) as writer:
                df_len.to_excel(writer, sheet_name="Lengths", index=False)
                df_geo_m.to_excel(writer, sheet_name="Geometry Main", index=False)
                df_geo_s.to_excel(writer, sheet_name="Geometry Stirrup", index=False)
            return output.getvalue()
        except Exception as e:
            last_err = e
            continue

    raise last_err


def generate_html_report(inputs: Dict[str, Any], res_df: pd.DataFrame) -> str:
    """
    Generate a print-ready HTML calculation booklet (English-only, LTR).
    Modes:
      - Professional: concise, checker/client ready
          - Educational: adds step-by-step notes and explanations
    """
    # -----------------------------
    # Mode normalization (accept legacy emoji labels)
    # -----------------------------
    raw_mode = str(inputs.get("user_mode", "Professional") or "Professional")
    mode_clean = raw_mode
    # remove common emoji prefixes and extra whitespace
    mode_clean = re.sub(r"[^\w\s-]", " ", mode_clean, flags=re.UNICODE)
    mode_clean = re.sub(r"\s+", " ", mode_clean).strip().lower()

    is_pro = "pro" in mode_clean
    is_int = False  # FIX12B: Intermediate mode removed from visible UI
    is_edu = "edu" in mode_clean or bool(inputs.get("is_edu", False))

    if not (is_pro or is_int or is_edu):
        is_pro = True

    if is_pro:
        mode_label = "Professional"
    else:
        mode_label = "Educational"

    # -----------------------------
    # Base inputs
    # -----------------------------
    edition = str(inputs.get("edition", "ACI 318-2019"))
    units = str(inputs.get("units", "SI"))
    fc = inputs.get("fc_d", "-")
    fy = inputs.get("fy_d", "-")
    lam = inputs.get("lam", "-")

    cover_mm = inputs.get("cover_mm", None)
    spacing_mm = inputs.get("spacing_mm", None)
    so_d = inputs.get("so_d", None)

    # ---- STABILITY ALIASES (for legacy variable names used in report/UI)
    cover_d = inputs.get('cover_d', None)
    spacing_d = inputs.get('spacing_d', None)
    cover = cover_d
    spacing = spacing_d
    so = so_d


    cb_mm = inputs.get("cb_mm", None)
    cb_gov = inputs.get("cb_gov", None)

    beta_raw = inputs.get("beta_raw", None)
    beta_eff = inputs.get("beta_eff", None)
    beta_is_capped = inputs.get("beta_is_capped", None)

    psi_e = inputs.get("psi_e", None)
    psi_s = inputs.get("psi_s", None)

    # top bar factor may be stored as psi_t_T / psi_t_B (tool uses both)
    psi_t_T = inputs.get("psi_t_T", None)
    psi_t_B = inputs.get("psi_t_B", None)

    Ktr_base = inputs.get("Ktr_base", None)
    alpha_used = inputs.get("alpha_used", None)
    special_seismic_region = bool(inputs.get("special_seismic_region", False))

    splice_mode = inputs.get("splice_mode", None)
    class_val = inputs.get("class_val", None)
    as_ratio = inputs.get("as_ratio", None)
    pct_spliced = inputs.get("pct_spliced", None)

    method = inputs.get("method", None)  # internal label if present

    # -----------------------------
    # Helpers
    # -----------------------------
    def _fmt(x, nd=2):
        try:
            if x is None or x == "-" or (isinstance(x, float) and (math.isnan(x) or math.isinf(x))):
                return "-"
            if isinstance(x, (int, np.integer)):
                return f"{int(x)}"
            if isinstance(x, (float, np.floating)):
                return f"{float(x):.{nd}f}"
            return str(x)
        except Exception:
            return str(x)

    def _html_escape(s: Any) -> str:
        s = "" if s is None else str(s)
        return (s.replace("&", "&amp;")
                 .replace("<", "&lt;")
                 .replace(">", "&gt;"))

    def _df_to_html_table(df: pd.DataFrame) -> str:
        if df is None or df.empty:
            return "<p class='muted'>No results available.</p>"
        # keep it simple and robust (no external CSS frameworks)
        cols = list(df.columns)
        th = "".join([f"<th>{_html_escape(c)}</th>" for c in cols])
        rows = []
        for _, r in df.iterrows():
            tds = "".join([f"<td>{_html_escape(r.get(c, ''))}</td>" for c in cols])
            rows.append(f"<tr>{tds}</tr>")
        return f"""
        <table class="table">
          <thead><tr>{th}</tr></thead>
          <tbody>
            {''.join(rows)}
          </tbody>
        </table>
        """

    # Date/time
    dt_str = _dt.now().strftime("%Y-%m-%d %H:%M:%S")

    # -----------------------------
    # Core report (English-only, LTR)
    # -----------------------------
    # Note: MathJax renders TeX; keep LTR layout with left alignment.
    html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0" />
<title>Rebar Development & Splice Length Report</title>

<script>
  window.MathJax = {{
    tex: {{ inlineMath: [['$', '$'], ['\\\\(', '\\\\)']] }},
    options: {{ skipHtmlTags: ['script','noscript','style','textarea','pre','code'] }}
  }};
</script>
<script src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>

<style>
  :root {{
    --border: #2b2b2b;
    --muted: #555;
  }}
  html, body {{
    direction: ltr;
    text-align: left;
    font-family: Arial, Helvetica, sans-serif;
    color: #000;
    background: #fff;
    margin: 0;
    padding: 0;
  }}
  .page {{
    padding: 18mm 16mm;
  }}
  h1 {{
    font-size: 22pt;
    margin: 0 0 6mm 0;
    letter-spacing: 0.2px;
  }}
  .sub {{
    color: var(--muted);
    font-size: 10.5pt;
    margin-bottom: 8mm;
    line-height: 1.35;
  }}
  h2 {{
    font-size: 14pt;
    margin: 10mm 0 3mm 0;
    border-bottom: 1px solid var(--border);
    padding-bottom: 2mm;
  }}
  h3 {{
    font-size: 12pt;
    margin: 6mm 0 2mm 0;
  }}
  .grid {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 6mm;
  }}
  .card {{
    border: 1px solid var(--border);
    padding: 4mm 4mm;
    border-radius: 2mm;
  }}
  .label {{
    color: var(--muted);
    font-size: 9.5pt;
    margin: 0 0 1mm 0;
  }}
  .value {{
    font-size: 11pt;
    margin: 0;
  }}
  .table {{
    width: 100%;
    border-collapse: collapse;
    margin-top: 2mm;
    font-size: 10pt;
  }}
  .table th, .table td {{
    border: 1px solid var(--border);
    padding: 2.2mm 2.4mm;
    vertical-align: top;
  }}
  .table thead th {{
    background: #f2f2f2;
    font-weight: 700;
  }}
  .muted {{
    color: var(--muted);
  }}
  .kpi {{
    font-size: 12pt;
    font-weight: 700;
  }}
  .small {{
    font-size: 9.5pt;
  }}
  .mono {{
    font-family: "Consolas", "Courier New", monospace;
  }}

  @media print {{
    @page {{ size: A4; margin: 14mm; }}
    .page {{ padding: 0; }}
    a {{ color: #000; text-decoration: none; }}
  }}

</style>
</head>

<body>
<div class="page">

<h1>Rebar Development & Splice Length Calculations</h1>
<div class="sub">
  <div><b>Design Code:</b> {_html_escape(edition)} &nbsp;&nbsp; <b>Units:</b> {_html_escape(units)} &nbsp;&nbsp; <b>Mode:</b> {_html_escape(mode_label)}</div>
  <div><b>Generated:</b> {_html_escape(dt_str)}</div>
</div>

<h2>1. Design Basis</h2>
<ul class="small">
  <li>ACI 318-19 / ACI 318M-25 Chapter 25: Reinforcement Details (development, hooks, and splices as applicable).</li>
  <li>Material properties and geometry are taken from the input data listed in Section 2.</li>
  <li>Report is generated by the software tool; final detailing remains the responsibility of the Engineer of Record.</li>
</ul>

<h2>2. Input Data</h2>
<div class="grid">
  <div class="card">
    <p class="label">Concrete strength</p>
    <p class="value">$f'_c$ = {_html_escape(fc)} ({_html_escape(units)})</p>
    <p class="label" style="margin-top:3mm;">Steel yield strength</p>
    <p class="value">$f_y$ = {_html_escape(fy)} ({_html_escape(units)})</p>
    <p class="label" style="margin-top:3mm;">Lightweight factor</p>
    <p class="value">$\\lambda$ = {_html_escape(lam)}</p>
  </div>

  <div class="card">
    <p class="label">Clear cover</p>
    <p class="value">{_fmt(cover_mm, 1)} mm</p>
    <p class="label" style="margin-top:3mm;">Clear bar spacing, s_clear</p>
    <p class="value">{_fmt(spacing_mm, 1)} mm</p>
    <p class="label" style="margin-top:3mm;">so / nearest surface distance</p>
    <p class="value">{_fmt(so_d, 1)} {('mm' if units.upper()=='SI' else 'in')}</p>
  </div>
</div>

<h3>2.1 Derived Geometry Parameters</h3>
<table class="table">
  <thead>
    <tr>
      <th>Parameter</th>
      <th>Value</th>
      <th>Notes</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>$c_b$</td>
      <td>{_fmt(cb_mm, 2)} mm</td>
      <td>Governing component: {_html_escape(cb_gov)}</td>
    </tr>
    <tr>
      <td>$K_{{tr}}$</td>
      <td>{_fmt(Ktr_base, 4)}</td>
      <td>Computed from transverse reinforcement configuration (if provided)</td>
    </tr>
    <tr>
      <td>$\\dfrac{{c_b + K_{{tr}}}}{{d_b}}$</td>
      <td>{_fmt(beta_raw, 3) if beta_raw is not None else "-"}</td>
      <td>ACI: $(c_b + K_{{tr}})/d_b$</td>
    </tr>
    <tr>
      <td>$\\min\\left[\\dfrac{{c_b + K_{{tr}}}}{{d_b}},\\,2.5\\right]$</td>
      <td>{_fmt(beta_eff, 3) if beta_eff is not None else "-"}</td>
      <td>Cap applied: {"YES" if beta_is_capped else "NO" if beta_is_capped is not None else "-"}</td>
    </tr>
    <tr>
      <td>$\\alpha$</td>
      <td>{_fmt(alpha_used, 3)}</td>
      <td>Bond coefficient / method flag: {_html_escape(method)}</td>
    </tr>
    <tr>
      <td>Special seismic detailing region</td>
      <td>{"YES" if special_seismic_region else "NO"}</td>
      <td>Gating only; verify Chapter 18 requirements separately</td>
    </tr>
  </tbody>
</table>

<h2>3. Results Summary</h2>
{_df_to_html_table(res_df)}

<h2>4. Methodology and Governing Equations</h2>
<p class="small muted">
The following equations are shown for documentation of the implemented approach. Actual governing terms and limits are applied as implemented in the tool.
</p>

<h3>4.1 Tension Development Length ($\\ell_d$)</h3>
<p class="small">
$\\ell_d$ for deformed bars in tension is calculated per ACI Chapter 25 using the standard development-length expression with applicable modification factors and the confinement term:
</p>
<p class="mono small">
$\\ell_d \\propto \\dfrac{{f_y}}{{\\lambda \\sqrt{{f'_c}}}}\\;\\dfrac{{\\psi_t\\,\\psi_e\\,\\psi_s}}{{(c_b + K_{{tr}})/d_b}}\\;d_b$
</p>

<table class="table">
  <thead>
    <tr><th>Factor</th><th>Value</th><th>Meaning</th></tr>
  </thead>
  <tbody>
    <tr><td>$\\psi_t$ (top bar)</td><td>{_fmt(psi_t_T, 3) if psi_t_T not in [None,'-'] else _fmt(psi_t_B, 3)}</td><td>Top-bar modification factor (if applicable)</td></tr>
    <tr><td>$\\psi_e$ (epoxy)</td><td>{_fmt(psi_e, 3)}</td><td>Epoxy modification factor (if applicable)</td></tr>
    <tr><td>$\\psi_s$ (bar size)</td><td>{_fmt(psi_s, 3)}</td><td>Bar-size modification factor (if applicable)</td></tr>
  </tbody>
</table>

<h3>4.2 Lap Splice Length ($\\ell_s$)</h3>
<table class="table">
  <thead>
    <tr><th>Splice Input</th><th>Value</th></tr>
  </thead>
  <tbody>
    <tr><td>Splice mode</td><td>{_html_escape(splice_mode)}</td></tr>
    <tr><td>Lap splice class</td><td>{_html_escape(class_val)}</td></tr>
    <tr><td>$A_s$ provided / required</td><td>{_fmt(as_ratio, 3)}</td></tr>
    <tr><td>% bars spliced</td><td>{_fmt(pct_spliced, 1)}%</td></tr>
  </tbody>
</table>

"""

    # -----------------------------
    # Mode-specific content
    # -----------------------------
    if is_int or is_edu:
        html += """
<h2>5. Engineering Checks</h2>
<ul class="small">
  <li>Confirm input geometry is consistent with detailing drawings: clear cover, clear longitudinal-bar spacing (s_clear), and transverse reinforcement spacing used for Ktr (s_Ktr).</li>
  <li>Verify transverse reinforcement parameters used for $K_{{tr}}$ match the actual detailing (legs, size, spacing).</li>
  <li>Confirm splice class selection is consistent with bar arrangement and percentage of splices in a given section.</li>
</ul>
"""

    if is_edu:
        html += f"""
<h2>6. Educational Notes</h2>
<ul class="small">
  <li><b>$c_b$</b> is governed by the minimum of clear cover and one-half clear spacing. Here, $c_b = {_fmt(cb_mm,2)}$ mm.</li>
  <li><b>$K_{{tr}}$</b> improves bond when confinement is present. If transverse reinforcement is not provided, $K_{{tr}}=0$.</li>
  <li><b>Upper limits</b> in the confinement term can govern; additional cover/spacing beyond the code cap does not reduce $\\ell_d$ further.</li>
  <li>Use this section for training; for submittals, use Professional mode.</li>
</ul>
"""

    # Close
    html += """
<h2>End of Report</h2>
<p class="small muted">
This report is valid only for the input parameters listed. Final design decisions and detailing remain the responsibility of the Engineer of Record.
</p>

</div>
</body>
</html>
"""
    return html


def inject_custom_css() -> None:
    """
    Material Design 3 dark theme — UI-only, no calculation logic modified.
    """
    st.markdown(
        """
<style>
/* ── Material Design 3 token surface ── */
:root {
    /* Backgrounds / surfaces */
    --aci-bg0:    #0d0d14;
    --aci-bg1:    #12121e;
    --aci-bg2:    #181826;
    --aci-card:   rgba(30, 30, 50, 0.95);
    --aci-card2:  rgba(22, 22, 38, 0.97);
    --aci-surf-lo: rgba(26, 26, 42, 0.98);
    --aci-surf-hi: rgba(38, 38, 60, 0.98);

    /* Borders */
    --aci-border:       rgba(130, 170, 255, 0.18);
    --aci-border-focus: rgba(130, 170, 255, 0.60);

    /* Text */
    --aci-text:   #e8eaff;
    --aci-muted:  #9a9ec8;
    --aci-subtle: #6a6e94;

    /* Accent palette  (indigo-blue primary) */
    --aci-blue:   #82aaff;
    --aci-cyan:   #4dc9f0;
    --aci-green:  #69f0ae;
    --aci-yellow: #ffd740;
    --aci-red:    #ff6e6e;
    --aci-violet: #c3b1e1;

    /* Elevation shadows */
    --aci-shadow-sm: 0 2px  8px rgba(0,0,0,.35);
    --aci-shadow-md: 0 6px 20px rgba(0,0,0,.42);
    --aci-shadow-lg: 0 14px 40px rgba(0,0,0,.52);

    /* Radii */
    --r-sm: 10px;
    --r-md: 14px;
    --r-lg: 20px;
    --r-xl: 28px;
}

html, body, [data-testid="stAppViewContainer"] {
    background:
        radial-gradient(ellipse at 8% 0%,   rgba(82,130,255,0.12) 0%, transparent 40%),
        radial-gradient(ellipse at 95% 5%,  rgba(77,201,240,0.08) 0%, transparent 35%),
        radial-gradient(ellipse at 50% 100%, rgba(60,60,120,0.10) 0%, transparent 55%),
        linear-gradient(180deg, var(--aci-bg0) 0%, var(--aci-bg1) 100%) !important;
    color: var(--aci-text) !important;
}

[data-testid="stHeader"] {
    background: rgba(10, 10, 18, 0.82) !important;
    backdrop-filter: blur(18px) saturate(1.4);
    border-bottom: 1px solid var(--aci-border);
    box-shadow: var(--aci-shadow-sm);
}

[data-testid="stSidebar"] {
    background: linear-gradient(180deg, rgba(14,14,24,0.99) 0%, rgba(18,18,30,0.99) 100%) !important;
    border-right: 1px solid var(--aci-border);
    box-shadow: 4px 0 24px rgba(0,0,0,.40);
}

[data-testid="stSidebar"] * {
    color: var(--aci-text) !important;
}

.block-container {
    padding-top: 1.2rem !important;
    padding-bottom: 2rem !important;
    max-width: 1800px !important;
}

/* ── Hero banner ── */
.aci-hero {
    display: grid;
    grid-template-columns: 1.35fr .95fr;
    gap: 20px;
    padding: 22px 24px;
    margin: 10px 0 20px 0;
    background: linear-gradient(135deg, rgba(28, 28, 52, .97) 0%, rgba(16, 16, 30, .98) 100%);
    border: 1px solid rgba(130,170,255,.22);
    border-radius: var(--r-xl);
    box-shadow: var(--aci-shadow-lg), inset 0 1px 0 rgba(255,255,255,.04);
    overflow: hidden;
    position: relative;
}
.aci-hero::before {
    content: "";
    position: absolute; inset: 0;
    background: radial-gradient(ellipse at 0% 0%, rgba(82,130,255,.14) 0%, transparent 55%);
    pointer-events: none;
}

.aci-title {
    font-size: 32px;
    line-height: 1.14;
    font-weight: 900;
    letter-spacing: -0.5px;
    color: var(--aci-text);
    margin: 0 0 8px 0;
}

.aci-subtitle {
    color: var(--aci-muted);
    font-size: 14.5px;
    line-height: 1.60;
    max-width: 900px;
}

/* ── Logo badge ── */
.aci-logo {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 52px;
    height: 52px;
    border-radius: 15px;
    margin-bottom: 12px;
    background: linear-gradient(135deg, #3d7eff 0%, #4dc9f0 100%);
    font-weight: 900;
    font-size: 19px;
    color: #fff;
    box-shadow: 0 4px 18px rgba(82,130,255,.45), inset 0 1px 0 rgba(255,255,255,.20);
    letter-spacing: -0.5px;
}

/* ── Chips (Material filled-tonal style) ── */
.aci-chip {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    padding: 5px 11px;
    margin: 5px 4px 0 0;
    border-radius: 999px;
    border: 1px solid var(--aci-border);
    background: rgba(82,130,255,.10);
    color: var(--aci-muted);
    font-size: 11.5px;
    font-weight: 700;
    letter-spacing: .2px;
    transition: background .15s;
}
.aci-chip.ok   { color: #a6f0c6; background: rgba(105,240,174,.12); border-color: rgba(105,240,174,.30); }
.aci-chip.warn { color: #ffe08a; background: rgba(255,215,64,.12);  border-color: rgba(255,215,64,.28);  }
.aci-chip.info { color: #b8d8ff; background: rgba(82,130,255,.14);  border-color: rgba(82,130,255,.35);  }
.aci-chip.err  { color: #ffadad; background: rgba(255,110,110,.12); border-color: rgba(255,110,110,.28); }

/* ── Panel card (Material surface-container) ── */
.aci-panel {
    background: var(--aci-card);
    border: 1px solid var(--aci-border);
    border-radius: var(--r-lg);
    padding: 16px 18px;
    box-shadow: var(--aci-shadow-md), inset 0 1px 0 rgba(255,255,255,.03);
    backdrop-filter: blur(12px);
    margin-bottom: 14px;
    transition: box-shadow .2s, border-color .2s;
}
.aci-panel:hover {
    border-color: rgba(130,170,255,.30);
    box-shadow: var(--aci-shadow-lg);
}

.aci-panel-title {
    font-size: 11.5px;
    font-weight: 800;
    letter-spacing: .55px;
    text-transform: uppercase;
    color: var(--aci-blue);
    margin-bottom: 10px;
    display: flex;
    align-items: center;
    gap: 7px;
}
.aci-panel-title::before {
    content: "";
    display: inline-block;
    width: 3px; height: 14px;
    border-radius: 2px;
    background: var(--aci-blue);
    opacity: .8;
}

/* ── 3-column grid ── */
.aci-grid-3 {
    display: grid;
    grid-template-columns: repeat(3, minmax(0,1fr));
    gap: 10px;
}

/* ── Stat card ── */
.aci-stat {
    background: rgba(18, 18, 36, .70);
    border: 1px solid rgba(130,170,255,.14);
    border-radius: var(--r-md);
    padding: 13px 14px;
    transition: background .15s, border-color .15s;
}
.aci-stat:hover {
    background: rgba(28, 28, 52, .80);
    border-color: rgba(130,170,255,.28);
}
.aci-stat .label {
    color: var(--aci-muted);
    font-size: 11.5px;
    font-weight: 600;
    letter-spacing: .15px;
}
.aci-stat .value {
    color: var(--aci-text);
    font-size: 20px;
    font-weight: 900;
    margin-top: 5px;
    letter-spacing: -0.3px;
}

/* ── Streamlit Metric ── */
div[data-testid="stMetric"] {
    background: rgba(26, 26, 46, .82);
    border: 1px solid var(--aci-border);
    border-radius: var(--r-md);
    padding: 14px 16px;
    box-shadow: var(--aci-shadow-sm);
    transition: border-color .18s, box-shadow .18s;
}
div[data-testid="stMetric"]:hover {
    border-color: rgba(130,170,255,.35);
    box-shadow: var(--aci-shadow-md);
}

/* ── Tabs ── */
div[data-testid="stTabs"] button {
    color: var(--aci-muted) !important;
    font-weight: 700;
    letter-spacing: .15px;
    transition: color .15s;
}
div[data-testid="stTabs"] button[aria-selected="true"] {
    color: var(--aci-blue) !important;
    border-bottom-color: var(--aci-blue) !important;
}
div[data-testid="stTabs"] button:hover {
    color: var(--aci-text) !important;
}

/* ── Buttons ── */
.stButton > button {
    border-radius: var(--r-sm) !important;
    background: linear-gradient(135deg, #2c5fff 0%, #1a3ed4 100%) !important;
    color: #ffffff !important;
    border: 1px solid rgba(82,130,255,.55) !important;
    font-weight: 800 !important;
    letter-spacing: .15px !important;
    box-shadow: 0 2px 10px rgba(44,95,255,.30), inset 0 1px 0 rgba(255,255,255,.12) !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #3d6fff 0%, #2451e8 100%) !important;
    border-color: rgba(130,170,255,.80) !important;
    box-shadow: 0 4px 18px rgba(44,95,255,.45), inset 0 1px 0 rgba(255,255,255,.14) !important;
}

.stDownloadButton > button {
    border-radius: var(--r-sm) !important;
    background: rgba(24, 24, 42, 0.96) !important;
    color: var(--aci-blue) !important;
    border: 1px solid rgba(82,130,255,.38) !important;
    font-weight: 700 !important;
    box-shadow: var(--aci-shadow-sm) !important;
}
.stDownloadButton > button:hover {
    background: rgba(30, 30, 54, 0.98) !important;
    border-color: rgba(82,130,255,.65) !important;
    box-shadow: var(--aci-shadow-md) !important;
}

hr { border-color: rgba(130,170,255,.12) !important; }

/* ── Widget labels ── */
[data-testid="stWidgetLabel"],
[data-testid="stWidgetLabel"] *,
label, label *,
div[data-testid="stNumberInput"] label,
div[data-testid="stNumberInput"] label *,
div[data-testid="stTextInput"] label,
div[data-testid="stTextInput"] label *,
div[data-testid="stSelectbox"] label,
div[data-testid="stSelectbox"] label *,
div[data-testid="stSlider"] label,
div[data-testid="stSlider"] label *,
div[data-testid="stCheckbox"] label,
div[data-testid="stCheckbox"] label *,
div[data-testid="stRadio"] label,
div[data-testid="stRadio"] label * {
    color: #c8d4f0 !important;
    opacity: 1 !important;
}

/* Caption / help text */
[data-testid="stCaptionContainer"],
[data-testid="stCaptionContainer"] *,
small, .stCaption, .help,
.css-q8sbsg, .css-1inwz65 {
    color: var(--aci-muted) !important;
    opacity: 1 !important;
}

/* Markdown body */
[data-testid="stMarkdownContainer"],
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] span,
[data-testid="stMarkdownContainer"] li {
    color: var(--aci-text) !important;
}

/* Input fields — Material filled style */
div[data-testid="stNumberInput"] input,
div[data-testid="stTextInput"] input,
div[data-testid="stDateInput"] input,
div[data-testid="stTimeInput"] input,
textarea {
    background-color: rgba(18, 18, 36, 0.95) !important;
    color: #f0f2ff !important;
    border: 1px solid rgba(130, 170, 255, 0.30) !important;
    border-radius: var(--r-sm) !important;
    box-shadow: inset 0 1px 3px rgba(0,0,0,.25) !important;
    min-height: 2.35rem !important;
    font-size: 0.94rem !important;
    transition: border-color .15s, box-shadow .15s !important;
}
div[data-testid="stNumberInput"] input:focus,
div[data-testid="stTextInput"] input:focus,
textarea:focus {
    border-color: var(--aci-border-focus) !important;
    box-shadow: inset 0 1px 3px rgba(0,0,0,.25), 0 0 0 2px rgba(82,130,255,.18) !important;
}

div[data-testid="stNumberInput"] button {
    min-height: 2.35rem !important;
    height: 2.35rem !important;
    min-width: 2.1rem !important;
    padding: 0 0.32rem !important;
}

div[data-testid="stNumberInput"] {
    margin-bottom: 0.15rem !important;
}

/* Selectbox */
div[data-baseweb="select"] > div,
div[data-baseweb="select"] div,
div[data-baseweb="popover"] div {
    background-color: rgba(18, 18, 36, 0.98) !important;
    color: #f0f2ff !important;
    border-color: rgba(130, 170, 255, 0.30) !important;
}

/* Disabled text */
[aria-disabled="true"], [disabled],
.st-emotion-cache-1gulkj5,
.st-emotion-cache-10trblm {
    color: var(--aci-subtle) !important;
    opacity: 1 !important;
}


/* ── Typography base ── */
html, body, [class*="css"], [class*="st-emotion"] {
    font-family: "Inter", "Segoe UI", "Roboto", "Arial", sans-serif !important;
}

h1, h2, h3, h4 {
    letter-spacing: -0.30px !important;
    color: var(--aci-text) !important;
    font-weight: 800 !important;
}

p, li { font-weight: 420; }

/* ── Button transitions ── */
.stButton > button,
.stDownloadButton > button {
    min-height: 40px !important;
    border-radius: var(--r-sm) !important;
    font-weight: 800 !important;
    letter-spacing: .20px !important;
    transition: transform .12s ease, box-shadow .15s ease, background .15s ease, border-color .15s ease !important;
}
.stButton > button:hover,
.stDownloadButton > button:hover {
    transform: translateY(-2px);
}
.stButton > button:active,
.stDownloadButton > button:active {
    transform: translateY(0) scale(.98) !important;
}

/* ── Tabs ── */
button[role="tab"] {
    border-radius: 10px 10px 0 0 !important;
    padding: 9px 14px !important;
    font-size: 13px !important;
}
button[role="tab"][aria-selected="true"] {
    background: rgba(82,130,255,.13) !important;
    box-shadow: inset 0 -2px 0 var(--aci-blue) !important;
}

/* ── DataFrame ── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--aci-border) !important;
    border-radius: var(--r-md) !important;
    overflow: hidden !important;
    background: rgba(14,14,26,.65) !important;
    box-shadow: var(--aci-shadow-sm) !important;
}

/* ── Metric text ── */
[data-testid="stMetric"] label,
[data-testid="stMetric"] label *,
[data-testid="stMetricValue"],
[data-testid="stMetricValue"] * {
    color: var(--aci-text) !important;
}

/* ── SVG shell (geometry diagram) ── */
.aci-svg-shell svg,
.aci-svg-shell svg text {
    font-family: "Inter", "Segoe UI", "Arial", sans-serif !important;
    text-rendering: geometricPrecision;
}
.aci-svg-shell svg {
    display: block;
    max-width: 100%;
    filter: drop-shadow(0 10px 20px rgba(0,0,0,.28));
}
.aci-svg-shell {
    background: linear-gradient(135deg, rgba(18, 18, 32, .97), rgba(12, 12, 22, .98));
    border: 1px solid rgba(130,170,255,.22);
    border-radius: var(--r-lg);
    padding: 12px;
    box-shadow: var(--aci-shadow-md);
    margin: 8px 0 14px 0;
}

/* ── Mini note / soft card ── */
.aci-mini-note {
    color: var(--aci-muted);
    font-size: 12px;
    line-height: 1.58;
}
.aci-soft-card {
    background: rgba(18, 18, 36, .60);
    border: 1px solid rgba(130,170,255,.14);
    border-radius: var(--r-md);
    padding: 12px 14px;
    margin: 6px 0;
}

/* ── Expanders ── */
[data-testid="stExpander"] summary,
[data-testid="stExpander"] summary *,
button[role="tab"],
button[role="tab"] * {
    color: var(--aci-text) !important;
    opacity: 1 !important;
}
[data-testid="stExpander"] summary,
[data-testid="stExpander"] summary * {
    white-space: normal !important;
    line-height: 1.45 !important;
}

/* ── Section title / subtitle ── */
.aci-section-title {
    display: block !important;
    width: 100% !important;
    margin: 16px 0 8px 0 !important;
    padding: 10px 16px !important;
    border-radius: var(--r-sm) !important;
    background: linear-gradient(135deg, rgba(82,130,255,.16) 0%, rgba(77,201,240,.06) 100%) !important;
    border: 1px solid rgba(82,130,255,.30) !important;
    color: var(--aci-text) !important;
    font-size: 15.5px !important;
    font-weight: 900 !important;
    line-height: 1.45 !important;
    letter-spacing: .05px !important;
}
.aci-section-subtitle {
    display: block !important;
    margin: 14px 0 8px 0 !important;
    color: #c6d6f4 !important;
    font-size: 13.5px !important;
    font-weight: 800 !important;
    line-height: 1.4 !important;
}

/* Avoid global font override breaking Streamlit Material Symbols. */
[data-testid="stIconMaterial"], [data-testid="stIconMaterial"] *,
span[class*="material"], i[class*="material"] {
    font-family: "Material Symbols Rounded", "Material Symbols Outlined", "Material Icons", sans-serif !important;
    font-weight: normal !important;
    font-style: normal !important;
    line-height: 1 !important;
    letter-spacing: normal !important;
    text-transform: none !important;
    white-space: nowrap !important;
    direction: ltr !important;
    font-feature-settings: "liga" !important;
}


/* ==========================================================================
   FIX12ZA — LIVE APP CSS PATCH
   Important: FIX12Z was accidentally inserted into the HTML report style block.
   This block is inserted inside inject_custom_css(), so it affects the live UI.
   Calculation logic is unchanged.
   ========================================================================== */

/* ── Expander: Material-style surface container ── */
[data-testid="stExpander"] {
    background: transparent !important;
}
[data-testid="stExpander"] details {
    background: transparent !important;
    border: 0 !important;
}
[data-testid="stExpander"] summary,
[data-testid="stExpander"] details > summary,
[data-testid="stExpander"] details[open] > summary,
[data-testid="stExpander"] summary:hover,
[data-testid="stExpander"] summary:focus,
[data-testid="stExpander"] summary:focus-visible,
[data-testid="stExpander"] summary:active {
    background: linear-gradient(135deg, rgba(28,28,52,.98) 0%, rgba(18,18,34,.99) 100%) !important;
    color: var(--aci-text) !important;
    -webkit-text-fill-color: var(--aci-text) !important;
    border: 1px solid rgba(82,130,255,.38) !important;
    border-radius: var(--r-md) !important;
    min-height: 46px !important;
    padding: 10px 16px !important;
    box-shadow: var(--aci-shadow-sm), inset 0 1px 0 rgba(255,255,255,.03) !important;
    transition: border-color .15s, box-shadow .15s !important;
}
[data-testid="stExpander"] summary:hover {
    border-color: rgba(82,130,255,.60) !important;
    box-shadow: var(--aci-shadow-md) !important;
}
[data-testid="stExpander"] summary *,
[data-testid="stExpander"] details > summary *,
[data-testid="stExpander"] details[open] > summary *,
[data-testid="stExpander"] summary:hover *,
[data-testid="stExpander"] summary:focus *,
[data-testid="stExpander"] summary:active * {
    color: var(--aci-text) !important;
    -webkit-text-fill-color: var(--aci-text) !important;
    text-shadow: none !important;
    font-weight: 650 !important;
}
[data-testid="stExpander"] summary svg,
[data-testid="stExpander"] summary svg *,
[data-testid="stExpander"] details > summary svg,
[data-testid="stExpander"] details > summary svg * {
    color: var(--aci-blue) !important;
    fill: var(--aci-blue) !important;
    stroke: var(--aci-blue) !important;
    opacity: 1 !important;
}
[data-testid="stExpander"] [data-testid="stExpanderDetails"],
[data-testid="stExpander"] details > div,
[data-testid="stExpander"] details[open] > div {
    background: rgba(20, 20, 36, .65) !important;
    color: var(--aci-text) !important;
    border: 1px solid rgba(82,130,255,.18) !important;
    border-top: 0 !important;
    border-radius: 0 0 var(--r-md) var(--r-md) !important;
}

/* ── Tables ── */
[data-testid="stMarkdownContainer"] table,
[data-testid="stMarkdownContainer"] thead,
[data-testid="stMarkdownContainer"] tbody,
[data-testid="stMarkdownContainer"] tr,
[data-testid="stMarkdownContainer"] th,
[data-testid="stMarkdownContainer"] td,
[data-testid="stTable"] table,
[data-testid="stTable"] thead,
[data-testid="stTable"] tbody,
[data-testid="stTable"] tr,
[data-testid="stTable"] th,
[data-testid="stTable"] td {
    background-color: rgba(18, 18, 34, .90) !important;
    color: var(--aci-text) !important;
    -webkit-text-fill-color: var(--aci-text) !important;
    border-color: rgba(82,130,255,.18) !important;
}
[data-testid="stMarkdownContainer"] th,
[data-testid="stTable"] th {
    background-color: rgba(28, 28, 52, .98) !important;
    color: var(--aci-blue) !important;
    -webkit-text-fill-color: var(--aci-blue) !important;
    font-weight: 700 !important;
    letter-spacing: .25px !important;
}
[data-testid="stMarkdownContainer"] td,
[data-testid="stTable"] td {
    font-weight: 430 !important;
}
[data-testid="stMarkdownContainer"] span,
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] div,
[data-testid="stMarkdownContainer"] li {
    background-color: transparent !important;
}
[data-testid="stMarkdownContainer"] code {
    background: rgba(14, 14, 28, .95) !important;
    color: var(--aci-cyan) !important;
    border: 1px solid rgba(82,130,255,.22) !important;
    border-radius: 5px !important;
    padding: 1px 5px !important;
}

/* ── DataFrame grid ── */
[data-testid="stDataFrame"],
[data-testid="stDataFrame"] * {
    color: var(--aci-text) !important;
    -webkit-text-fill-color: var(--aci-text) !important;
}
[data-testid="stDataFrame"] [role="gridcell"],
[data-testid="stDataFrame"] [role="columnheader"],
[data-testid="stDataFrame"] [role="rowheader"] {
    background: rgba(18, 18, 34, .88) !important;
    color: var(--aci-text) !important;
    -webkit-text-fill-color: var(--aci-text) !important;
}

/* ── Input stepper buttons ── */
div[data-baseweb="input"] > div,
div[data-baseweb="select"] > div,
input, textarea {
    background-color: rgba(18, 18, 36, .92) !important;
    color: #f0f2ff !important;
    -webkit-text-fill-color: #f0f2ff !important;
    border-color: rgba(82,130,255,.32) !important;
}
button[aria-label="Increment"],
button[aria-label="Decrement"] {
    background: rgba(26, 26, 46, .96) !important;
    color: var(--aci-blue) !important;
    border-color: rgba(82,130,255,.28) !important;
    transition: background .12s !important;
}
button[aria-label="Increment"]:hover,
button[aria-label="Decrement"]:hover {
    background: rgba(36, 36, 62, .98) !important;
}


/* ── Geo-help code inline ── */
.geo-help code {
    background: rgba(14, 14, 28, .96) !important;
    color: var(--aci-cyan) !important;
    border: 1px solid rgba(82,130,255,.22) !important;
    border-radius: 6px !important;
    padding: 1px 5px !important;
    font-size: .86em !important;
}

/* ── Geometry card ── */
.geo-card {
    padding: 12px 14px !important;
}
.geo-mini-note {
    color: var(--aci-subtle) !important;
    font-size: 11px !important;
    line-height: 1.48 !important;
}
.geo-bar-count-title {
    color: var(--aci-text) !important;
    font-weight: 700 !important;
    font-size: 12px !important;
    margin-bottom: 3px !important;
}
.geo-bar-count-sub {
    color: var(--aci-subtle) !important;
    font-size: 10px !important;
    line-height: 1.35 !important;
    margin-bottom: 6px !important;
}
.geo-canvas-grid-title {
    font-size: 13px !important;
    font-weight: 700 !important;
    color: var(--aci-text) !important;
    margin: 0 0 2px 0 !important;
}
.geo-canvas-grid-subtitle {
    font-size: 10px !important;
    font-weight: 400 !important;
    color: var(--aci-subtle) !important;
    margin: 0 0 6px 0 !important;
}
.geo-card div[data-testid="stNumberInput"] label p {
    font-size: 11.5px !important;
    color: #c0cfec !important;
    font-weight: 600 !important;
}
.geo-card div[data-testid="stNumberInput"] {
    margin-bottom: 3px !important;
}
.geo-card div[data-baseweb="input"] > div {
    min-height: 32px !important;
    border-radius: var(--r-sm) !important;
}
.geo-card input {
    font-size: 12.5px !important;
    font-weight: 600 !important;
}
.geo-card button[aria-label="Increment"],
.geo-card button[aria-label="Decrement"] {
    min-width: 28px !important;
    width: 28px !important;
}
.geo-ktr-wrap div[data-testid="stNumberInput"] label p,
.geo-bar-control-wrap div[data-testid="stNumberInput"] label p {
    font-size: 11px !important;
    color: var(--aci-muted) !important;
}
.geo-ktr-wrap div[data-baseweb="input"] > div,
.geo-bar-control-wrap div[data-baseweb="input"] > div {
    min-height: 30px !important;
    border-radius: var(--r-sm) !important;
}
/* Result-card hover — Material ripple-like border glow */
.geo-result-card:hover {
    border-color: rgba(82,130,255,.48) !important;
    box-shadow: 0 4px 18px rgba(82,130,255,.18), 0 6px 20px rgba(0,0,0,.30) !important;
}

/* ── Scrollbar — thin Material style ── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: rgba(18,18,30,.80); }
::-webkit-scrollbar-thumb { background: rgba(82,130,255,.35); border-radius: 99px; }
::-webkit-scrollbar-thumb:hover { background: rgba(82,130,255,.58); }

/* ── Focus-visible ring ── */
*:focus-visible {
    outline: 2px solid rgba(82,130,255,.60) !important;
    outline-offset: 2px !important;
}

</style>
        """,
        unsafe_allow_html=True,
    )

COVER_MD = """
<div style='max-width:780px;margin:48px auto 0 auto;text-align:center;padding:0 20px;'>
    <div style='display:inline-flex;align-items:center;justify-content:center;width:64px;height:64px;border-radius:18px;margin-bottom:20px;background:linear-gradient(135deg,#3d7eff 0%,#4dc9f0 100%);font-weight:900;font-size:22px;color:#fff;box-shadow:0 6px 24px rgba(82,130,255,.50),inset 0 1px 0 rgba(255,255,255,.22);letter-spacing:-0.5px;'>ACI</div>
    <h1 style='font-size:36px;font-weight:900;letter-spacing:-0.5px;color:#e8eaff;margin:0 0 6px 0;line-height:1.12;'>ACI 318 &#8211; Chapter 25</h1>
    <p style='font-size:17px;color:#9a9ec8;margin:0 0 32px 0;font-weight:400;line-height:1.6;'>Development, Anchorage, and Splicing of Reinforcement</p>
    <div style='width:56px;height:3px;border-radius:2px;margin:0 auto 32px auto;background:linear-gradient(90deg,#3d7eff,#4dc9f0);'></div>
    <div style='display:inline-flex;gap:10px;flex-wrap:wrap;justify-content:center;margin-bottom:28px;'>
        <span style='padding:5px 13px;border-radius:999px;font-size:12px;font-weight:700;background:rgba(82,130,255,.13);border:1px solid rgba(82,130,255,.32);color:#b8d8ff;letter-spacing:.2px;'>Revision 00</span>
        <span style='padding:5px 13px;border-radius:999px;font-size:12px;font-weight:700;background:rgba(105,240,174,.10);border:1px solid rgba(105,240,174,.28);color:#a6f0c6;letter-spacing:.2px;'>ACI 318-2019 &amp; 2025</span>
        <span style='padding:5px 13px;border-radius:999px;font-size:12px;font-weight:700;background:rgba(77,201,240,.10);border:1px solid rgba(77,201,240,.26);color:#a0e8ff;letter-spacing:.2px;'>SI &amp; KGF-CM</span>
    </div>
    <div style='background:rgba(28,28,48,.85);border:1px solid rgba(82,130,255,.18);border-radius:18px;padding:22px 28px;text-align:left;box-shadow:0 8px 28px rgba(0,0,0,.42);'>
        <p style='color:#c8d4f0;font-size:14.5px;line-height:1.70;margin:0 0 12px 0;'>
            This interface provides calculation tools for <strong style="color:#e8eaff;">development length</strong>,
            <strong style="color:#e8eaff;">anchorage</strong>, and <strong style="color:#e8eaff;">lap splices</strong>
            of reinforcement in accordance with ACI 318-2019 and ACI 318-2025, Chapter 25.
        </p>
        <p style='color:#6a6e94;font-size:12.5px;margin:0;'>
            Prepared by <strong style="color:#9a9ec8;">M. H. Torabi</strong> &nbsp;&middot;&nbsp; RAJAN PETRO FARAYAND
        </p>
    </div>
</div>
"""


def build_aci_parameter_reference_md(edition: str, units: str = "SI") -> str:
    """Build edition-aware sidebar reference.

    IMPORTANT:
    - This reference reports the **numbers and rules implemented in this app**.
    - Governing checks and computed results remain in the Detailed Report.
    """
    is_2019 = bool(edition) and edition.startswith("ACI318-2019")
    tables = ACI_318_2019_TABLES if is_2019 else ACI_318_2025_TABLES

    t_tension = tables.get("25.4.2.5", {})
    t_comp = tables.get("25.4.9.3", {})
    t_hook = tables.get("25.4.3.2", {})

    ed_tag = "ACI 318-2019 (SI)" if is_2019 else "ACI 318-2025 (SI)"

    # ---- Tension modifiers (25.4.2.5) ----
    lam = t_tension.get("lambda", {})
    psi_g = t_tension.get("psi_g", {})
    psi_s = t_tension.get("psi_s", {})
    psi_t = t_tension.get("psi_t", {})
    psi_e = t_tension.get("psi_e", {})
    psi_te_cap = (t_tension.get("psi_te_cap", {}) or {}).get("value", None)

    # ---- Compression (25.4.9.3) ----
    psi_r_c = t_comp.get("psi_r", {})
    lam_c = t_comp.get("lambda", {})

    # ---- Hooks (25.4.3.2) ----
    lam_h = t_hook.get("lambda", {})
    psi_e_h = t_hook.get("psi_e", {})
    psi_s_h = t_hook.get("psi_s", {})
    psi_cc_h = t_hook.get("psi_cc", {})
    psi_r_h = t_hook.get("psi_r", {})

    md = []
    md.append(f"### ACI Chapter 25 – Parameter Selection Reference (edition-aware)")
    md.append(f"**Edition selected:** `{ed_tag}`  ")
    md.append("")
    md.append("> This section shows the **numeric values and decision rules as implemented in this app** (so the user does not need to open the code).  ")
    md.append("> Governing checks and computed results are reported in the **Detailed report**.")
    md.append("")

    # Material inputs
    md.append("#### 1) Material inputs (entered by user)")
    md.append("| Symbol | What to enter |")
    md.append("|---|---|")
    md.append("| **f′c** | Specified concrete compressive strength used for Ch.25 development/splice/hook calculations (project concrete spec). |")
    md.append("| **fy** | Specified yield strength of reinforcement used for Ch.25 calculations (project rebar grade/spec). |")
    md.append("| **db** | Nominal bar diameter from selected bar size (or user-entered diameter). |")
    md.append("")

    # Geometry inputs
    md.append("#### 2) Geometry inputs used in formulas")
    md.append("| Symbol | How the value is taken in this app |")
    md.append("|---|---|")
    md.append("| **cover** | Clear cover (user input). Use the **minimum** clear cover in the development region. |")
    md.append("| **s_clear / spacing** | Clear spacing between parallel longitudinal bars (user input). Use the **minimum** clear spacing in the development region. This is not the stirrup spacing. |")
    md.append("| **so** | Bar-to-nearest-surface distance used in `c_b` selection (user input where applicable). Use the governing (smallest) value. |")
    md.append("| **c_b** | Computed as the governing cover/spacing parameter derived from `cover`, `s_clear`, and `so`. The detailed report indicates the governing contributor. |")
    md.append("| **Ktr** | Computed transverse reinforcement index from `Atr`, `s_Ktr`, and `n` (if provided). |")
    md.append("")

    # Tension modifiers
    md.append("#### 3) Tension development modifiers (Table 25.4.2.5 – implemented values)")
    md.append("| Factor | Implemented numeric values | Selection rule (math) |")
    md.append("|---|---|---|")
    md.append(f"| **λ** | normal = **{lam.get('normal','?')}**, lightweight = **{lam.get('lightweight','?')}** | λ = 1.0 (normalweight), λ = 0.75 (lightweight). |")
    md.append(f"| **ψt** | top-cast = **{psi_t.get('top_cast','?')}**, other = **{psi_t.get('other','?')}** | ψt = 1.3 (top-cast), ψt = 1.0 (otherwise). |")
    md.append(f"| **ψe** | uncoated/zinc = **{psi_e.get('uncoated_or_zinc','?')}**, epoxy (good cover/spacing) = **{psi_e.get('epoxy_other','?')}**, epoxy (poor cover/spacing) = **{psi_e.get('epoxy_low_cover_or_spacing','?')}** | ψe = 1.0 (uncoated/zinc). If epoxy: ψe = 1.2 when cover ≥ 3·db OR spacing ≥ 6·db; else ψe = 1.5. |")
    md.append(f"| **ψs** | No.22+ = **{psi_s.get('no22_and_larger','?')}**, No.19− or wire = **{psi_s.get('no19_and_smaller_or_wire','?')}** | Selected by bar size group per ACI for tension development. |")
    md.append(f"| **ψg** | 280 = **{psi_g.get('grade_280','?')}**, 420 = **{psi_g.get('grade_420','?')}**, 550 = **{psi_g.get('grade_550','?')}**, 690 = **{psi_g.get('grade_690','?')}** | Selected by reinforcement grade corresponding to fy. |")
    if psi_te_cap is not None:
        md.append(f"| **cap** | **ψt·ψe ≤ {psi_te_cap}** | Enforced as (ψt·ψe) ≤ (cap) by reducing ψe when needed. |")
    md.append("")

    # Ktr formula (improved)
    md.append("#### 4) Ktr (transverse reinforcement index – implemented equation)")
    md.append("")
    md.append("The transverse reinforcement index $K_{tr}$ is calculated as:")
    md.append("")
    md.append("$$")
    md.append("K_{tr} = \\frac{40\\,A_{tr}}{s\\,n}")
    md.append("$$")
    md.append("*(SI units: $A_{tr}$ in mm², $s$ in mm)*")
    md.append("")
    md.append("| Symbol | Description | Input method in this app |")
    md.append("|---|---|---|")
    md.append("| **$A_{tr}$** | Total cross‑sectional area of all transverse reinforcement crossing the potential splitting plane within a spacing $s$ | User enters directly as `Atr` (mm²) |")
    md.append("| **$s_{Ktr}$** | Maximum center-to-center spacing of transverse reinforcement used in the Ktr expression within the development length | User enters directly as Ktr transverse spacing (mm). It is different from s_clear between longitudinal bars. |")
    md.append("| **$n$** | Number of longitudinal bars being developed along the splitting plane | User enters as `n` (developed bars) |")
    md.append("")
    md.append("The constant 40 has units of 1/mm and is applicable only in SI (MPa, mm).")
    md.append("")

    # Compression
    md.append("#### 5) Compression development modifiers (Table 25.4.9.3 – implemented values)")
    md.append("| Factor | Implemented numeric values | Selection rule (math) |")
    md.append("|---|---|---|")
    md.append(f"| **λ** | normal = **{lam_c.get('normal','?')}**, lightweight = **{lam_c.get('lightweight','?')}** | λ = 1.0 (normalweight), λ = 0.75 (lightweight). |")
    md.append(f"| **ψr** | meets confinement = **{psi_r_c.get('meets_confinement','?')}**, other = **{psi_r_c.get('other','?')}** | ψr = 0.75 if confinement qualifies; otherwise ψr = 1.0. |")
    md.append("")

    # Hooks
    md.append("#### 6) Standard hook development modifiers (Table 25.4.3.2 – implemented values)")
    md.append(f"- Hook equation constant **C** is edition-dependent in this app: 2019 uses **C=23**; 2025 uses **C=4.16** (corrected per ACI official confirmation, Jan 2026).")
    md.append("")
    md.append("| Factor | Implemented numeric values | Selection rule (math) |")
    md.append("|---|---|---|")
    md.append(f"| **λ** | normal = **{lam_h.get('normal','?')}**, lightweight = **{lam_h.get('lightweight','?')}** | λ = 1.0 (normalweight), λ = 0.75 (lightweight). |")
    # ψe hook
    if is_2019:
        md.append(f"| **ψe** | uncoated/zinc = **{psi_e_h.get('uncoated_or_zinc','?')}**, epoxy/dual = **{psi_e_h.get('epoxy_or_dual','?')}** | ψe = 1.0 (uncoated/zinc), ψe = 1.2 (epoxy/dual). |")
        md.append(f"| **ψr** | meets confinement = **{psi_r_h.get('meets_confinement','?')}**, other = **{psi_r_h.get('other','?')}** | ψr selected by hook confinement per chosen edition. |")
        # 2019: ψo and ψc live in 25.4.3.2 for this app
        psi_o = t_hook.get("psi_o", {})
        psi_c = t_hook.get("psi_c", {})
        md.append(f"| **ψo** | ends in confined core / adequate side cover = **{psi_o.get('ends_in_core_or_adequate_side_cover','?')}**, other = **{psi_o.get('other','?')}** | Selected by `hook_end_in_core` / cover condition input. |")
        md.append(f"| **ψc** | meets cover condition = **{psi_c.get('meets_cover','?')}**, other = **{psi_c.get('other','?')}** | Implemented as `ψc=0.7` when (90° hook AND cover_ok), else 1.0. |")
    else:
        md.append(f"| **ψe** | uncoated/zinc = **{psi_e_h.get('uncoated_or_zinc','?')}**, epoxy/dual = **{psi_e_h.get('epoxy_or_dual','?')}** | Selected by `epoxy` input. |")
        # ψs group mapping
        md.append(f"| **ψs** | ≤No.29 = **{psi_s_h.get('no29_and_smaller','?')}**, No.32–36 = **{psi_s_h.get('no32_no36','?')}**, No.43 = **{psi_s_h.get('no43','?')}**, No.57 = **{psi_s_h.get('no57','?')}** | Selected by bar size group derived from `db`. |")
        md.append(f"| **ψcc** | meets cover = **{psi_cc_h.get('meets_cover','?')}**, other = **{psi_cc_h.get('other','?')}** | Implemented as `ψcc=0.7` when (90° hook AND cover_ok), else 1.0. |")
        md.append(f"| **ψr** | meets tie confinement = **{psi_r_h.get('meets_tie_confinement','?')}**, other = **{psi_r_h.get('other','?')}** | Selected by tie confinement condition input (spacing criterion reported by the app when applicable). |")
    md.append("")
    md.append("#### 7) Output symbols (computed)")
    md.append("| Symbol | Meaning |")
    md.append("|---|---|")
    md.append("| **ℓd** | Tension development length (straight bar). |")
    md.append("| **ℓdc** | Compression development length. |")
    md.append("| **ℓdh** | Standard hook development length. |")
    md.append("| **ℓs** | Lap splice length derived from development length and splice class selection. |")

    return "\n".join(md)







# =========================================================
# Chapter 25 — Diff-only panel (2019 vs 2025, SI)
# Textual + Parametric (NO calculations)
# =========================================================
def render_ch25_diff_only_panel_si():
    st.markdown("## ACI 318 Chapter 25 — Differences (SI)")
    st.caption("Comparison mode: ACI 318-2019 vs ACI 318-2025 | Differences only")

    st.markdown("### §25.4.2 — Development of deformed bars in tension (General)")
    st.markdown(
        "- **2019:** coefficient = 23; ψ-factors limited to ψt, ψe\n"
        "- **2025:** coefficient = 21; ψ-factors explicitly include **ψs** and **ψcc**\n"
        "- **Type:** numeric + parametric\n"
        "- **Note:** current app calculations intentionally keep ld identical; difference reported here only."
    )

    st.markdown("### §25.4.3 — Standard hooks in tension")
    st.markdown(
        "- **2019:** ψe, ψr, ψo\n"
        "- **2025:** ψe, ψr, ψo **+ ψs**\n"
        "- **Type:** parametric"
    )

    st.markdown("### §25.4.9 — Development of bars in compression")
    st.markdown(
        "- Minimum ℓdc = **200 mm** in both editions\n"
        "- **Type:** clarification (no numeric change)"
    )

    st.markdown("### §25.5 — Lap splices")
    st.markdown(
        "- No numeric change\n"
        "- **Type:** conceptual clarification"
    )

    st.markdown("### Detailing reference")
    st.markdown(
        "- **2019:** SP-66\n"
        "- **2025:** MNL-66\n"
        "- **Type:** reference update"
    )

    st.info("Only code-level differences are shown. Identical items are suppressed.")


# ---------------------------------------------------------------------------
# PHASE 2 — CENTRAL QA / WARNING PANEL
# UI/report support only. These helpers do not modify engineering formulas.
# ---------------------------------------------------------------------------

def add_qa_message(
    messages: List[Dict[str, Any]],
    level: str,
    code: str,
    message: str,
    source: str = "Calculation engine metadata",
    value: Any = None,
    action: str = "",
) -> None:
    """Append one normalized engineering QA message."""
    messages.append({
        "level": str(level).upper(),
        "code": str(code),
        "message": str(message),
        "source": str(source),
        "value": value,
        "action": str(action),
    })


def build_engineering_qa_messages(
    calc_inputs: Dict[str, Any],
    df_res: Optional[pd.DataFrame],
    units: str,
    edition: str,
) -> List[Dict[str, Any]]:
    """
    Build a central QA/warning list from the same calculation snapshot used
    by screen outputs and exports. This is a display/QA layer only.

    Important:
    - This function must not recalculate ACI design results.
    - It only reads already-computed values such as cb, Ktr, beta, and result tables.
    """
    messages: List[Dict[str, Any]] = []

    ci = calc_inputs or {}
    unit_len = "mm" if str(units).upper() == "SI" else "display length unit"

    add_qa_message(
        messages,
        "INFO",
        "ENGINE_VERSION",
        f"Engine/UI version: {globals().get('ENGINE_VERSION', 'unknown')}",
        "Application metadata",
    )
    add_qa_message(
        messages,
        "INFO",
        "CODE_EDITION",
        f"Selected code edition: {edition}",
        "User-selected input",
    )

    # Basic input sanity checks based on the locked calculation snapshot.
    numeric_checks = [
        ("fc_mpa", "Concrete strength f'c"),
        ("fy_mpa", "Reinforcement yield strength fy"),
        ("cover_mm", "Clear cover"),
        ("spacing_mm", "s_clear (clear bar spacing)"),
        ("so_mm", "so"),
    ]
    for key, label in numeric_checks:
        val = ci.get(key)
        try:
            fval = float(val)
            if fval <= 0:
                add_qa_message(
                    messages,
                    "CRITICAL",
                    f"INVALID_{key.upper()}",
                    f"{label} is not positive.",
                    "Input validation",
                    val,
                    "Correct the input before using design output.",
                )
        except Exception:
            add_qa_message(
                messages,
                "WARNING",
                f"MISSING_{key.upper()}",
                f"{label} is not available in the calculation snapshot.",
                "Input snapshot",
                val,
                "Recalculate and confirm the input snapshot.",
            )

    # c_b / governing geometry source.
    cb_val = ci.get("cb_mm")
    cb_gov = ci.get("cb_gov")
    if cb_val is not None:
        add_qa_message(
            messages,
            "INFO",
            "CB_SOURCE",
            f"c_b used by engine = {float(cb_val):.2f} {unit_len}; governing geometry source = {cb_gov or 'not stored'}.",
            "Calculation snapshot",
            cb_val,
        )

    # Ktr information.
    ktr_val = ci.get("Ktr_base")
    if ktr_val is not None:
        try:
            ktr_float = float(ktr_val)
            if abs(ktr_float) < 1e-12:
                add_qa_message(
                    messages,
                    "INFO",
                    "KTR_ZERO",
                    "Ktr stored by the engine is zero or not credited.",
                    "Calculation snapshot",
                    ktr_float,
                )
            else:
                add_qa_message(
                    messages,
                    "OK",
                    "KTR_ACTIVE",
                    f"Ktr stored by the engine = {ktr_float:.4g}.",
                    "Calculation snapshot",
                    ktr_float,
                )
        except Exception:
            add_qa_message(
                messages,
                "WARNING",
                "KTR_UNREADABLE",
                "Ktr value exists but could not be parsed.",
                "Calculation snapshot",
                ktr_val,
            )

    # beta cap status from the engine. No independent recalculation here.
    beta_raw = ci.get("beta_raw")
    beta_eff = ci.get("beta_eff")
    beta_is_capped = ci.get("beta_is_capped")
    if beta_raw is not None or beta_eff is not None:
        try:
            msg = f"β raw = {float(beta_raw):.3f}, β used = {float(beta_eff):.3f}."
        except Exception:
            msg = f"β raw = {beta_raw}, β used = {beta_eff}."
        if bool(beta_is_capped):
            add_qa_message(
                messages,
                "WARNING",
                "BETA_CAPPED",
                msg + " The engine has capped β.",
                "Calculation snapshot / engine rule",
                beta_eff,
                "Review cover, spacing, and transverse reinforcement assumptions.",
            )
        else:
            add_qa_message(
                messages,
                "OK",
                "BETA_NOT_CAPPED",
                msg + " β cap is not active.",
                "Calculation snapshot / engine rule",
                beta_eff,
            )

    # Modifiers / inputs that materially affect development length.
    if bool(ci.get("top_bar")):
        add_qa_message(
            messages,
            "INFO",
            "TOP_BAR_ACTIVE",
            "Top-bar condition is active in the calculation snapshot.",
            "User-selected input / engine modifier",
        )
    if bool(ci.get("epoxy")):
        add_qa_message(
            messages,
            "INFO",
            "EPOXY_ACTIVE",
            "Epoxy-coated reinforcement condition is active in the calculation snapshot.",
            "User-selected input / engine modifier",
        )

    lam_val = ci.get("lambda")
    try:
        if lam_val is not None and float(lam_val) < 1.0:
            add_qa_message(
                messages,
                "INFO",
                "LIGHTWEIGHT_ACTIVE",
                f"λ = {float(lam_val):.2f}; lightweight concrete factor is active.",
                "User-selected input / engine modifier",
                lam_val,
            )
    except Exception:
        pass

    # Result-table based QA. This reads screen results only.
    if df_res is None or getattr(df_res, "empty", True):
        add_qa_message(
            messages,
            "WARNING",
            "NO_RESULTS_TABLE",
            "No result table is available for QA review.",
            "Streamlit session state",
            None,
            "Run Calculate / Update.",
        )
    else:
        add_qa_message(
            messages,
            "OK",
            "RESULTS_TABLE_AVAILABLE",
            f"Results table available with {len(df_res)} row(s) and {len(df_res.columns)} column(s).",
            "Streamlit session state",
        )

        # Detect stored capped flags if present in result columns.
        for col in df_res.columns:
            if "cap @ 2.5" in str(col).lower():
                try:
                    if any(str(v).strip().lower() in ("yes", "true", "1") for v in df_res[col].tolist()):
                        add_qa_message(
                            messages,
                            "WARNING",
                            "BETA_CAP_TABLE_FLAG",
                            "At least one result row indicates β cap @ 2.5 is active.",
                            "Results table",
                        )
                except Exception:
                    pass

        # Hook/splice status columns, if the existing engine already created them.
        for status_col in ("Hook status", "Splice status"):
            if status_col in df_res.columns:
                bad_values = []
                for v in df_res[status_col].tolist():
                    sv = str(v).strip()
                    if sv and sv.upper() not in ("OK", "PASS", "-", "N/A", "NA"):
                        bad_values.append(sv)
                if bad_values:
                    add_qa_message(
                        messages,
                        "WARNING",
                        status_col.upper().replace(" ", "_"),
                        f"{status_col} contains non-OK values: {', '.join(sorted(set(bad_values))[:5])}.",
                        "Results table",
                        ", ".join(sorted(set(bad_values))[:5]),
                        "Review the related detailed output.",
                    )
                else:
                    add_qa_message(
                        messages,
                        "OK",
                        status_col.upper().replace(" ", "_"),
                        f"{status_col} entries are OK or not controlling.",
                        "Results table",
                    )

    if str(edition).strip() in ("ACI 318-25", "ACI318-2025", "ACI 318-2025"):
        add_qa_message(
            messages,
            "INFO",
            "ACI_2025_REVIEW_NOTE",
            "ACI 318-25 edition is selected. Any edition difference shown later must come from the verified calculation branch, not hardcoded comparison values.",
            "QA discipline note",
        )

    return messages


def qa_level_badge_html(level: str) -> str:
    """Return a compact HTML badge for QA severity."""
    lev = str(level).upper()
    if lev == "CRITICAL":
        cls = "crit"
    elif lev == "WARNING":
        cls = "warn"
    elif lev == "OK":
        cls = "ok"
    else:
        cls = "info"
    return f'<span class="aci-chip {cls}" style="margin:0;">{lev}</span>'


def render_engineering_qa_panel(messages: List[Dict[str, Any]]) -> None:
    """Render central QA messages in a professional review panel."""
    if not messages:
        st.info("No QA messages are available yet. Run Calculate / Update.")
        return

    counts = {}
    for m in messages:
        counts[str(m.get("level", "INFO")).upper()] = counts.get(str(m.get("level", "INFO")).upper(), 0) + 1

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("OK", counts.get("OK", 0))
    c2.metric("INFO", counts.get("INFO", 0))
    c3.metric("WARNING", counts.get("WARNING", 0))
    c4.metric("CRITICAL", counts.get("CRITICAL", 0))

    rows = []
    for m in messages:
        rows.append({
            "Level": str(m.get("level", "")),
            "Code": str(m.get("code", "")),
            "Message": str(m.get("message", "")),
            "Source": str(m.get("source", "")),
            "Value": "" if m.get("value", None) is None else str(m.get("value")),
            "Action": str(m.get("action", "")),
        })

    df_qa = pd.DataFrame(rows)
    st.dataframe(df_qa, width="stretch", hide_index=True)

    if counts.get("CRITICAL", 0) > 0:
        st.error("CRITICAL QA messages exist. Treat the affected design output as not valid until resolved.")
    elif counts.get("WARNING", 0) > 0:
        st.warning("WARNING messages exist. Review governing limits and assumptions before using the output.")
    else:
        st.success("No critical/warning QA messages detected in this run.")



def render_svg_component(svg_html: str, height: int = 460) -> None:
    """
    Render SVG/HTML through Streamlit components with a CAD-like local toolbar.

    UI-only scope:
    - zoom in / zoom out
    - fit/reset
    - dim/normal display
    - no engineering calculation is performed here
    """
    components.html(
        f"""
        <html>
        <head>
        <style>
            html, body {{
                margin: 0;
                padding: 0;
                background: transparent;
                overflow: hidden;
                font-family: Segoe UI, Inter, Arial, sans-serif;
            }}
            .viewer {{
                width: 100%;
                height: {max(height - 8, 260)}px;
                box-sizing: border-box;
                background: linear-gradient(135deg, rgba(4,12,22,.92), rgba(8,23,38,.92));
                border: 1px solid rgba(120,170,220,.18);
                border-radius: 18px;
                overflow: hidden;
                position: relative;
            }}
            .toolbar {{
                position: absolute;
                top: 10px;
                right: 12px;
                z-index: 20;
                display: flex;
                gap: 6px;
                padding: 6px;
                border-radius: 999px;
                background: rgba(3, 12, 22, .76);
                border: 1px solid rgba(120,170,220,.30);
                box-shadow: 0 10px 28px rgba(0,0,0,.25);
                backdrop-filter: blur(8px);
            }}
            .toolbar button {{
                min-width: 32px;
                height: 30px;
                border-radius: 999px;
                border: 1px solid rgba(40,213,255,.38);
                background: rgba(12,34,55,.96);
                color: #eaf3ff;
                cursor: pointer;
                font-weight: 900;
                line-height: 1;
                padding: 0 9px;
                transition: transform .12s ease, background .12s ease, border-color .12s ease;
            }}
            .toolbar button:hover {{
                transform: translateY(-1px);
                background: rgba(28,134,255,.72);
                border-color: rgba(40,213,255,.85);
            }}
            .zoom-readout {{
                min-width: 48px;
                text-align: center;
                color: #b8cbe2;
                font-size: 12px;
                font-weight: 800;
                align-self: center;
                padding: 0 4px;
            }}
            .stage-wrap {{
                width: 100%;
                height: 100%;
                overflow: auto;
                box-sizing: border-box;
                padding: 0;
            }}
            .stage {{
                transform-origin: top left;
                transition: transform .18s ease, filter .18s ease;
                min-width: 100%;
            }}
            .viewer.dim .stage {{
                filter: brightness(.82) contrast(1.08);
            }}
            .aci-svg-shell {{
                box-sizing: border-box;
                width: 100%;
                background: linear-gradient(135deg, rgba(7,20,34,.96), rgba(5,14,25,.96));
                border: 1px solid rgba(120,170,220,.26);
                border-radius: 18px;
                padding: 6px;
                box-shadow: 0 18px 46px rgba(0,0,0,.30);
            }}
            svg text {{
                font-family: Segoe UI, Inter, Arial, sans-serif;
                paint-order: stroke;
                stroke: rgba(0,0,0,.18);
                stroke-width: .45px;
            }}
            svg line, svg path, svg rect, svg circle {{
                vector-effect: non-scaling-stroke;
            }}
        </style>
        </head>
        <body>
            <div id="viewer" class="viewer">
                <div class="toolbar" title="Local SVG viewer controls">
                    <button onclick="zoomOut()" title="Zoom out">−</button>
                    <span id="zr" class="zoom-readout">100%</span>
                    <button onclick="zoomIn()" title="Zoom in">+</button>
                    <button onclick="fitView()" title="Fit / reset">Fit</button>
                    <button onclick="toggleDim()" title="Dim / normal">◐</button>
                </div>
                <div id="wrap" class="stage-wrap">
                    <div id="stage" class="stage">{svg_html}</div>
                </div>
            </div>

            <script>
                let scale = 1.0;
                const stage = document.getElementById("stage");
                const readout = document.getElementById("zr");
                const viewer = document.getElementById("viewer");

                function applyZoom() {{
                    stage.style.transform = "scale(" + scale + ")";
                    readout.textContent = Math.round(scale * 100) + "%";
                }}

                function zoomIn() {{
                    scale = Math.min(2.0, scale + 0.10);
                    applyZoom();
                }}

                function zoomOut() {{
                    scale = Math.max(0.60, scale - 0.10);
                    applyZoom();
                }}

                function fitView() {{
                    scale = 1.0;
                    applyZoom();
                    document.getElementById("wrap").scrollTop = 0;
                    document.getElementById("wrap").scrollLeft = 0;
                }}

                function toggleDim() {{
                    viewer.classList.toggle("dim");
                }}

                applyZoom();
            </script>
        </body>
        </html>
        """,
        height=height,
        scrolling=False,
    )

# ---------------------------------------------------------------------------
# PHASE 3 — DYNAMIC SVG GEOMETRY ENGINE
# Visualization layer only. These functions do not perform ACI design
# calculations; they read the existing calculation snapshot and result table.
# ---------------------------------------------------------------------------

def _num_from_any(value: Any, default: float = 0.0) -> float:
    try:
        if value is None:
            return default
        if isinstance(value, (int, float)):
            if math.isnan(float(value)):
                return default
            return float(value)
        s = str(value).replace(",", "")
        m = re.search(r"[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?", s)
        return float(m.group(0)) if m else default
    except Exception:
        return default


def _first_result_value(df_res: Optional[pd.DataFrame], key_tokens: List[str], default: float = 0.0) -> float:
    # Reads from the existing result table only; avoids a second calculation path.
    try:
        if df_res is None or getattr(df_res, "empty", True):
            return default
        row = df_res.iloc[0].to_dict()
        for col, val in row.items():
            name = str(col)
            if all(tok in name for tok in key_tokens):
                return _num_from_any(val, default)
    except Exception:
        pass
    return default


def build_live_geometry_data(
    calc_inputs: Dict[str, Any],
    df_res: Optional[pd.DataFrame],
    units: str,
) -> Dict[str, Any]:
    # Display-only geometry object from the locked calculation snapshot.
    ci = calc_inputs or {}
    unit_tag = "mm" if str(units).upper() == "SI" else "cm"

    db = _num_from_any(ci.get("db_mm"), 20.0)
    cover = _num_from_any(ci.get("cover_mm"), 40.0)
    spacing = _num_from_any(ci.get("spacing_mm"), 100.0)
    so = _num_from_any(ci.get("so_mm"), spacing)
    beam_b_mm = max(1.0, _num_from_any(ci.get("beam_b_mm"), 400.0))
    beam_h_mm = max(1.0, _num_from_any(ci.get("beam_h_mm"), 600.0))
    stirrup_db_mm = max(1.0, _num_from_any(ci.get("stirrup_db_mm"), 8.0))
    cb = _num_from_any(ci.get("cb_mm"), min(cover, spacing / 2.0, so / 2.0))
    beta_raw = _num_from_any(ci.get("beta_raw"), 0.0)
    beta_eff = _num_from_any(ci.get("beta_eff"), beta_raw)
    ktr = _num_from_any(ci.get("Ktr_base"), 0.0)

    ld_t = _first_result_value(df_res, ["ℓd_T"], default=0.0)
    if ld_t <= 0:
        ld_t = _first_result_value(df_res, ["ℓd", "T"], default=0.0)

    # FIX12C: display-only; read both simplified/general ld values
    # from the existing result table. No ACI value is recalculated here.
    ld_t_simp = _first_result_value(df_res, ["ℓd_T_Simp"], default=0.0)
    if ld_t_simp <= 0:
        ld_t_simp = _first_result_value(df_res, ["Simp_ℓd_T"], default=0.0)
    if ld_t_simp <= 0:
        ld_t_simp = _first_result_value(df_res, ["Simp", "ℓd", "T"], default=0.0)

    ld_t_gen = _first_result_value(df_res, ["ℓd_T_Gen"], default=0.0)
    if ld_t_gen <= 0:
        ld_t_gen = _first_result_value(df_res, ["Gen_ℓd_T"], default=0.0)
    if ld_t_gen <= 0:
        ld_t_gen = _first_result_value(df_res, ["Gen", "ℓd", "T"], default=0.0)

    # Bottom bar ld (both methods) — for the 4-line dimension layout
    ld_b_simp = _first_result_value(df_res, ["ℓd_B", "Simp"], default=0.0)
    if ld_b_simp <= 0:
        ld_b_simp = _first_result_value(df_res, ["Simp_ℓd_B"], default=0.0)
    ld_b_gen = _first_result_value(df_res, ["ℓd_B", "Gen"], default=0.0)
    if ld_b_gen <= 0:
        ld_b_gen = _first_result_value(df_res, ["Gen_ℓd_B"], default=0.0)

    ldc = _first_result_value(df_res, ["ℓdc"], default=0.0)
    ldh = _first_result_value(df_res, ["ℓdh"], default=0.0)
    ls_t = _first_result_value(df_res, ["ℓs_T"], default=0.0)

    return {
        "unit_tag": unit_tag,
        "db": db,
        "n_long": int(ci.get("n_long", 4) or 4),
        "cover": cover,
        "spacing": spacing,
        "beam_b_mm": beam_b_mm,
        "beam_h_mm": beam_h_mm,
        "stirrup_db_mm": stirrup_db_mm,
        "s_tr": _num_from_any(ci.get("s_tr_mm"), spacing),
        "so": so,
        "cb": cb,
        "cb_gov": str(ci.get("cb_gov", "not stored")),
        "beta_raw": beta_raw,
        "beta_eff": beta_eff,
        "beta_is_capped": bool(ci.get("beta_is_capped")),
        "ktr": ktr,
        "ld_t": ld_t,
        "ld_t_simp": ld_t_simp,
        "ld_t_gen": ld_t_gen,
        "ld_b_simp": ld_b_simp,
        "ld_b_gen": ld_b_gen,
        "ldc": ldc,
        "ldh": ldh,
        "ls_t": ls_t,
        "epoxy": bool(ci.get("epoxy")),
        "top_bar": bool(ci.get("top_bar")),
        "lambda": ci.get("lambda", ci.get("lam", 1.0)),
    }



def _beam_layout_points(top_count: int, bottom_count: int, left_count: int, right_count: int, x_left: float, x_right: float, y_top: float, y_bottom: float) -> List[Tuple[float, float, str]]:
    # Display-only. Counts are per face. Top/bottom include corners; side duplicates are merged.
    top_count = max(2, int(top_count or 2)); bottom_count = max(2, int(bottom_count or 2))
    left_count = max(2, int(left_count or 2)); right_count = max(2, int(right_count or 2))
    pts: Dict[Tuple[int, int], Tuple[float, float, str]] = {}
    def add_pt(x: float, y: float, tag: str) -> None:
        pts[(round(x), round(y))] = (x, y, tag)
    for i in range(top_count):
        add_pt(x_left + (x_right - x_left) * i / max(1, top_count - 1), y_top, "top")
    for i in range(bottom_count):
        add_pt(x_left + (x_right - x_left) * i / max(1, bottom_count - 1), y_bottom, "bottom")
    for i in range(left_count):
        add_pt(x_left, y_top + (y_bottom - y_top) * i / max(1, left_count - 1), "left")
    for i in range(right_count):
        add_pt(x_right, y_top + (y_bottom - y_top) * i / max(1, right_count - 1), "right")
    return list(pts.values())


def _layout_defaults_from_n_long(n_long: int) -> Tuple[int, int, int, int]:
    # Four corners first; extras: top, bottom, then sides.
    n = max(4, int(n_long or 4)); top = bottom = left = right = 2; extra = n - 4
    add_top = min(extra, 4); top += add_top; extra -= add_top
    add_bottom = min(extra, 4); bottom += add_bottom; extra -= add_bottom
    add_left = extra // 2; add_right = extra - add_left
    left += add_left; right += add_right
    return top, bottom, left, right


def normalize_beam_layout(layout: Optional[Dict[str, Any]]) -> Dict[str, int]:
    def face_count(value: Any) -> int:
        try:
            return max(2, min(20, int(value)))
        except Exception:
            return 2

    source = layout or {}
    return {
        "top": face_count(source.get("top")),
        "bottom": face_count(source.get("bottom")),
        "left": face_count(source.get("left")),
        "right": face_count(source.get("right")),
    }


def svg_dynamic_ld_geometry(
    gd: Dict[str, Any],
    show_cover: bool = True,
    show_cb: bool = True,
    show_spacing: bool = True,
    show_ktr: bool = True,
    show_ld: bool = True,
) -> str:
    # UI-only longitudinal beam section.
    # FIX14B: re-based from FIX14A. Do not import the later FIX15 canvas.
    # Top/bottom longitudinal bars in the elevation now follow the cross-section
    # Top/Bottom face counts. This keeps the section/elevation visually aligned
    # while leaving all ACI calculations untouched.
    db = max(1.0, _num_from_any(gd.get("db"), 20.0))
    cover = _num_from_any(gd.get("cover"), 40.0)
    spacing = _num_from_any(gd.get("spacing"), 100.0)  # clear spacing between longitudinal bars (ACI 25.2 display)
    s_tr = _num_from_any(gd.get("s_tr"), spacing)  # transverse reinforcement spacing used for Ktr display
    so = _num_from_any(gd.get("so"), spacing)
    cb = _num_from_any(gd.get("cb"), min(cover, spacing / 2.0, so / 2.0))
    beta = _num_from_any(gd.get("beta_eff"), 0.0)
    ld = _num_from_any(gd.get("ld_t"), 0.0)
    ld_simp   = _num_from_any(gd.get("ld_t_simp"), 0.0)
    ld_gen    = _num_from_any(gd.get("ld_t_gen"),  0.0)
    ld_b_simp = _num_from_any(gd.get("ld_b_simp"), 0.0)
    ld_b_gen  = _num_from_any(gd.get("ld_b_gen"),  0.0)
    unit = str(gd.get("unit_tag", "mm"))
    epoxy = bool(gd.get("epoxy"))
    capped = bool(gd.get("beta_is_capped"))
    layout = normalize_beam_layout(gd.get("beam_layout") or {})
    top_count = layout["top"]
    bottom_count = layout["bottom"]
    left_count = layout["left"]
    right_count = layout["right"]
    # Display rule: the longitudinal side/elevation view is governed by the side-face
    # longitudinal bars visible along the beam depth. Use the larger of Left/Right
    # face counts, so the elevation remains synchronized with the cross-section.
    elevation_count = max(left_count, right_count)

    rebar_color1 = "#2367d8" if epoxy else "#202020"
    rebar_color2 = "#6fd6ff" if epoxy else "#bdbdbd"
    beam_b_mm = max(1.0, _num_from_any(gd.get("beam_b_mm"), 400.0))
    beam_h_mm = max(1.0, _num_from_any(gd.get("beam_h_mm"), 600.0))
    stirrup_db_mm = max(1.0, _num_from_any(gd.get("stirrup_db_mm"), 8.0))
    beam_x, beam_w = 78, 675
    beam_center_y = 221.0
    max_conc_w, max_conc_h = 312.0, 190.0
    visual_depth_scale = min(max_conc_w / beam_b_mm, max_conc_h / beam_h_mm)
    # The cross-section SVG is drawn in a narrower Streamlit column and a 520-wide
    # viewBox, while this elevation uses a 980-wide viewBox in a 1.70x column.
    # Compensate internal height so the final on-screen concrete depths align.
    screen_scale_comp = 980.0 / (520.0 * 1.70)
    beam_h = beam_h_mm * visual_depth_scale * screen_scale_comp
    beam_y = beam_center_y - beam_h / 2.0
    depth_scale = beam_h / beam_h_mm
    cover_px = cover * depth_scale
    stirrup_db_px = stirrup_db_mm * depth_scale
    db_px = db * depth_scale
    bar_h = max(1.5, db_px)
    bar_center_offset = (cover + stirrup_db_mm + db / 2.0) * depth_scale
    bar_center_offset = min(max(bar_center_offset, bar_h / 2.0 + 4.0), max(bar_h / 2.0 + 4.0, beam_h / 2.0 - bar_h / 2.0 - 4.0))
    stirrup_center_offset = cover_px + stirrup_db_px / 2.0
    stirrup_center_offset = min(max(stirrup_center_offset, 3.0), max(3.0, beam_h / 2.0 - 3.0))
    cage_x, cage_y, cage_w, cage_h = 98, beam_y + stirrup_center_offset, 635, max(1.0, beam_h - 2.0 * stirrup_center_offset)
    top_bar_y = beam_y + bar_center_offset
    bottom_bar_y = beam_y + beam_h - bar_center_offset
    stirrup_display_w = max(1.5, min(10.0, stirrup_db_px))
    stirrup_xs = [122, 258, 410, 562, 704]

    def reinforcement_elevation() -> str:
        def count_markers(x0: float, y: float, count: int) -> str:
            shown = min(max(0, int(count)), 8)
            marks = "".join(
                f'<circle cx="{x0 + i * 10.0:.1f}" cy="{y:.1f}" r="3.2" fill="{rebar_color1}" stroke="#dbe7f2" stroke-width="0.7"/>\n'
                for i in range(shown)
            )
            if int(count) > shown:
                marks += f'<text x="{x0 + shown * 10.0 + 5.0:.1f}" y="{y+3.5:.1f}" fill="#dbe7f2" font-size="9" font-weight="650">+{int(count)-shown}</text>\n'
            return marks

        def level_positions(y0: float, y1: float, count: int) -> List[float]:
            if count <= 1:
                return [(y0 + y1) / 2.0]
            return [y0 + (y1 - y0) * i / (count - 1) for i in range(count)]

        def bar_layer(y: float, label: str = "", count: int = 0, label_y: float = 0.0, opacity: float = 1.0) -> str:
            bar_x = beam_x + 14.0
            bar_w = beam_w - 28.0
            bar_y = y - bar_h / 2.0
            bar_rx = min(bar_h / 2.0, 3.0)
            out = (
                f'<rect x="{bar_x:.1f}" y="{bar_y:.1f}" width="{bar_w:.1f}" height="{bar_h:.2f}" '
                f'rx="{bar_rx:.2f}" fill="url(#rebarGrad)" fill-opacity="{opacity:.2f}" stroke="{rebar_color1}" stroke-width="0.45"/>\n'
            )
            if label:
                out += f'<text x="{beam_x + beam_w - 8:.1f}" y="{label_y+3.5:.1f}" text-anchor="end" fill="#f4f8ff" font-size="11" font-weight="700">{label} = {int(count)} bars</text>\n'
            return out

        levels = level_positions(top_bar_y, bottom_bar_y, elevation_count)
        out = ""
        for i, y in enumerate(levels):
            is_top = i == 0
            is_bottom = i == len(levels) - 1
            if is_top:
                out += bar_layer(y, "Top", top_count, y + bar_h / 2.0 + 17.0, 1.0)
            elif is_bottom:
                out += bar_layer(y, "Bottom", bottom_count, y - bar_h / 2.0 - 13.0, 1.0)
            else:
                out += bar_layer(y, opacity=0.68)
        out += f'<text x="{beam_x + 18:.1f}" y="{beam_y + beam_h - 11.0:.1f}" fill="#dbe7f2" font-size="10" font-weight="650" opacity=".82">Elevation levels = max(L,R) = {int(elevation_count)}</text>\n'
        return out

    stirrups = ""
    for x in stirrup_xs:
        stirrups += f'<rect x="{x-stirrup_display_w/2:.1f}" y="{cage_y:.1f}" width="{stirrup_display_w:.1f}" height="{cage_h:.1f}" rx="2.0" fill="none" stroke="#dbe7f2" stroke-width="0.9" opacity=".48"/><line x1="{x-stirrup_display_w/2+1.3:.1f}" y1="{cage_y+6:.1f}" x2="{x+stirrup_display_w/2-1.3:.1f}" y2="{cage_y+6:.1f}" stroke="#dbe7f2" stroke-width="0.7" opacity=".42"/><line x1="{x-stirrup_display_w/2+1.3:.1f}" y1="{cage_y+cage_h-6:.1f}" x2="{x+stirrup_display_w/2-1.3:.1f}" y2="{cage_y+cage_h-6:.1f}" stroke="#dbe7f2" stroke-width="0.7" opacity=".42"/>\n'
    cover_block = ""
    if show_cover:
        top_cover_box_y = max(90.0, beam_y - 8.0)
        bottom_cover_box_y = min(326.0, beam_y + beam_h - 36.0)
        cover_block = f'<line x1="824" y1="{beam_y:.1f}" x2="824" y2="{beam_y+cover_px:.1f}" stroke="#eaf3ff" stroke-width="1.0" marker-start="url(#arrowW)" marker-end="url(#arrowW)"/><rect x="842" y="{top_cover_box_y:.1f}" width="108" height="42" rx="8" fill="#081a2b" stroke="#345a78" opacity=".96"/><text x="896" y="{top_cover_box_y+19:.1f}" text-anchor="middle" fill="#eaf3ff" font-size="12" font-weight="600">Top cover</text><text x="896" y="{top_cover_box_y+35:.1f}" text-anchor="middle" fill="#9ed7ff" font-size="12" font-weight="600">{cover:.0f} {unit}</text><line x1="824" y1="{beam_y+beam_h-cover_px:.1f}" x2="824" y2="{beam_y+beam_h:.1f}" stroke="#eaf3ff" stroke-width="1.0" marker-start="url(#arrowW)" marker-end="url(#arrowW)"/><rect x="842" y="{bottom_cover_box_y:.1f}" width="108" height="42" rx="8" fill="#081a2b" stroke="#345a78" opacity=".96"/><text x="896" y="{bottom_cover_box_y+19:.1f}" text-anchor="middle" fill="#eaf3ff" font-size="12" font-weight="600">Bottom cover</text><text x="896" y="{bottom_cover_box_y+35:.1f}" text-anchor="middle" fill="#9ed7ff" font-size="12" font-weight="600">{cover:.0f} {unit}</text>'
    cb_block = f'<line x1="330" y1="{(top_bar_y+bottom_bar_y)/2:.1f}" x2="530" y2="{(top_bar_y+bottom_bar_y)/2:.1f}" stroke="#28d5ff" stroke-width="2.0" marker-start="url(#arrowB)" marker-end="url(#arrowB)"/><text x="430" y="{(top_bar_y+bottom_bar_y)/2-14:.1f}" fill="#28d5ff" font-size="12" font-weight="650" text-anchor="middle">cb = {cb:.0f} {unit}</text>' if show_cb else ""
    spacing_y = min(367.0, beam_y + beam_h + 17.0)
    spacing_block = f'<line x1="122" y1="{spacing_y:.1f}" x2="258" y2="{spacing_y:.1f}" stroke="#1c86ff" stroke-width="1.5" marker-start="url(#arrowB)" marker-end="url(#arrowB)"/><text x="190" y="{spacing_y+14:.1f}" text-anchor="middle" fill="#70c6ff" font-size="12" font-weight="600">stirrup spacing s = {s_tr:.0f} {unit}</text><text x="500" y="{spacing_y+14:.1f}" text-anchor="middle" fill="#9fb2c8" font-size="11.5" font-weight="500">s_clear = {spacing:.0f} {unit} · so = {so:.0f} {unit}</text>' if show_spacing else ""
    ktr_w = max(130, min(350, 110 + beta * 90))
    ktr_block = f'<rect x="{430-ktr_w/2:.1f}" y="{cage_y+18}" width="{ktr_w:.1f}" height="{cage_h-36}" rx="8" fill="#34e36f" fill-opacity="0.10" stroke="#34e36f" stroke-width="1.2" stroke-dasharray="6 5"/><text x="430" y="{cage_y+cage_h-35}" text-anchor="middle" fill="#34e36f" font-size="13" font-weight="700">Ktr zone · β = {beta:.2f}</text>' if show_ktr else ""
    ld_block = ""
    if show_ld:
        _xl, _xr = 154.0, 706.0
        _span    = _xr - _xl  # 552 px
        _max_ref = max(
            ld_simp   if ld_simp   > 0 else (ld   if ld   > 0 else 1.0),
            ld_b_simp if ld_b_simp > 0 else 0.0,
            1.0,
        )

        def _cx(v: float) -> float:
            """Scale ld value to x-pixel endpoint."""
            return min(_xr, _xl + _span * v / _max_ref)

        def _optB(ya: float, x2: float, color: str, mk_end: str, mk_start: str,
                  label: str, mask_w: float = 244.0) -> str:
            """Option-B inline masked dimension line."""
            mid = (_xl + x2) / 2.0
            mx  = max(_xl + 14.0, min(mid - mask_w / 2.0, max(_xl + 15.0, x2 - mask_w - 6.0)))
            rx  = mx + mask_w
            mcx = mx + mask_w / 2.0
            return (
                f'<line x1="{_xl:.1f}" y1="{ya:.1f}" x2="{mx:.1f}" y2="{ya:.1f}" stroke="{color}" stroke-width="2.0" marker-start="{mk_start}"/>'
                f'<line x1="{rx:.1f}" y1="{ya:.1f}" x2="{x2:.1f}" y2="{ya:.1f}" stroke="{color}" stroke-width="2.0" marker-end="{mk_end}"/>'
                f'<rect x="{mx:.1f}" y="{ya-11.0:.1f}" width="{mask_w:.0f}" height="22" rx="7" fill="#071727" stroke="{color}" stroke-width="1.2"/>'
                f'<text x="{mcx:.1f}" y="{ya+5.0:.1f}" text-anchor="middle" fill="{color}" font-size="12.5" font-weight="800">{label}</text>'
            )

        has_top_both = ld_simp > 0 and ld_gen > 0
        has_bot_simp = ld_b_simp > 0
        has_bot_gen  = ld_b_gen  > 0
        has_bot_any  = has_bot_simp or has_bot_gen

        if has_top_both and has_bot_any:
            # ── FULL 4-LINE OPTION-B LAYOUT ──────────────────────────
            x2_st = _xr
            x2_gt = _cx(ld_gen)
            x2_sb = _cx(ld_b_simp) if has_bot_simp else 0.0
            x2_gb = _cx(ld_b_gen)  if has_bot_gen  else 0.0
            # TOP zone label
            ld_block += f'<text x="{_xl:.0f}" y="14" fill="rgba(255,77,77,.50)" font-size="9" font-weight="800" letter-spacing="1.5">TOP BARS</text>'
            # Row 1 — Simplified Top  y=30
            ld_block += _optB(30.0, x2_st, "#ff4d4d", "url(#arrowR)", "url(#arrowRl)", f"ℓd Simplified · Top = {ld_simp:.0f} {unit}")
            ld_block += (
                f'<line x1="{_xl:.1f}" y1="30" x2="{_xl:.1f}" y2="{beam_y:.1f}" stroke="#ff4d4d" stroke-width="1.0" stroke-dasharray="5,4" opacity=".50"/>'
                f'<line x1="{x2_st:.1f}" y1="30" x2="{x2_st:.1f}" y2="{beam_y:.1f}" stroke="#ff4d4d" stroke-width="1.0" stroke-dasharray="5,4" opacity=".50"/>'
            )
            # Row 2 — General Top  y=58
            ld_block += _optB(58.0, x2_gt, "#ffb020", "url(#arrowO)", "url(#arrowOl)", f"ℓd General · Top = {ld_gen:.0f} {unit}", 236.0)
            ld_block += f'<line x1="{x2_gt:.1f}" y1="58" x2="{x2_gt:.1f}" y2="{beam_y:.1f}" stroke="#ffb020" stroke-width="1.0" stroke-dasharray="5,4" opacity=".45"/>'
            # BOTTOM rows (below spacing_block)
            ext_y0 = spacing_y + 14.0
            if has_bot_simp:
                y3 = spacing_y + 30.0
                ld_block += (
                    f'<line x1="{_xl:.1f}" y1="{ext_y0:.1f}" x2="{_xl:.1f}" y2="{y3:.1f}" stroke="#ff4d4d" stroke-width="1.0" stroke-dasharray="5,4" opacity=".50"/>'
                    f'<line x1="{x2_sb:.1f}" y1="{ext_y0:.1f}" x2="{x2_sb:.1f}" y2="{y3:.1f}" stroke="#ff4d4d" stroke-width="1.0" stroke-dasharray="5,4" opacity=".50"/>'
                )
                ld_block += _optB(y3, x2_sb, "#ff4d4d", "url(#arrowR)", "url(#arrowRl)", f"ℓd Simplified · Bot = {ld_b_simp:.0f} {unit}")
            if has_bot_gen:
                y4 = spacing_y + (58.0 if has_bot_simp else 30.0)
                ld_block += f'<line x1="{x2_gb:.1f}" y1="{ext_y0:.1f}" x2="{x2_gb:.1f}" y2="{y4:.1f}" stroke="#ffb020" stroke-width="1.0" stroke-dasharray="5,4" opacity=".45"/>'
                ld_block += _optB(y4, x2_gb, "#ffb020", "url(#arrowO)", "url(#arrowOl)", f"ℓd General · Bot = {ld_b_gen:.0f} {unit}", 236.0)
                ld_block += f'<text x="{_xl:.0f}" y="{y4+20.0:.1f}" fill="rgba(255,176,32,.50)" font-size="9" font-weight="800" letter-spacing="1.5">BOTTOM BARS</text>'

        elif has_top_both:
            # ── 2-LINE TOP ONLY ───────────────────────────────────────
            x2_gt = _cx(ld_gen)
            ld_block += _optB(30.0, _xr, "#ff4d4d", "url(#arrowR)", "url(#arrowRl)", f"ℓd Simplified = {ld_simp:.0f} {unit}")
            ld_block += (
                f'<line x1="{_xl:.1f}" y1="30" x2="{_xl:.1f}" y2="{beam_y:.1f}" stroke="#ff4d4d" stroke-width="1.0" stroke-dasharray="5,4" opacity=".50"/>'
                f'<line x1="{_xr:.1f}" y1="30" x2="{_xr:.1f}" y2="{beam_y:.1f}" stroke="#ff4d4d" stroke-width="1.0" stroke-dasharray="5,4" opacity=".50"/>'
            )
            ld_block += _optB(58.0, x2_gt, "#ffb020", "url(#arrowO)", "url(#arrowOl)", f"ℓd General = {ld_gen:.0f} {unit}", 236.0)
            ld_block += f'<line x1="{x2_gt:.1f}" y1="58" x2="{x2_gt:.1f}" y2="{beam_y:.1f}" stroke="#ffb020" stroke-width="1.0" stroke-dasharray="5,4" opacity=".45"/>'

        else:
            # ── SINGLE VALUE FALLBACK ─────────────────────────────────
            lv = ld_simp if ld_simp > 0 else ld
            ld_block = (
                f'<line x1="154" y1="76" x2="706" y2="76" stroke="#ff4d4d" stroke-width="2.0" marker-start="url(#arrowRl)" marker-end="url(#arrowR)"/>'
                f'<line x1="154" y1="76" x2="154" y2="{beam_y:.1f}" stroke="#ff4d4d" stroke-dasharray="5 5"/>'
                f'<line x1="706" y1="76" x2="706" y2="{beam_y:.1f}" stroke="#ff4d4d" stroke-dasharray="5 5"/>'
                f'<text x="430" y="68" text-anchor="middle" fill="#ff4d4d" font-size="22" font-weight="700">ℓd = {lv:.0f} {unit}</text>'
            )
    cap_note = '<text x="690" y="467" fill="#ffb020" font-size="12" font-weight="650">WARNING: β capped by engine</text>' if capped else ''
    return f'''
<div class="aci-svg-shell"><svg viewBox="0 0 980 500" width="100%" height="500" xmlns="http://www.w3.org/2000/svg" role="img"><defs>
<pattern id="concPat" patternUnits="userSpaceOnUse" width="42" height="42"><rect width="42" height="42" fill="#c8bda9"/><circle cx="9" cy="12" r="1.9" fill="#756e64" opacity=".75"/><circle cx="27" cy="8" r="1.3" fill="#5f584f" opacity=".65"/><circle cx="34" cy="28" r="2.2" fill="#8a8072" opacity=".72"/><path d="M5,35 L14,31 L19,37" stroke="#675f56" stroke-width="1" opacity=".45"/></pattern>
<linearGradient id="rebarGrad" x1="0" x2="1"><stop offset="0" stop-color="{rebar_color1}"/><stop offset=".23" stop-color="{rebar_color2}"/><stop offset=".55" stop-color="{rebar_color1}"/><stop offset=".82" stop-color="{rebar_color2}"/><stop offset="1" stop-color="{rebar_color1}"/></linearGradient>
<marker id="arrowR"  markerWidth="8" markerHeight="8" refX="4"   refY="4"   orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="#ff4d4d"/></marker>
<marker id="arrowRl" markerWidth="8" markerHeight="8" refX="4"   refY="4"   orient="auto-start-reverse"><path d="M0,0 L8,4 L0,8 z" fill="#ff4d4d"/></marker>
<marker id="arrowO"  markerWidth="8" markerHeight="8" refX="4"   refY="4"   orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="#ffb020"/></marker>
<marker id="arrowOl" markerWidth="8" markerHeight="8" refX="4"   refY="4"   orient="auto-start-reverse"><path d="M0,0 L8,4 L0,8 z" fill="#ffb020"/></marker>
<marker id="arrowB"  markerWidth="8" markerHeight="8" refX="4"   refY="4"   orient="auto"><path d="M0,0 L8,4 L0,8 z" fill="#1c86ff"/></marker>
<marker id="arrowW"  markerWidth="7" markerHeight="7" refX="3.5" refY="3.5" orient="auto"><path d="M0,0 L7,3.5 L0,7 z" fill="#eaf3ff"/></marker>
</defs>
<rect x="6" y="6" width="968" height="488" rx="20" fill="#071727" stroke="#2a4764" stroke-width="1.2"/>
<rect x="18" y="28" width="248" height="26" rx="7" fill="#071727" fill-opacity="0.92" stroke="#82aaff" stroke-width="0.8"/>
<text x="30" y="46" fill="#e8eaff" font-size="15" font-weight="700">Longitudinal Beam Section</text><text x="742" y="42" fill="#34e36f" font-size="14" font-weight="700">LIVE UPDATE</text>
<rect x="{beam_x}" y="{beam_y}" width="{beam_w}" height="{beam_h}" rx="8" fill="url(#concPat)" stroke="#14171c" stroke-width="2"/>
<rect x="{cage_x}" y="{cage_y}" width="{cage_w}" height="{cage_h}" rx="6" fill="none" stroke="#eef6ff" stroke-width="1.0" opacity=".22"/>
{ld_block}{ktr_block}{stirrups}{reinforcement_elevation()}{cb_block}{spacing_block}{cover_block}
<rect x="34" y="464" width="285" height="18" rx="7" fill="#0b2034" stroke="#294761"/><text x="176" y="477" text-anchor="middle" fill="#eaf3ff" font-size="12" font-weight="650">db = {db:.0f} {unit} · tie db = {stirrup_db_mm:.0f} {unit}</text>
<rect x="335" y="454" width="440" height="30" rx="8" fill="#0b2034" stroke="#294761"/><text x="555" y="474" text-anchor="middle" fill="#eaf3ff" font-size="13" font-weight="650">cb source = {gd.get('cb_gov', 'not stored')} · Ktr = {_num_from_any(gd.get('ktr'),0.0):.2f}</text>{cap_note}
</svg></div>
'''

def svg_dynamic_cross_section(gd: Dict[str, Any], show_cover: bool = True, show_spacing: bool = True) -> str:
    # FIX12J: cross-section with closed stirrup outside all longitudinal bars.
    db = max(1.0, _num_from_any(gd.get("db"), 20.0))
    cover = _num_from_any(gd.get("cover"), 40.0)
    spacing = _num_from_any(gd.get("spacing"), 100.0)
    unit = str(gd.get("unit_tag", "mm"))
    epoxy = bool(gd.get("epoxy"))
    layout = normalize_beam_layout(gd.get("beam_layout") or {})
    top_n = layout["top"]
    bot_n = layout["bottom"]
    left_n = layout["left"]
    right_n = layout["right"]
    bar_fill = "#1c86ff" if epoxy else "#22262d"
    beam_b_mm = max(1.0, _num_from_any(gd.get("beam_b_mm"), 400.0))
    beam_h_mm = max(1.0, _num_from_any(gd.get("beam_h_mm"), 600.0))
    stirrup_db_mm = max(1.0, _num_from_any(gd.get("stirrup_db_mm"), 8.0))
    max_conc_w, max_conc_h = 312.0, 190.0
    display_scale = min(max_conc_w / beam_b_mm, max_conc_h / beam_h_mm)
    conc_w = beam_b_mm * display_scale
    conc_h = beam_h_mm * display_scale
    conc_x = 260.0 - conc_w / 2.0
    conc_y = 126.0 + (max_conc_h - conc_h) / 2.0
    cover_px = cover * display_scale
    stirrup_db_px = stirrup_db_mm * display_scale
    db_px = db * display_scale
    bar_r = max(1.5, db_px / 2.0)
    bar_center_offset = (cover + stirrup_db_mm + db / 2.0) * display_scale
    stirrup_outer_x, stirrup_outer_y = conc_x + cover_px, conc_y + cover_px
    stirrup_center_offset = cover_px + stirrup_db_px / 2.0
    stirrup_x, stirrup_y = conc_x + stirrup_center_offset, conc_y + stirrup_center_offset
    stirrup_w = max(1.0, conc_w - 2.0 * stirrup_center_offset)
    stirrup_h = max(1.0, conc_h - 2.0 * stirrup_center_offset)
    tie_stroke = max(1.0, stirrup_db_px)
    stirrup_inner_x, stirrup_inner_y = stirrup_x + stirrup_db_px / 2.0, stirrup_y + stirrup_db_px / 2.0
    stirrup_inner_w = max(1.0, stirrup_w - stirrup_db_px)
    stirrup_inner_h = max(1.0, stirrup_h - stirrup_db_px)
    bx0, bx1 = conc_x + bar_center_offset, conc_x + conc_w - bar_center_offset
    by0, by1 = conc_y + bar_center_offset, conc_y + conc_h - bar_center_offset
    pts: List[Tuple[float, float]] = []
    def add_pt(x: float, y: float) -> None:
        key = (round(float(x), 1), round(float(y), 1))
        if key not in pts:
            pts.append(key)
    def linspace(a: float, b: float, n: int) -> List[float]:
        if n <= 1:
            return [(a + b) / 2.0]
        return [a + (b - a) * i / (n - 1) for i in range(n)]
    for x in linspace(bx0, bx1, top_n): add_pt(x, by0)
    for x in linspace(bx0, bx1, bot_n): add_pt(x, by1)
    for y in linspace(by0, by1, left_n): add_pt(bx0, y)
    for y in linspace(by0, by1, right_n): add_pt(bx1, y)
    bars_svg = "".join([f'<circle cx="{x:.1f}" cy="{y:.1f}" r="{bar_r:.1f}" fill="{bar_fill}" stroke="#05070a" stroke-width="1.2"/>\n' for x, y in pts])
    clear_spacing_values = []
    if top_n > 1:
        clear_spacing_values.append((((bx1 - bx0) / (top_n - 1)) - db_px) / display_scale)
    if bot_n > 1:
        clear_spacing_values.append((((bx1 - bx0) / (bot_n - 1)) - db_px) / display_scale)
    if left_n > 1:
        clear_spacing_values.append((((by1 - by0) / (left_n - 1)) - db_px) / display_scale)
    if right_n > 1:
        clear_spacing_values.append((((by1 - by0) / (right_n - 1)) - db_px) / display_scale)
    drawn_clear_spacing = max(0.0, min(clear_spacing_values)) if clear_spacing_values else 0.0
    spacing_warning = drawn_clear_spacing + 0.5 < spacing
    cover_block = ""
    if show_cover:
        cover_block = f'<line x1="{conc_x:.1f}" y1="{stirrup_outer_y:.1f}" x2="{stirrup_outer_x:.1f}" y2="{stirrup_outer_y:.1f}" stroke="#34e36f" stroke-width="1.2" marker-start="url(#aG)" marker-end="url(#aG)"/><line x1="{stirrup_outer_x:.1f}" y1="{conc_y:.1f}" x2="{stirrup_outer_x:.1f}" y2="{stirrup_outer_y:.1f}" stroke="#34e36f" stroke-width="1.2" marker-start="url(#aG)" marker-end="url(#aG)"/><text x="260" y="{conc_y - 6.0:.1f}" text-anchor="middle" fill="#9cffaa" font-size="12" font-weight="700">cover = {cover:.0f} mm to outside of stirrup</text>'
    spacing_block = ""
    if show_spacing:
        spacing_block = f'<line x1="{bx0:.1f}" y1="{conc_y+conc_h+30:.1f}" x2="{bx1:.1f}" y2="{conc_y+conc_h+30:.1f}" stroke="#1c86ff" stroke-width="1.5" marker-start="url(#aB)" marker-end="url(#aB)"/>'
    clear_width_inside = max(0.0, beam_b_mm - 2.0 * (cover + stirrup_db_mm))
    clear_depth_inside = max(0.0, beam_h_mm - 2.0 * (cover + stirrup_db_mm))
    total_drawn_bars = len(pts)
    b_dim_y = min(342.0, conc_y + conc_h + 23.0)
    h_dim_x = min(486.0, conc_x + conc_w + 42.0)
    inside_w_y = stirrup_inner_y + stirrup_inner_h / 2.0
    inside_h_x = stirrup_inner_x + stirrup_inner_w / 2.0
    dimension_block = (
        f'<line x1="{conc_x:.1f}" y1="{b_dim_y:.1f}" x2="{conc_x + conc_w:.1f}" y2="{b_dim_y:.1f}" stroke="#dbe7f2" stroke-width="1.0" marker-start="url(#aD)" marker-end="url(#aD)"/>'
        f'<line x1="{conc_x:.1f}" y1="{b_dim_y-7:.1f}" x2="{conc_x:.1f}" y2="{b_dim_y+7:.1f}" stroke="#dbe7f2" stroke-width="0.8"/>'
        f'<line x1="{conc_x + conc_w:.1f}" y1="{b_dim_y-7:.1f}" x2="{conc_x + conc_w:.1f}" y2="{b_dim_y+7:.1f}" stroke="#dbe7f2" stroke-width="0.8"/>'
        f'<text x="260" y="{b_dim_y-5:.1f}" text-anchor="middle" fill="#dbe7f2" font-size="10.5" font-weight="650">b = {beam_b_mm:.0f} mm</text>'
        f'<line x1="{h_dim_x:.1f}" y1="{conc_y:.1f}" x2="{h_dim_x:.1f}" y2="{conc_y + conc_h:.1f}" stroke="#dbe7f2" stroke-width="1.0" marker-start="url(#aD)" marker-end="url(#aD)"/>'
        f'<line x1="{h_dim_x-7:.1f}" y1="{conc_y:.1f}" x2="{h_dim_x+7:.1f}" y2="{conc_y:.1f}" stroke="#dbe7f2" stroke-width="0.8"/>'
        f'<line x1="{h_dim_x-7:.1f}" y1="{conc_y + conc_h:.1f}" x2="{h_dim_x+7:.1f}" y2="{conc_y + conc_h:.1f}" stroke="#dbe7f2" stroke-width="0.8"/>'
        f'<text x="{h_dim_x+13:.1f}" y="{conc_y + conc_h/2.0:.1f}" text-anchor="middle" fill="#dbe7f2" font-size="10.5" font-weight="650" transform="rotate(90 {h_dim_x+13:.1f} {conc_y + conc_h/2.0:.1f})">h = {beam_h_mm:.0f} mm</text>'
        f'<line x1="{stirrup_inner_x:.1f}" y1="{inside_w_y:.1f}" x2="{stirrup_inner_x + stirrup_inner_w:.1f}" y2="{inside_w_y:.1f}" stroke="#7fb6e8" stroke-width="0.9" stroke-dasharray="3 3" marker-start="url(#aDimB)" marker-end="url(#aDimB)" opacity=".86"/>'
        f'<line x1="{inside_h_x:.1f}" y1="{stirrup_inner_y:.1f}" x2="{inside_h_x:.1f}" y2="{stirrup_inner_y + stirrup_inner_h:.1f}" stroke="#7fb6e8" stroke-width="0.9" stroke-dasharray="3 3" marker-start="url(#aDimB)" marker-end="url(#aDimB)" opacity=".86"/>'
    )
    side_label_offset = 42.0 if max(left_n, right_n) >= 4 else 34.0
    left_label_x = max(30.0, conc_x - side_label_offset)
    right_label_x = min(490.0, conc_x + conc_w + side_label_offset)
    group_block = (
        f'<text x="260" y="{max(77.0, conc_y-19.0):.1f}" text-anchor="middle" fill="#f4f8ff" font-size="11" font-weight="700">T face</text>'
        f'<text x="260" y="{conc_y+conc_h-6.0:.1f}" text-anchor="middle" fill="#203040" font-size="10.5" font-weight="800">B face</text>'
        f'<text x="{left_label_x:.1f}" y="{conc_y+conc_h/2.0:.1f}" text-anchor="middle" fill="#f4f8ff" font-size="11" font-weight="700" transform="rotate(-90 {left_label_x:.1f} {conc_y+conc_h/2.0:.1f})">L = {left_n}</text>'
        f'<text x="{right_label_x:.1f}" y="{conc_y+conc_h/2.0:.1f}" text-anchor="middle" fill="#f4f8ff" font-size="11" font-weight="700" transform="rotate(90 {right_label_x:.1f} {conc_y+conc_h/2.0:.1f})">R = {right_n}</text>'
    )
    face_marker_block = ""
    spacing_status = (
        "Warning: drawn clear spacing < input s_clear"
        if spacing_warning
        else "ACI cb uses input s_clear"
    )
    spacing_status_fill = "#ffcf5a" if spacing_warning else "#8fa4b8"
    legend_y = 360.0
    legend_h = 60.0 if spacing_warning else 52.0
    warning_line = (
        f'<text x="260" y="{legend_y+42:.1f}" text-anchor="middle" fill="{spacing_status_fill}" font-size="8.8" font-weight="700">{spacing_status}</text>'
        if spacing_warning
        else ""
    )
    count_y = legend_y + (56.0 if spacing_warning else 44.0)
    legend_block = (
        f'<rect x="72" y="{legend_y:.1f}" width="376" height="{legend_h:.1f}" rx="7" fill="#0b2034" stroke="#294761" stroke-width="1.0"/>'
        f'<text x="164" y="{legend_y+14:.1f}" text-anchor="middle" fill="#eaf3ff" font-size="9.2" font-weight="700">db = {db:.0f} mm / tie db = {stirrup_db_mm:.0f} mm</text>'
        f'<text x="356" y="{legend_y+14:.1f}" text-anchor="middle" fill="#d4e5f4" font-size="9.0" font-weight="650">inside tie faces = {clear_width_inside:.0f} x {clear_depth_inside:.0f} mm</text>'
        f'<text x="260" y="{legend_y+29:.1f}" text-anchor="middle" fill="#9fb2c8" font-size="9.0" font-weight="600">input s_clear = {spacing:.0f} mm / drawn clear = {drawn_clear_spacing:.0f} mm</text>'
        f'{warning_line}'
        f'<text x="260" y="{count_y:.1f}" text-anchor="middle" fill="#eaf3ff" font-size="9.1" font-weight="700">drawn count = {total_drawn_bars} bars / corners shared</text>'
    )
    return f'''
<div class="aci-svg-shell">
<svg viewBox="0 0 520 425" width="100%" height="425" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <pattern id="concCS" patternUnits="userSpaceOnUse" width="36" height="36"><rect width="36" height="36" fill="#c8bda9"/><circle cx="8" cy="10" r="1.4" fill="#6b6258" opacity=".7"/><circle cx="26" cy="20" r="1.8" fill="#887d70" opacity=".7"/><path d="M5,29 L14,25 L20,31" stroke="#675f56" stroke-width="1" opacity=".42"/></pattern>
    <linearGradient id="barCS" x1="0" x2="1"><stop offset="0" stop-color="{bar_fill}"/><stop offset=".50" stop-color="#d7f2ff"/><stop offset="1" stop-color="{bar_fill}"/></linearGradient>
    <filter id="shadowCS"><feDropShadow dx="0" dy="2" stdDeviation="2" flood-color="#000" flood-opacity=".45"/></filter>
    <marker id="aG" markerWidth="6" markerHeight="6" refX="3" refY="3" orient="auto"><path d="M0,0 L6,3 L0,6 z" fill="#34e36f"/></marker>
    <marker id="aB" markerWidth="6" markerHeight="6" refX="3" refY="3" orient="auto"><path d="M0,0 L6,3 L0,6 z" fill="#1c86ff"/></marker>
    <marker id="aD" markerWidth="7" markerHeight="7" refX="3.5" refY="3.5" orient="auto"><path d="M0,0 L7,3.5 L0,7 z" fill="#dbe7f2"/></marker>
    <marker id="aDimB" markerWidth="6" markerHeight="6" refX="3" refY="3" orient="auto"><path d="M0,0 L6,3 L0,6 z" fill="#7fb6e8"/></marker>
  </defs>
  <rect x="6" y="6" width="508" height="413" rx="20" fill="#071727" stroke="#2a4764"/>
  <text x="24" y="38" fill="#f4f8ff" font-size="19" font-weight="700">Beam Cross Section</text>
  <text x="24" y="59" fill="#9fb2c8" font-size="13">Closed stirrup surrounds bars; counts include corners.</text>
  <rect x="{conc_x}" y="{conc_y}" width="{conc_w}" height="{conc_h}" rx="10" fill="url(#concCS)" stroke="#12161b" stroke-width="2.0"/>
  <rect x="{stirrup_x}" y="{stirrup_y}" width="{stirrup_w}" height="{stirrup_h}" rx="6" fill="none" stroke="#e9f1f8" stroke-width="{tie_stroke:.1f}" filter="url(#shadowCS)" opacity=".98"/>
  <rect x="{stirrup_inner_x}" y="{stirrup_inner_y}" width="{stirrup_inner_w}" height="{stirrup_inner_h}" rx="5" fill="none" stroke="#8d969e" stroke-width="0.8" opacity=".75"/>
  <line x1="260" y1="{conc_y}" x2="260" y2="{conc_y+conc_h}" stroke="#72cbff" stroke-width="1" stroke-dasharray="4 5" opacity=".35"/>
  <line x1="{conc_x}" y1="{conc_y+conc_h/2}" x2="{conc_x+conc_w}" y2="{conc_y+conc_h/2}" stroke="#72cbff" stroke-width="1" stroke-dasharray="4 5" opacity=".35"/>
  {dimension_block}{bars_svg}{face_marker_block}{cover_block}{spacing_block}{group_block}{legend_block}
</svg>
</div>
'''

def svg_dynamic_hook_preview(gd: Dict[str, Any], angle: int = 90) -> str:
    db = max(8.0, _num_from_any(gd.get("db"), 20.0))
    ldh = _num_from_any(gd.get("ldh"), 0.0)
    unit = str(gd.get("unit_tag", "mm"))
    sw = max(18, min(30, db * 1.1))

    # FIX10A — UI-only correction:
    # The dashed bend-diameter guide must sit on the actual bend region.
    # It is a graphical guide only; all hook lengths and bend values remain
    # sourced from the existing calculation/result tables.
    if angle == 90:
        path = "M138 305 L138 138 Q138 88 188 88 L342 88"
        bend_cx, bend_cy, bend_r = 188, 138, 50
        bend_label_x, bend_label_y = 216, 126
    elif angle == 135:
        path = "M138 305 L138 198 Q138 152 176 124 L332 38"
        bend_cx, bend_cy, bend_r = 184, 166, 50
        bend_label_x, bend_label_y = 210, 157
    else:
        path = "M138 305 L138 138 Q138 88 188 88 L287 88 Q337 88 337 138 L337 305"
        bend_cx, bend_cy, bend_r = 287, 138, 50
        bend_label_x, bend_label_y = 287, 138

    label = f"ℓdh = {ldh:.0f} {unit}" if ldh > 0 else "ℓdh from result table"

    return f'''
<div class="aci-svg-shell hook-svg-shell">
<svg viewBox="0 0 420 360" width="100%" height="360" xmlns="http://www.w3.org/2000/svg">
  <defs>
    <linearGradient id="hookBar{angle}" x1="0" x2="1">
      <stop offset="0" stop-color="#242424"/><stop offset=".45" stop-color="#b7b7b7"/><stop offset="1" stop-color="#242424"/>
    </linearGradient>
    <marker id="aR{angle}" markerWidth="10" markerHeight="10" refX="5" refY="5" orient="auto"><path d="M0,0 L10,5 L0,10 z" fill="#ff4d4d"/></marker>
  </defs>
  <rect x="6" y="6" width="408" height="348" rx="18" fill="#071727" stroke="#2a4764"/>
  <text x="24" y="38" fill="#f4f8ff" font-size="18" font-weight="950">Hook {angle}° Preview</text>

  <circle cx="{bend_cx}" cy="{bend_cy}" r="{bend_r}" fill="rgba(52,227,111,.055)" stroke="#34e36f" stroke-width="1.7" stroke-dasharray="7 6"/>
  <text x="{bend_label_x}" y="{bend_label_y}" text-anchor="middle" fill="#34e36f" font-size="13" font-weight="900">Inside bend Ø</text>

  <path d="{path}" fill="none" stroke="url(#hookBar{angle})" stroke-width="{sw:.1f}" stroke-linecap="round" stroke-linejoin="round"/>
  <path d="{path}" fill="none" stroke="#f0f0f0" stroke-width="1.6" stroke-dasharray="12 10" opacity=".58"/>

  <line x1="72" y1="305" x2="72" y2="92" stroke="#ff4d4d" stroke-width="2" marker-start="url(#aR{angle})" marker-end="url(#aR{angle})"/>
  <text x="42" y="205" fill="#ff4d4d" font-size="16" font-weight="950" transform="rotate(-90 42 205)">{label}</text>
</svg>
</div>
'''


def render_live_geometry_panel(
    calc_inputs: Dict[str, Any],
    df_res: Optional[pd.DataFrame],
    qa_messages: Optional[List[Dict[str, Any]]],
    units: str,
    edition: str,
) -> None:
    """
    FIX12V — Geometry layout rebuilt from the stable FIX12U baseline.

    UI-only scope:
    - no ACI formula is modified
    - no result-table value is overwritten
    - bar-count controls are graphics-only except for the local Ktr preview inputs
    """
    st.markdown(
        """
<style>
/* ── PRO GEO PANEL ── professional engineering software geometry layout ── */

/* Panel header */
.geo-panel-header {
    display: flex; align-items: center; gap: 14px;
    background: linear-gradient(135deg, rgba(12,34,58,.96) 0%, rgba(7,20,37,.98) 100%);
    border: 1px solid rgba(40,180,255,.28);
    border-radius: 16px; padding: 14px 20px; margin-bottom: 14px;
    box-shadow: 0 8px 28px rgba(0,0,0,.32), inset 0 1px 0 rgba(255,255,255,.04);
}
.geo-panel-icon {
    width: 44px; height: 44px; border-radius: 12px; flex-shrink: 0;
    background: linear-gradient(135deg, #0f70ff 0%, #28d5ff 100%);
    display: flex; align-items: center; justify-content: center;
    font-size: 20px; box-shadow: 0 0 20px rgba(28,134,255,.38);
}
.geo-panel-title-block { flex: 1; }
.geo-panel-title { font-size: 17px; font-weight: 800; color: #f4f9ff; letter-spacing: -.1px; margin: 0 0 3px 0; }
.geo-panel-subtitle { font-size: 11.5px; color: #7fa8cc; line-height: 1.4; margin: 0; }
.geo-panel-badge {
    padding: 5px 11px; border-radius: 8px; font-size: 10.5px; font-weight: 800;
    letter-spacing: .4px; text-transform: uppercase;
    background: rgba(52,227,111,.12); color: #34e36f;
    border: 1px solid rgba(52,227,111,.30);
}

/* Section dividers / mini-titles */
.geo-mini-title {
    font-size: 10px; letter-spacing: .08em; text-transform: uppercase;
    color: #5aa8e0; font-weight: 800;
    margin: 14px 0 6px 0;
    display: flex; align-items: center; gap: 7px;
}
.geo-mini-title::after {
    content: ''; flex: 1; height: 1px;
    background: linear-gradient(to right, rgba(86,139,184,.30), transparent);
}

/* Card base */
.geo-card {
    border: 1px solid rgba(86,139,184,.22);
    border-radius: 12px;
    background: linear-gradient(160deg, rgba(10,28,48,.90) 0%, rgba(6,18,32,.94) 100%);
    padding: 12px 14px 10px 14px;
    margin-bottom: 10px;
    box-shadow: 0 4px 16px rgba(0,0,0,.22), inset 0 1px 0 rgba(255,255,255,.03);
}

/* KPI result strip */
.geo-result-strip {
    display: grid;
    grid-template-columns: repeat(6, minmax(100px,1fr));
    gap: 8px; margin: 6px 0 10px 0;
}
.geo-result-card {
    border: 1px solid rgba(86,139,184,.22);
    border-radius: 11px;
    background: linear-gradient(160deg, rgba(10,30,52,.90), rgba(6,18,34,.94));
    padding: 10px 12px 8px 12px;
    min-height: 70px;
    box-shadow: 0 4px 14px rgba(0,0,0,.20);
    transition: border-color .15s ease, box-shadow .15s ease;
    position: relative; overflow: hidden;
}
.geo-result-card::before {
    content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px;
    background: linear-gradient(to right, rgba(40,213,255,.45), rgba(28,134,255,.20));
    border-radius: 11px 11px 0 0;
}
.geo-result-card.hi::before { background: linear-gradient(to right, #ff4d4d, #ff8c00); }
.geo-result-card.ok::before { background: linear-gradient(to right, #34e36f, #00c4a0); }
.geo-result-card .k {
    font-size: 9.5px; color: #7fa8cc; font-weight: 800;
    text-transform: uppercase; letter-spacing: .06em;
}
.geo-result-card .v {
    font-size: 1.08rem; color: #f0f8ff; font-weight: 800;
    margin-top: 3px; letter-spacing: -.1px;
}
.geo-result-card .s { font-size: 9px; color: #5d82a0; margin-top: 2px; }

/* Canvas panel titles */
.geo-canvas-title {
    font-size: 13px; color: #dceeff; font-weight: 750;
    margin: 0 0 2px 0; letter-spacing: .02em;
}
.geo-canvas-sub { font-size: 10px; color: #5d82a0; margin-bottom: 6px; }

/* Help note */
.geo-help {
    font-size: 11px; line-height: 1.55; color: #7a9abb;
    border-left: 2px solid rgba(40,180,255,.30);
    padding-left: 10px; margin: 0 0 8px 0;
}
.geo-help code {
    background: rgba(8,22,36,.95) !important; color: #54d4ff !important;
    border: 1px solid rgba(84,207,255,.22) !important;
    border-radius: 4px !important; padding: 1px 5px !important;
    font-size: .88em !important;
}

/* Bar count controls */
.geo-bar-control-wrap {
    border: 1px solid rgba(86,139,184,.22);
    border-radius: 11px;
    background: linear-gradient(160deg, rgba(8,24,42,.88), rgba(5,16,28,.92));
    padding: 10px 12px 8px 12px; margin-top: 8px;
    box-shadow: 0 3px 12px rgba(0,0,0,.18);
}
.geo-bar-control-title {
    font-size: 11.5px; color: #dceeff; font-weight: 750; margin-bottom: 2px;
}
.geo-bar-control-note { font-size: 9.5px; color: #5d82a0; margin-bottom: 6px; }
.geo-bar-control-wrap div[data-testid="stNumberInput"] { margin-bottom: .14rem !important; }
.geo-bar-control-wrap div[data-testid="stNumberInput"] label { font-size: .72rem !important; }
.geo-bar-control-wrap div[data-testid="stNumberInput"] label * { font-size: .72rem !important; }
.geo-bar-control-wrap div[data-testid="stNumberInput"] input {
    min-height: 1.72rem !important; height: 1.72rem !important;
    font-size: .78rem !important;
}
.geo-bar-control-wrap div[data-testid="stNumberInput"] button {
    min-height: 1.72rem !important; height: 1.72rem !important;
    min-width: 1.35rem !important; padding: 0 .14rem !important;
}

/* Ktr control card */
.geo-ktr-wrap {
    border: 1px solid rgba(52,227,111,.18);
    border-radius: 11px;
    background: linear-gradient(160deg, rgba(8,26,14,.82), rgba(6,18,10,.90));
    padding: 10px 12px 8px 12px; margin-bottom: 0;
    box-shadow: 0 3px 12px rgba(0,0,0,.14);
}
.geo-ktr-title { font-size: 11.5px; color: #34e36f; font-weight: 750; margin-bottom: 2px; }
.geo-ktr-note { font-size: 9.5px; color: #3a8a52; margin-bottom: 6px; }

/* Display layer checkboxes */
.geo-layers-wrap {
    border: 1px solid rgba(86,139,184,.18);
    border-radius: 11px;
    background: rgba(8,22,38,.72);
    padding: 9px 12px 7px 12px; margin-bottom: 8px;
}
.geo-layers-title { font-size: 11px; color: #a0bdd4; font-weight: 750; margin-bottom: 4px; text-transform: uppercase; letter-spacing: .06em; }

div[data-testid="stCheckbox"] label { font-size: .80rem !important; }
div[data-testid="stCheckbox"] { margin-bottom: .08rem !important; }
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div { min-height: 2.10rem !important; font-size: .86rem !important; }

@media (max-width: 1400px) { .geo-result-strip { grid-template-columns: repeat(3, minmax(110px,1fr)); } }
@media (max-width: 900px)  { .geo-result-strip { grid-template-columns: repeat(2, minmax(110px,1fr)); } }
</style>
        """,
        unsafe_allow_html=True,
    )

    # ── Professional panel header ──
    st.markdown(
        "<div class='geo-panel-header'>"
        "<div class='geo-panel-icon'>⬡</div>"
        "<div class='geo-panel-title-block'>"
        "<div class='geo-panel-title'>Section Geometry Viewer</div>"
        "<div class='geo-panel-subtitle'>Live ACI 318 Ch.25 · Longitudinal section + cross-section · synchronized with calculation snapshot</div>"
        "</div>"
        "<div class='geo-panel-badge'>● LIVE</div>"
        "</div>",
        unsafe_allow_html=True,
    )

    main_list_for_geometry = st.session_state.get("main_list", []) or []
    selected_geo_db = shared_single_bar_selectbox(
        "Bar diameter for geometry (mm)",
        main_list_for_geometry,
        key="selected_bar_geometry_mm",
    )

    df_res_for_geometry = filter_df_res_for_selected_bar(df_res, selected_geo_db)
    calc_inputs_for_geometry = dict(calc_inputs or {})
    if selected_geo_db is not None:
        calc_inputs_for_geometry["db_mm"] = float(selected_geo_db)
    gd = build_live_geometry_data(calc_inputs_for_geometry, df_res_for_geometry, units)
    default_top, default_bottom, default_left, default_right = _layout_defaults_from_n_long(int(gd.get("n_long", 4) or 4))

    layout_sig = int(gd.get("n_long", 4) or 4)
    layout_keys = ("bars_top_face", "bars_bottom_face", "bars_left_face", "bars_right_face")
    if any(key not in st.session_state for key in layout_keys):
        st.session_state["bars_top_face"] = default_top
        st.session_state["bars_bottom_face"] = default_bottom
        st.session_state["bars_left_face"] = default_left
        st.session_state["bars_right_face"] = default_right
        st.session_state["geometry_layout_n_long_sig"] = layout_sig
    elif "geometry_layout_n_long_sig" not in st.session_state:
        st.session_state["geometry_layout_n_long_sig"] = layout_sig

    bars_top_face = int(st.session_state.get("bars_top_face", default_top))
    bars_bottom_face = int(st.session_state.get("bars_bottom_face", default_bottom))
    bars_left_face = int(st.session_state.get("bars_left_face", default_left))
    bars_right_face = int(st.session_state.get("bars_right_face", default_right))
    beam_layout = normalize_beam_layout({
        "top": bars_top_face,
        "bottom": bars_bottom_face,
        "left": bars_left_face,
        "right": bars_right_face,
    })
    bars_top_face = beam_layout["top"]
    bars_bottom_face = beam_layout["bottom"]
    bars_left_face = beam_layout["left"]
    bars_right_face = beam_layout["right"]
    st.session_state["beam_layout_graphics"] = beam_layout
    gd["beam_layout"] = beam_layout

    # ── Controls row: display layers + Ktr ──
    snap_now = st.session_state.get("calc_inputs", {}) or {}
    geo_Atr_default = float(snap_now.get("Atr_mm2", 0.0) or 0.0)
    geo_s_tr_disp_default = from_si_len_to_display(float(snap_now.get("s_tr_mm", 0.0) or 0.0), units)

    control_col, ktr_col = st.columns([1.0, 1.55], gap="medium")

    with control_col:
        st.markdown(
            "<div class='geo-help'>"
            "<b>Display discipline:</b> "
            "<code>s_clear</code> = clear spacing for c<sub>b</sub>/ψe; "
            "<code>s_Ktr</code> = transverse reinforcement spacing for K<sub>tr</sub>. "
            "Graphics are live from calculation snapshot — no ACI formulas are modified."
            "</div>",
            unsafe_allow_html=True,
        )
        st.markdown("<div class='geo-layers-wrap'>", unsafe_allow_html=True)
        st.markdown("<div class='geo-layers-title'>Canvas Overlays</div>", unsafe_allow_html=True)
        d1, d2, d3 = st.columns(3, gap="small")
        with d1:
            show_ld = st.checkbox("ℓd arrow", True, key="geo_show_ld")
            show_cover = st.checkbox("Cover", True, key="geo_show_cover")
        with d2:
            show_cb = st.checkbox("c_b", True, key="geo_show_cb")
            show_spacing = st.checkbox("s_clear / s_Ktr", True, key="geo_show_spacing")
        with d3:
            show_ktr = st.checkbox("Ktr zone", True, key="geo_show_ktr")
        st.markdown("</div>", unsafe_allow_html=True)

    with ktr_col:
        st.markdown("<div class='geo-ktr-wrap'>", unsafe_allow_html=True)
        st.markdown("<div class='geo-ktr-title'>⬡ Transverse Reinforcement / Ktr Preview</div>", unsafe_allow_html=True)
        st.markdown("<div class='geo-ktr-note'>Preview only — n is taken from the selected developed / lap-spliced group below.</div>", unsafe_allow_html=True)
        k1, k2, k3 = st.columns([1, 1, 1.2], gap="small")
        with k1:
            geo_Atr = st.number_input("A_tr (mm²)", min_value=0.0, value=float(st.session_state.get("geo_Atr_mm2", geo_Atr_default)), step=10.0, key="geo_Atr_mm2")
        with k2:
            _geo_unit_len = "mm" if str(units).upper() == "SI" else "cm"
            geo_s_tr_disp = st.number_input(f"s_Ktr ({_geo_unit_len})", min_value=0.0, value=float(st.session_state.get("geo_s_tr_display", geo_s_tr_disp_default)), step=10.0 if units == "SI" else 1.0, key="geo_s_tr_display")
        group_counts_for_ktr = {
            "Top face": int(bars_top_face),
            "Bottom face": int(bars_bottom_face),
            "Left side": int(bars_left_face),
            "Right side": int(bars_right_face),
        }
        ktr_group_options = list(group_counts_for_ktr.keys()) + ["Manual / Custom"]
        ktr_group = st.session_state.get("ktr_developed_group", "Top face")
        if ktr_group not in ktr_group_options:
            ktr_group = "Top face"
        with k3:
            ktr_group = st.selectbox(
                "Developed group (n)",
                ktr_group_options,
                index=ktr_group_options.index(ktr_group),
                key="ktr_developed_group",
            )
        geo_n_manual_default = int(snap_now.get("n_long", 1) or 1)
        geo_n_manual = int(st.session_state.get("geo_n_long_manual", geo_n_manual_default) or geo_n_manual_default)
        if ktr_group == "Manual / Custom":
            geo_n_manual = int(st.number_input("Manual n", min_value=1, max_value=100, value=int(geo_n_manual), step=1, key="geo_n_long_manual"))
            geo_n_long = int(geo_n_manual)
        else:
            geo_n_long = int(group_counts_for_ktr[ktr_group])
        geo_n_long = max(1, int(geo_n_long))
        st.caption(f"n = {geo_n_long:d}  ·  source: {ktr_group}")
        st.markdown("</div>", unsafe_allow_html=True)

    geo_s_tr_mm = to_si_len_from_display(float(geo_s_tr_disp), units)
    try:
        geo_ktr = ktr_aci(float(geo_Atr), float(geo_s_tr_mm), int(geo_n_long))
    except Exception:
        geo_ktr = 0.0

    if "calc_inputs" not in st.session_state or not isinstance(st.session_state.get("calc_inputs"), dict):
        st.session_state["calc_inputs"] = dict(calc_inputs or {})
    st.session_state.calc_inputs["Atr_mm2"] = float(geo_Atr)
    st.session_state.calc_inputs["s_tr_mm"] = float(geo_s_tr_mm)
    st.session_state.calc_inputs["n_long"] = int(geo_n_long)
    st.session_state.calc_inputs["Ktr_base"] = float(geo_ktr)

    gd["bars_top_face"] = int(bars_top_face)
    gd["bars_bottom_face"] = int(bars_bottom_face)
    gd["bars_left_face"] = int(bars_left_face)
    gd["bars_right_face"] = int(bars_right_face)
    gd["n_long"] = int(geo_n_long)
    gd["s_tr"] = float(geo_s_tr_mm)
    gd["ktr_developed_group"] = str(ktr_group)
    gd["ktr"] = float(geo_ktr)

    def strip_card(k: str, v: Any, s: str = "", accent: str = "") -> str:
        cls = f"geo-result-card {accent}".strip()
        return (
            f"<div class='{cls}'>"
            f"<div class='k'>{k}</div>"
            f"<div class='v'>{_safe_str(v)}</div>"
            f"<div class='s'>{s}</div>"
            f"</div>"
        )

    beta_val = float(gd.get("beta_eff", 0.0) or 0.0)
    ld_val   = float(gd.get("ld_t", 0.0) or 0.0)
    beta_accent = "hi" if beta_val >= 2.5 else ("ok" if beta_val >= 1.8 else "")
    ld_accent   = "hi" if ld_val > 0 else ""

    st.markdown("<div class='geo-mini-title'>Key Parameters</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='geo-result-strip'>"
        + strip_card("d_b", f"{gd['db']:.0f} {gd['unit_tag']}", "selected bar")
        + strip_card("c_b", f"{gd['cb']:.1f} {gd['unit_tag']}", f"governs: {gd.get('cb_gov', '—')}")
        + strip_card("K_tr", f"{geo_ktr:.2f} {gd['unit_tag']}", f"A_tr={float(geo_Atr):.1f}  s={float(geo_s_tr_disp):.1f}", "ok" if geo_ktr > 0 else "")
        + strip_card("n (Ktr)", int(geo_n_long), str(ktr_group))
        + strip_card("β = (c_b+K_tr)/d_b", f"{beta_val:.3f}", "capped at 2.5" if beta_val >= 2.5 else "ACI Eq. 25.4.2.4b", beta_accent)
        + strip_card("ℓd (result)", f"{ld_val:.0f} {gd['unit_tag']}", "from result table", ld_accent)
        + "</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div class='geo-mini-title'>Section Dimensions</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='geo-result-strip' style='grid-template-columns:repeat(4,minmax(110px,1fr));'>"
        + strip_card("Width  b", f"{float(gd.get('beam_b_mm', 400.0) or 400.0):.0f} mm", "ACI § 25.2")
        + strip_card("Depth  h", f"{float(gd.get('beam_h_mm', 600.0) or 600.0):.0f} mm", "ACI § 25.2")
        + strip_card("Stirrup d_b", f"{float(gd.get('stirrup_db_mm', 8.0) or 8.0):.0f} mm", "closed tie/stirrup")
        + strip_card("Cover (clear)", f"{float(gd.get('cover', 40.0) or 40.0):.1f} {gd['unit_tag']}", "to stirrup face")
        + "</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div class='geo-mini-title'>Geometry Canvas</div>", unsafe_allow_html=True)
    gv1, gv2 = st.columns([1.70, 1.0], gap="medium")
    with gv1:
        st.markdown(
            "<div class='geo-canvas-title'>▸ Longitudinal Beam Section</div>"
            "<div class='geo-canvas-sub'>Elevation view · bar layers, ℓd, cover, c_b, and Ktr zone · synced with section counts</div>",
            unsafe_allow_html=True,
        )
        render_svg_component(
            svg_dynamic_ld_geometry(
                gd,
                show_cover=show_cover,
                show_cb=show_cb,
                show_spacing=show_spacing,
                show_ktr=show_ktr,
                show_ld=show_ld,
            ),
            height=500,
        )

    with gv2:
        st.markdown(
            "<div class='geo-canvas-title'>▸ Beam Cross-Section</div>"
            "<div class='geo-canvas-sub'>Cut view · dimensions, cover, stirrup cage, and bar layout · graphics only</div>",
            unsafe_allow_html=True,
        )
        render_svg_component(
            svg_dynamic_cross_section(gd, show_cover=show_cover, show_spacing=show_spacing),
            height=470,
        )
        st.markdown("<div class='geo-bar-control-wrap'>", unsafe_allow_html=True)
        st.markdown(
            "<div class='geo-bar-control-title'>Bar Counts per Face "
            "<span style='font-weight:500;color:#5d82a0;font-size:10px;'>(corner bars included)</span></div>",
            unsafe_allow_html=True,
        )
        st.markdown("<div class='geo-bar-control-note'>Affects cross-section drawing and Ktr group n — graphics only, no ACI recalculation</div>", unsafe_allow_html=True)
        bc1, bc2 = st.columns(2, gap="small")
        with bc1:
            st.number_input("Top face", min_value=2, max_value=20, step=1, key="bars_top_face")
            st.number_input("Left face", min_value=2, max_value=20, step=1, key="bars_left_face")
        with bc2:
            st.number_input("Bottom face", min_value=2, max_value=20, step=1, key="bars_bottom_face")
            st.number_input("Right face", min_value=2, max_value=20, step=1, key="bars_right_face")
        st.markdown("</div>", unsafe_allow_html=True)

    with st.expander("▸ Hook Geometry Preview  (90° / 135° / 180°)", expanded=False):
        h1, h2, h3 = st.columns(3)
        with h1:
            render_svg_component(svg_dynamic_hook_preview(gd, 90), height=390)
        with h2:
            render_svg_component(svg_dynamic_hook_preview(gd, 135), height=390)
        with h3:
            render_svg_component(svg_dynamic_hook_preview(gd, 180), height=390)

    with st.expander("Data snapshot", expanded=False):
        snap = {
            "db_mm": gd["db"],
            "cover": gd["cover"],
            "s_clear_mm": gd["spacing"],
            "so": gd["so"],
            "cb": gd["cb"],
            "cb_governing_source": gd["cb_gov"],
            "bars_top_face": int(bars_top_face),
            "bars_bottom_face": int(bars_bottom_face),
            "bars_left_face": int(bars_left_face),
            "bars_right_face": int(bars_right_face),
            "Ktr_preview": float(geo_ktr),
            "ktr_n_source": str(ktr_group),
            "n_used_for_Ktr": int(geo_n_long),
            "beta_raw": gd["beta_raw"],
            "beta_used": gd["beta_eff"],
            "beta_capped": gd["beta_is_capped"],
            "ld_t_from_results": gd["ld_t"],
            "ldh_from_results": gd["ldh"],
            "edition": edition,
        }
        st.json(snap)

def _safe_str(value: Any, default: str = "not stored") -> str:
    """Return a readable string for report cards."""
    if value is None:
        return default
    try:
        if isinstance(value, float):
            if math.isnan(value):
                return default
            return f"{value:g}"
        return str(value)
    except Exception:
        return default


def _result_row_dict(df_res: Optional[pd.DataFrame]) -> Dict[str, Any]:
    """Return first result row as a dict, if available."""
    try:
        if df_res is not None and not getattr(df_res, "empty", True):
            return df_res.iloc[0].to_dict()
    except Exception:
        pass
    return {}


def _html_kv_table(items: List[tuple]) -> str:
    """Build a compact HTML key-value table for Streamlit markdown."""
    rows = []
    for k, v in items:
        rows.append(f'<div class="k">{k}</div><div class="v">{_safe_str(v)}</div>')
    return '<div class="aci-report-kv">' + "".join(rows) + "</div>"


def render_calc_transparency_report(
    calc_inputs: Dict[str, Any],
    df_res: Optional[pd.DataFrame],
    qa_messages: Optional[List[Dict[str, Any]]],
    units: str,
    edition: str,
) -> None:
    """
    Render a guaranteed-visible calculation transparency panel.

    This function is intentionally read-only:
    - reads st.session_state calc_inputs / df_res / qa_messages
    - does not recalculate ACI values
    - does not introduce new coefficients or clause claims
    """
    ci = calc_inputs or {}
    qa_messages = qa_messages or {}

    st.markdown("### Calculation Transparency / Data Lineage")
    st.caption(
        "Read-only report layer. It displays stored calculation data and result-table values; "
        "it does not perform a second ACI calculation."
    )

    st.info(
        "This panel should never be blank. If detailed snapshots are missing, "
        "the panel still shows session-state availability and the result table status."
    )

    # Always-visible status cards
    status_rows = [
        {"Item": "Selected edition", "Value": _safe_str(edition)},
        {"Item": "Selected units", "Value": _safe_str(units)},
        {"Item": "Engine version", "Value": _safe_str(globals().get("ENGINE_VERSION", "unknown"))},
        {"Item": "calc_inputs available", "Value": str(bool(ci))},
        {"Item": "calc_inputs keys", "Value": str(len(ci))},
        {"Item": "df_res available", "Value": str(df_res is not None and not getattr(df_res, "empty", True))},
        {"Item": "qa_messages count", "Value": str(len(qa_messages) if isinstance(qa_messages, list) else 0)},
    ]
    st.dataframe(_arrow_safe_df(pd.DataFrame(status_rows)), hide_index=True, width="stretch")

    st.markdown("#### Result Table Used by Screen Output")
    if df_res is not None and not getattr(df_res, "empty", True):
        st.dataframe(_arrow_safe_df(df_res), hide_index=True, width="stretch")
    else:
        st.warning("No result table is stored yet. Press Calculate / Update first.")

    st.markdown("#### Stored Calculation Inputs")
    if ci:
        st.dataframe(
            _arrow_safe_df(pd.DataFrame([{"Parameter": str(k), "Stored value": _safe_str(v)} for k, v in ci.items()])),
            hide_index=True,
            width="stretch",
        )
    else:
        st.warning(
            "calc_inputs is empty or not stored in session_state. "
            "The calculation may still have produced df_res, but the input snapshot was not persisted."
        )

    st.markdown("#### Key Transparency Fields")
    key_names = [
        "fc_mpa", "fy_mpa", "db_mm", "cover_mm", "spacing_mm", "so_mm",
        "cb_mm", "cb_gov", "Ktr_base", "beta_raw", "beta_eff",
        "beta_is_capped", "epoxy", "top_bar", "lambda", "lam"
    ]
    key_rows = [{"Field": k, "Value": _safe_str(ci.get(k, "not stored"))} for k in key_names]
    st.dataframe(_arrow_safe_df(pd.DataFrame(key_rows)), hide_index=True, width="stretch")

    st.markdown("#### QA Messages")
    if isinstance(qa_messages, list) and qa_messages:
        st.dataframe(
            _arrow_safe_df(pd.DataFrame([
                {
                    "Level": _safe_str(m.get("level")),
                    "Code": _safe_str(m.get("code")),
                    "Message": _safe_str(m.get("message")),
                    "Source": _safe_str(m.get("source")),
                    "Action": _safe_str(m.get("action")),
                }
                for m in qa_messages
            ])),
            hide_index=True,
            width="stretch",
        )
    else:
        st.info("No QA messages are stored for this run.")

    st.markdown("---")
    st.warning(
        "Formula-by-formula symbolic breakdown is intentionally not added here unless each "
        "formula and coefficient is verified against the official ACI text."
    )


# ---------------------------------------------------------------------------
# FIX11D — Safe shared selected bar state across single-bar tabs
# UI/state layer only. No ACI design formula is modified.
# ---------------------------------------------------------------------------

def get_shared_single_bar_db(main_list: List[float], fallback: Optional[float] = None) -> Optional[float]:
    """Return the shared selected bar diameter for single-bar UI/report tabs."""
    if not main_list:
        return fallback

    normalized = [float(x) for x in main_list]
    stored = st.session_state.get("selected_single_bar_db_mm", None)

    if stored is not None:
        try:
            stored_f = float(stored)
            if stored_f in normalized:
                return stored_f
            nearest = min(normalized, key=lambda x: abs(x - stored_f))
            st.session_state["selected_single_bar_db_mm"] = nearest
            return nearest
        except Exception:
            pass

    if fallback is not None:
        try:
            fallback_f = float(fallback)
            if fallback_f in normalized:
                st.session_state["selected_single_bar_db_mm"] = fallback_f
                return fallback_f
        except Exception:
            pass

    st.session_state["selected_single_bar_db_mm"] = normalized[0]
    return normalized[0]


def shared_single_bar_selectbox(label: str, main_list: List[float], key: str) -> Optional[float]:
    """Render a selectbox synchronized through st.session_state."""
    if not main_list:
        st.warning("No main bar list is available yet.")
        return None

    normalized = [float(x) for x in main_list]
    current = get_shared_single_bar_db(normalized)
    try:
        index = normalized.index(float(current))
    except Exception:
        index = 0

    selected = st.selectbox(
        label,
        normalized,
        index=index,
        key=key,
        format_func=lambda x: f"{float(x):.0f} mm",
    )
    if selected is not None:
        st.session_state["selected_single_bar_db_mm"] = float(selected)
    return float(selected) if selected is not None else None


# ---------------------------------------------------------------------------
# FIX11E — Result-row filtering for shared selected bar in Geometry Main
# UI/state layer only. No ACI design formula is modified.
# ---------------------------------------------------------------------------

def filter_df_res_for_selected_bar(df_res: Optional[pd.DataFrame], selected_db: Optional[float]) -> Optional[pd.DataFrame]:
    """
    Return a one-row result table matching selected_db when a recognizable
    bar-diameter column exists. If no matching column is found, return df_res.

    This is used only for display synchronization in Geometry Main.
    """
    if df_res is None or getattr(df_res, "empty", True) or selected_db is None:
        return df_res

    try:
        selected = float(selected_db)
    except Exception:
        return df_res

    candidate_cols = []
    for col in df_res.columns:
        name = str(col).lower()
        if (
            "db" in name
            or "diameter" in name
            or "bar" in name
            or "قطر" in name
        ):
            candidate_cols.append(col)

    for col in candidate_cols:
        try:
            vals = df_res[col].apply(lambda v: _num_from_any(v, default=float("nan")))
            mask = vals.sub(selected).abs() < 1e-6
            if bool(mask.any()):
                return df_res.loc[mask].head(1).copy()
        except Exception:
            continue

    return df_res


# ---------------------------------------------------------------------------
# FIX12A — Professional result header and result cards
# Display/report layer only. No ACI design formula is modified.
# This block is placed before main() so functions exist at runtime.
# ---------------------------------------------------------------------------

def _first_row_dict_safe(df: Optional[pd.DataFrame]) -> Dict[str, Any]:
    """Return first row dictionary from a dataframe, if possible."""
    try:
        if df is not None and not getattr(df, "empty", True):
            return df.iloc[0].to_dict()
    except Exception:
        pass
    return {}


def _find_result_value_by_tokens(row: Dict[str, Any], token_sets: List[List[str]], default: str = "—") -> tuple:
    """
    Find a result value from an existing result row using token matching.
    Returns (value, source_column). This is reporting only, not calculation.
    """
    if not row:
        return default, "No result row"

    normalized = [(str(k), str(k).lower(), v) for k, v in row.items()]
    for tokens in token_sets:
        low_tokens = [str(t).lower() for t in tokens]
        for original, low_name, value in normalized:
            if all(t in low_name for t in low_tokens):
                return value, original
    return default, "Not detected"


def _format_card_value(value: Any) -> str:
    """Format a value for compact display while preserving existing source data."""
    try:
        if value is None:
            return "—"
        if isinstance(value, float):
            if math.isnan(value):
                return "—"
            return f"{value:,.0f}" if abs(value) >= 10 else f"{value:g}"
        if isinstance(value, int):
            return f"{value:,}"
        s = str(value)
        return s if len(s) <= 28 else s[:25] + "…"
    except Exception:
        return str(value)


def render_professional_run_header(
    edition: str,
    units: str,
    calc_inputs: Dict[str, Any],
    qa_messages: Optional[List[Dict[str, Any]]],
) -> None:
    """Render a compact professional run header."""
    ci = calc_inputs or {}
    qa_messages = qa_messages or []
    critical = sum(1 for m in qa_messages if str(m.get("level", "")).upper() == "CRITICAL")
    warning = sum(1 for m in qa_messages if str(m.get("level", "")).upper() == "WARNING")
    status_cls = "aci-chip-ok" if critical == 0 and warning == 0 else "aci-chip-warn"
    status_txt = "OK" if critical == 0 and warning == 0 else f"{critical} critical / {warning} warning"

    db_txt = _safe_str(ci.get("db_mm", st.session_state.get("selected_single_bar_db_mm", "not stored")))
    fc_txt = _safe_str(ci.get("fc_mpa", "not stored"))
    fy_txt = _safe_str(ci.get("fy_mpa", "not stored"))
    beta_txt = _safe_str(ci.get("beta_eff", "not stored"))

    st.markdown(
        f"""
<div class="aci-run-header">
  <div class="aci-run-title">ACI Rebar Development & Splice Designer</div>
  <div class="aci-run-subtitle">
    Results-only calculation workspace. Header and cards below read from the existing calculation snapshot and result table.
  </div>
  <div class="aci-chip-row">
    <span class="aci-chip">Edition: {edition}</span>
    <span class="aci-chip">Units: {units}</span>
    <span class="aci-chip">snapshot db: {db_txt}</span>
    <span class="aci-chip">f'c: {fc_txt}</span>
    <span class="aci-chip">fy: {fy_txt}</span>
    <span class="aci-chip">β used: {beta_txt}</span>
    <span class="aci-chip {status_cls}">QA: {status_txt}</span>
  </div>
</div>
        """,
        unsafe_allow_html=True,
    )


def render_length_result_cards(df_res: Optional[pd.DataFrame]) -> None:
    """Render high-level result cards using only the existing result table."""
    row = _first_row_dict_safe(df_res)
    cards = [
        ("ℓd tension", [["ℓd_t"], ["ld_t"], ["tension"], ["ℓd", "t"]]),
        ("ℓdc compression", [["ℓdc"], ["ldc"], ["compression"]]),
        ("ℓdh hook", [["ℓdh"], ["ldh"], ["hook"]]),
        ("ℓs tension", [["ℓs_t"], ["ls_t"], ["splice", "tension"]]),
        ("ℓs compression", [["ℓs_c"], ["ls_c"], ["splice", "compression"]]),
        ("c_b", [["c_b"], ["cb"]]),
        ("Ktr", [["ktr"]]),
        ("β", [["β"], ["beta"]]),
    ]

    card_html = []
    for label, token_sets in cards:
        value, source = _find_result_value_by_tokens(row, token_sets)
        card_html.append(
            f"""
<div class="aci-result-card">
  <div class="label">{label}</div>
  <div class="value">{_format_card_value(value)}</div>
  <div class="source">source: {source}</div>
</div>
            """
        )

    st.markdown(
        f"""
<div class="aci-card-note">
Cards below are a visual summary of the existing result table. If a card shows “—”, the corresponding column was not detected by name.
</div>
<div class="aci-result-card-grid">
{''.join(card_html)}
</div>
        """,
        unsafe_allow_html=True,
    )



def _render_results_topbar():
    """نوار افقی تمام‌عرض: کارت‌های خلاصه نتایج + هشدارها + diff 2019/2025."""
    df_rp  = st.session_state.get("df_res", None)
    utag_p = "mm" if (st.session_state.get("calc_inputs", {}) or {}).get("units", "SI") == "SI" else "cm"
    if df_rp is None or getattr(df_rp, "empty", True):
        return

    row0 = df_rp.iloc[0].to_dict()

    with st.container(border=True):
        st.markdown(
            "<div style='font-size:13px;font-weight:700;color:#eaf3ff;"
            "margin-bottom:8px;'>Results summary "
            "<span style='font-weight:400;color:#9fb2c8;font-size:11px;'>"
            "(smallest bar size)</span></div>",
            unsafe_allow_html=True,
        )
        # کارت‌های خلاصه — افقی، تمام‌عرض
        items = [
            ("Development ℓd",  "ld_T"),
            ("Compression ℓdc", "ldc"),
            ("Hook 90° ℓdh",    "ldh 90°"),
            ("Hook 135° ℓdh",   "ldh 135°"),
            ("Hook 180° ℓdh",   "ldh 180°"),
            ("Lap splice ℓs",   "ls_T"),
            ("Lap splice comp.","ls_c"),
        ]
        cols = st.columns(len(items))
        for col, (label, key_part) in zip(cols, items):
            col_key = f"\u2113{key_part} ({utag_p})"
            val = row0.get(col_key, "\u2014")
            val_str = f"{val} {utag_p}" if isinstance(val, (int, float)) else str(val)
            with col:
                st.markdown(
                    f"<div style='font-size:10px;color:#9fb2c8;line-height:1.3;'>{label}</div>"
                    f"<div style='font-size:17px;font-weight:800;color:#34c98e;'>{val_str}</div>",
                    unsafe_allow_html=True,
                )

        # هشدارها + diff در دو ستون
        warn_col, diff_col = st.columns([1.4, 1])
        with warn_col:
            qa_msgs = st.session_state.get("qa_messages", [])
            crits = [m for m in qa_msgs if m.get("level") == "CRITICAL"]
            warns = [m for m in qa_msgs if m.get("level") == "WARNING"]
            infos = [m for m in qa_msgs if m.get("level") == "INFO"
                     and "minimum" in str(m.get("message", "")).lower()]
            if crits or warns or infos:
                for m in crits:
                    st.error(m.get("message", ""), icon="\U0001f6a8")
                for m in warns[:3]:
                    st.warning(m.get("message", ""), icon="\u26a0\ufe0f")
                for m in infos[:2]:
                    st.info(m.get("message", ""), icon="\u2139\ufe0f")
            else:
                st.success("All requirements satisfied.", icon="\u2705")

        with diff_col:
            if "_ld_2019" in row0:
                st.markdown(
                    "<div style='font-size:11px;font-weight:700;color:#9fb2c8;"
                    "margin-bottom:4px;'>ACI 318-19 vs 318-25</div>",
                    unsafe_allow_html=True,
                )
                for lbl, k19, k25 in [
                    ("\u2113d",  "_ld_2019",  "_ld_2025"),
                    ("\u2113dc", "_ldc_2019", "_ldc_2025"),
                    ("\u2113dh", "_ldh_2019", "_ldh_2025"),
                    ("\u2113s",  "_ls_2019",  "_ls_2025"),
                ]:
                    v19 = row0.get(k19, "\u2014")
                    v25 = row0.get(k25, "\u2014")
                    try:
                        d  = int(v25) - int(v19)
                        ds = f"+{d}" if d > 0 else str(d)
                        dc = "#f97316" if d > 0 else ("#34c98e" if d < 0 else "#9fb2c8")
                    except Exception:
                        ds, dc = "\u2014", "#9fb2c8"
                    st.markdown(
                        "<div style='display:grid;grid-template-columns:1fr 50px 50px 44px;"
                        "font-size:11px;padding:2px 0;'>"
                        f"<span style='color:#c8dff5'>{lbl}</span>"
                        f"<span style='color:#9fb2c8'>{v19}</span>"
                        f"<span style='color:#eaf3ff'>{v25}</span>"
                        f"<span style='color:{dc};font-weight:700'>{ds}</span>"
                        "</div>",
                        unsafe_allow_html=True,
                    )


def _render_results_panel():
    """Right column: Results Summary + Warnings + 2019 vs 2025 diff."""
    df_rp  = st.session_state.get("df_res", None)
    utag_p = "mm" if (st.session_state.get("calc_inputs", {}) or {}).get("units", "SI") == "SI" else "cm"

    st.markdown(
        "<div style='font-size:13px;font-weight:700;color:#eaf3ff;"
        "margin-bottom:8px;'>Results summary</div>",
        unsafe_allow_html=True,
    )

    if df_rp is None or getattr(df_rp, "empty", True):
        st.info("Set inputs and press **Calculate / Update**.")
        return

    row0 = df_rp.iloc[0].to_dict()
    items = [
        ("Development length (ℓd)",   "ld_T"),
        ("Compression (ℓdc)",          "ldc"),
        ("Hook 90°",                   "ldh 90°"),
        ("Hook 135°",                  "ldh 135°"),
        ("Hook 180°",                  "ldh 180°"),
        ("Lap splice tension",               "ls_T"),
        ("Lap splice comp.",                 "ls_c"),
    ]
    for label, key_part in items:
        col_key = f"ℓ{key_part} ({utag_p})"
        val = row0.get(col_key, "—")
        color = "#34c98e" if isinstance(val, (int, float)) else "#9fb2c8"
        val_str = f"{val} {utag_p}" if isinstance(val, (int, float)) else str(val)
        st.markdown(
            "<div style='display:flex;justify-content:space-between;"
            "padding:5px 0;border-bottom:0.5px solid rgba(120,170,220,.15);"
            "font-size:12px;'>"
            f"<span style='color:#9fb2c8'>{label}</span>"
            f"<span style='color:{color};font-weight:700'>{val_str}</span>"
            "</div>",
            unsafe_allow_html=True,
        )

    qa_msgs = st.session_state.get("qa_messages", [])
    crits = [m for m in qa_msgs if m.get("level") == "CRITICAL"]
    warns = [m for m in qa_msgs if m.get("level") == "WARNING"]
    infos = [m for m in qa_msgs if m.get("level") == "INFO"
             and "minimum" in str(m.get("message", "")).lower()]
    if crits or warns or infos:
        st.markdown(
            "<div style='font-size:12px;font-weight:700;color:#ffb020;"
            "margin:10px 0 4px;'>⚠ Warnings / Notes</div>",
            unsafe_allow_html=True,
        )
        for m in crits:
            st.error(m.get("message", ""), icon="🚨")
        for m in warns:
            st.warning(m.get("message", ""), icon="⚠️")
        for m in infos:
            st.info(m.get("message", ""), icon="ℹ️")
    else:
        st.success("All requirements satisfied.", icon="✅")

    if "_ld_2019" in row0:
        st.markdown(
            "<div style='font-size:12px;font-weight:700;color:#9fb2c8;"
            "margin:10px 0 4px;'>ACI 318-19 vs 318-25</div>",
            unsafe_allow_html=True,
        )
        for lbl, k19, k25 in [
            ("ℓd",  "_ld_2019",  "_ld_2025"),
            ("ℓdc", "_ldc_2019", "_ldc_2025"),
            ("ℓdh", "_ldh_2019", "_ldh_2025"),
            ("ℓs",  "_ls_2019",  "_ls_2025"),
        ]:
            v19 = row0.get(k19, "—")
            v25 = row0.get(k25, "—")
            try:
                d  = int(v25) - int(v19)
                ds = f"+{d}" if d > 0 else str(d)
                dc = "#f97316" if d > 0 else ("#34c98e" if d < 0 else "#9fb2c8")
            except Exception:
                ds, dc = "—", "#9fb2c8"
            st.markdown(
                "<div style='display:grid;grid-template-columns:1fr 40px 40px 38px;"
                "font-size:11px;padding:4px 0;"
                "border-bottom:0.5px solid rgba(120,170,220,.10);'>"
                f"<span style='color:#c8dff5'>{lbl}</span>"
                f"<span style='color:#9fb2c8'>{v19}</span>"
                f"<span style='color:#eaf3ff'>{v25}</span>"
                f"<span style='color:{dc};font-weight:700'>{ds}</span>"
                "</div>",
                unsafe_allow_html=True,
            )


def main():
    db_mm = None; class_val = None; alpha_used = None; lam = None
    compare_mode = False; show_diff_only = False; is_compare_mode = False
    top_bar = False; epoxy = False; lightweight = False
    hook_conf_mode = "Manual"; hook_conf_manual = False; s_hook_d = 0.0
    Atr = 0.0; s_tr = 0.0; n_long = 1

    st.set_page_config(
        page_title="ACI Rebar Designer – Chapter 25",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_custom_css()

    if "data_calculated" not in st.session_state:
        st.session_state.data_calculated = False
        st.session_state.df_res = None
        st.session_state.df_bosb_main = None
        st.session_state.df_bosb_stir = None
        st.session_state.main_list = []
        st.session_state.calc_inputs = {}
        st.session_state.excel_blob = None
        st.session_state.html_report = None
        st.session_state.html_report_filename = None
        st.session_state.class_val = "Class B"
        st.session_state.alpha_used = 1.3
        st.session_state.export_errors = {"excel": None, "html": None}
        st.session_state.calc_run_id = None
        st.session_state.calc_run_ts = None
        st.session_state.excel_filename = None

    # =======================================================================
    #  SIDEBAR — Code | Mode | Options
    # =======================================================================
    with st.sidebar:
        st.markdown(
            "<div style='display:flex;align-items:center;gap:8px;margin-bottom:2px'>"
            "<div style='background:#1c86ff;color:#fff;font-weight:900;font-size:13px;"
            "padding:4px 8px;border-radius:6px;'>ACI</div>"
            "<div><div style='font-size:14px;font-weight:700;'>Rebar Designer</div>"
            "<div style='font-size:10px;color:#9fb2c8;'>Chapter 25</div></div></div>",
            unsafe_allow_html=True,
        )
        st.divider()

        st.markdown(
            "<div style='font-size:10px;font-weight:700;color:#9fb2c8;"
            "text-transform:uppercase;letter-spacing:.7px;margin-bottom:4px;'>Code</div>",
            unsafe_allow_html=True,
        )
        _ced, _cun = st.columns(2)
        with _ced:
            edition = st.selectbox("Code", ["ACI318-2025", "ACI318-2019"],
                                   label_visibility="collapsed", key="sb_edition")
        with _cun:
            units = st.selectbox("Units", ["SI", "KGF-CM"],
                                 label_visibility="collapsed", key="sb_units")
        st.divider()

        st.markdown(
            "<div style='font-size:10px;font-weight:700;color:#9fb2c8;"
            "text-transform:uppercase;letter-spacing:.7px;margin-bottom:4px;'>Mode</div>",
            unsafe_allow_html=True,
        )
        user_mode = st.radio(
            "Mode",
            ["🚀 Professional", "🎓 Educational"],
            index=0,
            format_func=lambda m: {
                "🚀 Professional": "Results only",
                "🎓 Educational":  "Detailed breakdown",
            }.get(m, m),
            label_visibility="collapsed",
            key="sb_mode",
        )
        st.divider()

        st.markdown(
            "<div style='font-size:10px;font-weight:700;color:#9fb2c8;"
            "text-transform:uppercase;letter-spacing:.7px;margin-bottom:4px;'>Options</div>",
            unsafe_allow_html=True,
        )
        special_seismic_region = st.checkbox(
            "Special seismic region", value=False,
            help="Hides certain Chapter 25 outputs restricted by Chapter 18.",
            key="sb_seismic",
        )
        rounding_mode = st.selectbox(
            "Rounding", ["Exact", "UP to 50", "NEAREST 50"],
            index=1, key="sb_round",
        )
        st.session_state.setdefault("ch25_append_full_edition_comparison", False)
        st.checkbox(
            "Append 2019↔2025 comparison",
            help="Appends full side-by-side comparison to Detailed Report.",
            key="ch25_append_full_edition_comparison",
        )
        st.divider()
        _show_ref = st.checkbox("Show parameter reference", value=False, key="show_ch25_ref")
        if _show_ref:
            with st.container(border=True):
                st.markdown(build_aci_parameter_reference_md(edition, units))
                _ci_r = st.session_state.get("calc_inputs", {}) or {}
                if isinstance(_ci_r, dict) and _ci_r:
                    _rows_r = [{"Parameter": k, "Last used value": str(v)}
                               for k, v in _ci_r.items()]
                    st.dataframe(_arrow_safe_df(pd.DataFrame(_rows_r)),
                                 use_container_width=True, hide_index=True)

    # Cover page — only before first calculation
    if not st.session_state.get("data_calculated", False):
        st.markdown(COVER_MD, unsafe_allow_html=True)
        st.divider()
    else:
        st.markdown(
            "<div style='display:flex;align-items:center;gap:10px;margin-bottom:10px;'>"
            "<span style='font-size:13px;font-weight:700;color:#eaf3ff;'>"
            "ACI 318 – Chapter 25</span>"
            "<span style='font-size:11px;color:#9fb2c8;'>"
            "Development, Anchorage &amp; Splices</span></div>",
            unsafe_allow_html=True,
        )

    # =======================================================================
    #  SINGLE-COLUMN LAYOUT — full-width inputs; results summary bar on top
    # =======================================================================
    unit_len = "mm" if units == "SI" else "cm"

    # نوار خلاصه نتایج (بالای فرم، تمام‌عرض) — فقط بعد از اولین محاسبه
    if st.session_state.get("data_calculated", False):
        _render_results_topbar()

    # col_inp اکنون یک container تمام‌عرض است (col_out حذف شد)
    col_inp = st.container()
    col_out = None  # سازگاری با کد پایین

    with col_inp:
        edu_modifiers = {}
        # ********************************************************************
        unit_area = f"{unit_len}²"
        unit_str = "MPa" if units == "SI" else "kgf/cm²"
        apply_round = rounding_mode != "Exact"

        # Default values (same as v60)
        fc_d, fy_d = (30.0, 400.0) if units == "SI" else (250.0, 4000.0)
        cover_d, spacing_d, so_d = (40.0, 60.0, 0.0) if units == "SI" else (4.0, 6.0, 0.0)
        beam_b_mm = float((st.session_state.get("calc_inputs", {}) or {}).get("beam_b_mm", 400.0) or 400.0)
        beam_h_mm = float((st.session_state.get("calc_inputs", {}) or {}).get("beam_h_mm", 600.0) or 600.0)
        stirrup_db_mm = float((st.session_state.get("calc_inputs", {}) or {}).get("stirrup_db_mm", 8.0) or 8.0)
        lam, epoxy = 1.0, False
        psi_t_T, psi_t_B = 1.3, 1.0
        method_options = [
            "Both methods (compare: General & Simplified)",
            "General method (ACI Eq. 25.4.2.4a)",
            "Simplified method (ACI Table 25.4.2.3)",
        ]
        ld_method = method_options[0]  # Default: Both methods; keeps Ktr inputs enabled by default.
        splice_mode = "Auto"
        as_ratio, pct_spliced = 1.0, 100.0
        Atr, s_tr, n_long = 0.0, 0.0, 1

        # FIX12N: Ktr has ONE input source only: Geometry Main tab.
        # Other panels are read-only for Ktr to avoid calculation/report mismatches.
        _ktr_source = st.session_state.get("calc_inputs", {}) or {}
        Atr = float(_ktr_source.get("Atr_mm2", 0.0) or 0.0)
        s_tr = from_si_len_to_display(float(_ktr_source.get("s_tr_mm", 0.0) or 0.0), units)
        n_long = int(_ktr_source.get("n_long", 1) or 1)

        psi_g_val = 1.0
        hook_angle = 90
        has_conf, cover_ok, psi_r_comp = False, False, 1.0
        enable_headed, psi_p_h, psi_o_h, psi_c_h = False, 1.0, 1.0, 1.0
        bars_str = "10,12,14,16,18,20,22,25,28,32"
        stirrup_str = "8,10,12,14,16,18,20,22,25"

        # Store educational-mode inputs in this dictionary so they remain available during calculation.
        edu_inputs = {
            "cover_d": cover_d,
            "spacing_d": spacing_d,
            "so_d": so_d,
            "beam_b_mm": beam_b_mm,
            "beam_h_mm": beam_h_mm,
            "stirrup_db_mm": stirrup_db_mm,
            "epoxy": epoxy,
            "top_bar": False,
            "lam": lam,
            "hook_angle": hook_angle,
            "hook_conf_mode": hook_conf_mode,
            "hook_conf_manual": hook_conf_manual,
            "s_hook_d": s_hook_d,
            "Atr": Atr,
            "s_tr": s_tr,
            "n_long": n_long,
        }

    # ================================
        # MAIN INPUT FORM (UI-only changes)
        # ================================
        def _on_submit_main_form():
            """Submit callback: lock a calculation snapshot for deterministic UI/exports."""
            st.session_state["calc_ready"] = True
            st.session_state["calc_triggered"] = True
            # Release hygiene: stable per-run identifier for filenames and artifacts
            st.session_state["calc_run_id"] = now_tehran().strftime("%m%d-%H%M_TEHRAN")
            st.session_state["calc_run_ts"] = now_tehran().strftime("%Y-%m-%d %H:%M Tehran")

        with st.container():  # was st.form; replaced to allow download buttons in tabs

            # V70 UX: Detailed-breakdown modifiers moved from sidebar to main input area
            if str(user_mode).startswith("🎓"):
                st.markdown('<div class="aci-section-title">🎓 Detailed breakdown – Chapter 25 modifiers</div>', unsafe_allow_html=True)
                with st.container(border=True):
                    st.caption(
                        "Select conditions to determine ψ-factors from ACI Chapter 25 tables. "
                        "Defaults are conservative when conditions are unknown."
                    )
                    # λ inputs moved to item detailed inputs (25.4). Keep conservative defaults here.
                    edu_modifiers["lambda_key"] = "normal"

                    if str(edition).startswith("ACI318-2025") and units == "SI":
                        # Hooked bars – size group ψs is determined automatically from db (Table 25.4.3.2)
                        psi_s_group = "Auto (from db)"
                        edu_modifiers["hook_psi_s_group"] = psi_s_group
                        psi_cc_cond = st.selectbox(
                            "Hooked bars – cover condition ψcc (Table 25.4.3.2)",
                            ["Other / unknown (ψcc=1.0)", "Meets cover requirements (ψcc=0.7)"],
                            index=0,
                            key="edu_hook_psicc",
                        )
                        edu_modifiers["hook_psi_cc_key"] = "meets_cover" if psi_cc_cond.startswith("Meets") else "other"

                        psi_r_cond = st.selectbox(
                            "Hooked bars – confinement ψr (Table 25.4.3.2)",
                            ["Other / unknown (ψr=1.0)", "Meets confinement requirements (ψr=0.8)"],
                            index=0,
                            key="edu_hook_psir",
                        )
                        edu_modifiers["hook_psi_r_key"] = "meets_tie_confinement" if psi_r_cond.startswith("Meets") else "other"

                    st.markdown("**Straight bars – tension modifiers (Table 25.4.2.5)**")
                    st.caption(
                        "These options show the app's automatic factor selection. Override is disabled for parameters derived directly from db/fy to prevent inconsistent inputs. "
                        "If you need manual override for ψ-factors, use the Results only report settings."
                    )
                    col_m1, col_m2 = st.columns(2)
                    with col_m1:
                        # λ inputs moved to item detailed inputs (25.4). Keep conservative defaults here.
                        edu_modifiers["lambda_mode"] = "input"

                        psi_t_mode = st.selectbox(
                            "ψt (top-bar factor)",
                            ["Evaluate both (Top & Other)", "Assume Top-cast only", "Assume Other only"],
                            index=0,
                            key="edu_psit_mode",
                        )
                        edu_modifiers["psi_t_mode"] = psi_t_mode

                    with col_m2:
                        psi_e_mode = st.selectbox(
                            "ψe (epoxy factor)",
                            [
                                "Auto (from epoxy + cover/spacing)",
                                "Force epoxy-other (Table value)",
                                "Force epoxy-low cover/spacing (Table value)",
                            ],
                            index=0,
                            key="edu_psie_mode",
                        )
                        edu_modifiers["psi_e_mode"] = psi_e_mode

                        # ψs is determined automatically from db (Table 25.4.2.5)
                        psi_s_mode = "Auto (from db)"
                        edu_modifiers["psi_s_mode"] = psi_s_mode

                        # ψg is determined automatically from fy (Table 25.4.2.5)
                        psi_g_mode = "Auto (from fy)"
                        edu_modifiers["psi_g_mode"] = psi_g_mode
                        st.session_state["edu_modifiers"] = edu_modifiers
            if user_mode.startswith("🚀"):
                st.markdown("### ⚙️ Design inputs (Results only)")
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.markdown('<div class="aci-section-title">1) Materials</div>', unsafe_allow_html=True)
                    fc_d = st.number_input(f"Concrete f'c ({unit_str})", value=fc_d)
                    fy_d = st.number_input(f"Steel fy ({unit_str})", value=fy_d)
                lam = float((st.session_state.get('calc_inputs', {}) or {}).get('lam', 1.0) or 1.0)
                with c2:
                    st.markdown('<div class="aci-section-title">2) Geometry</div>', unsafe_allow_html=True)
                    beam_b_mm = st.number_input("Beam width b (mm)", min_value=1.0, value=float(beam_b_mm), step=10.0)
                    beam_h_mm = st.number_input("Beam depth h (mm)", min_value=1.0, value=float(beam_h_mm), step=10.0)
                    cover_d = st.number_input(f"Clear cover ({unit_len})", value=cover_d)
                    spacing_d = st.number_input(f"Clear spacing ({unit_len})", value=spacing_d)
                    epoxy = st.checkbox("Epoxy coated bars?", value=False)
                    st.session_state['epoxy'] = bool(epoxy)
                with c3:
                    st.markdown('<div class="aci-section-title">3) Splice class</div>', unsafe_allow_html=True)
                    splice_mode = st.selectbox(
                        "Splice class mode", ["Auto", "Class A", "Class B"], index=0
                    )
                    if splice_mode == "Auto":
                        as_ratio = st.number_input("As,prov / As,req", value=1.0)
                        pct_spliced = st.number_input("% bars spliced", value=100.0)
                with c4:
                    st.markdown('<div class="aci-section-title">4) Bar sets</div>', unsafe_allow_html=True)
                    bars_str = st.text_input("Main bars (mm)", value=bars_str)

                    stirrup_str = st.text_input("Stirrups / ties (mm)", value=stirrup_str)
                    stirrup_options = parse_bars_mm(stirrup_str) or [8.0]
                    stirrup_index = min(range(len(stirrup_options)), key=lambda i: abs(float(stirrup_options[i]) - float(stirrup_db_mm)))
                    stirrup_db_mm = float(st.selectbox(
                        "Selected stirrup/tie db (mm)",
                        stirrup_options,
                        index=stirrup_index,
                        help="Graphics-only selected closed stirrup/tie diameter for the Cross Section SVG.",
                    ))

                    ld_method = st.selectbox(
                        "Development method (ℓd)",
                        method_options,
                        index=0,
                        help="General: ACI Eq. 25.4.2.4a. Simplified: ACI Table 25.4.2.3 (only if conditions are satisfied).",
                        key="ld_method",
                    )

                    is_simplified_mode = str(ld_method).startswith("Simplified method")
                st.markdown('<div class="aci-section-title">Advanced confinement & hooks</div>', unsafe_allow_html=True)
                with st.container(border=True):
                    ac1, ac2 = st.columns(2)
                    with ac1:
                        ktr_read = st.session_state.get("calc_inputs", {}) or {}
                        Atr = float(ktr_read.get("Atr_mm2", 0.0) or 0.0)
                        s_tr = from_si_len_to_display(float(ktr_read.get("s_tr_mm", 0.0) or 0.0), units)
                        n_long = int(ktr_read.get("n_long", 1) or 1)
                        Ktr_base_preview = ktr_aci(Atr, to_si_len_from_display(float(s_tr), units), n_long)
                        st.markdown("**Ktr source**")
                        st.caption("Controlled only in Geometry Main tab to avoid duplicate inputs.")
                        st.write(f"A_tr = {Atr:.2f} mm²")
                        st.write(f"s = {s_tr:.2f} {unit_len}")
                        st.write(f"n = {n_long}")
                        st.write(f"Ktr = {Ktr_base_preview:.2f} mm")
                    with ac2:
                        hook_angle = ac2.selectbox("Hook angle", [90, 135, 180])

                        hook_conf_mode = ac2.selectbox(
                            "Hook confinement mode",
                            ["Manual", "Auto (Needs verification)"],
                            index=0,
                        )
                        hook_conf_manual = ac2.checkbox("Hook confinement provided (manual)")
                        s_hook_d = ac2.number_input(
                            f"Hook tie spacing s_hook ({unit_len})",
                            min_value=0.0,
                            value=0.0,
                            help="Spacing of transverse reinforcement relevant to the hooked bar region (detail-specific).",
                        )
                        enable_headed = ac2.checkbox("Include headed bar check?")

            elif user_mode.startswith("ℹ️"):
                st.markdown("### ℹ️ Intermediate mode – guided inputs")
                col_i1, col_i2 = st.columns(2)
                with col_i1:
                    st.markdown('<div class="aci-section-title">1) Materials</div>', unsafe_allow_html=True)
                    fc_d = st.number_input(f"Concrete f'c ({unit_str})", value=fc_d, help="Specified compressive strength used in design.")
                    fy_d = st.number_input(f"Steel fy ({unit_str})", value=fy_d, help="Yield strength of longitudinal bars.")
                    lam = float(lam or 1.0)  # λ moved to item detailed inputs
                with col_i2:
                    st.markdown('<div class="aci-section-title">2) Cover & spacing</div>', unsafe_allow_html=True)
                    cover_d = st.number_input(f"Clear cover ({unit_len})", value=cover_d, help="From bar surface to tension face.")
                    spacing_d = st.number_input(
                        f"Clear spacing between bars ({unit_len})", value=spacing_d,
                        help="Clear distance between bar surfaces."
                    )
                    epoxy = st.toggle("Epoxy coated bars?", value=False)

                st.markdown('<div class="aci-section-title">3) Splice class</div>', unsafe_allow_html=True)
                splice_mode = st.radio(
                    "Splice class",
                    ["Auto", "Class A", "Class B"],
                    index=0,
                    help="Auto uses Table 25.5.2.1 based on As,prov/As,req and % spliced.",
                )
                if splice_mode == "Auto":
                    as_ratio = st.number_input("As,prov / As,req", value=1.0)
                    pct_spliced = st.number_input("% of bars spliced", value=100.0)

                _method_for_disable = st.session_state.get("ld_method", ld_method)
                is_simplified_mode = str(_method_for_disable).startswith("Simplified method")
                st.markdown('<div class="aci-section-title">Advanced confinement options</div>', unsafe_allow_html=True)
                with st.container(border=True):
                    c_adv1, c_adv2 = st.columns(2)
                    with c_adv1:
                        ktr_read = st.session_state.get("calc_inputs", {}) or {}
                        Atr = float(ktr_read.get("Atr_mm2", 0.0) or 0.0)
                        s_tr = from_si_len_to_display(float(ktr_read.get("s_tr_mm", 0.0) or 0.0), units)
                        st.markdown("**Ktr source**")
                        st.caption("Controlled only in Geometry Main tab.")
                        st.write(f"A_tr = {Atr:.2f} mm²")
                        st.write(f"s = {s_tr:.2f} {unit_len}")
                    with c_adv2:
                        n_long = int((st.session_state.get("calc_inputs", {}) or {}).get("n_long", 1) or 1)
                        st.write(f"n = {n_long}")
                        hook_angle = c_adv2.selectbox("Hook angle for ℓdh", [90, 135, 180])

                st.divider()
                st.markdown('<div class="aci-section-title">4) Bar sets</div>', unsafe_allow_html=True)
                bars_str = st.text_input("Main bars (mm)", value=bars_str)
                stirrup_str = st.text_input("Stirrups / ties (mm)", value=stirrup_str)

                ld_method = st.selectbox(
                    "Development method (ℓd)",
                    method_options,
                    index=0,
                    help="Both: show General + Simplified side-by-side (when simplified conditions are satisfied). "
                         "General: ACI Eq. 25.4.2.4a. Simplified: ACI Table 25.4.2.3 (only if conditions are satisfied).",
                    key="ld_method",
                )

            else:  # Detailed breakdown mode
                st.markdown("### 🎓 Detailed breakdown – guided inputs (ordered by ACI Chapter 25)")
                st.caption("Inputs are organized according to the ACI 318 Chapter 25 workflow.")

                # -----------------------------
                # 25.1 — General & Materials
                # -----------------------------
                with st.expander("25.1) General & Materials", expanded=True):

                    col_g1, col_g2 = st.columns(2)
                    with col_g1:
                        fc_d = st.number_input(
                            f"Concrete f'c ({unit_str})",
                            value=fc_d,
                            help="Specified compressive strength used for the selected design checks."
                        )
                    with col_g2:
                        fy_d = st.number_input(
                            f"Steel fy ({unit_str})",
                            value=fy_d,
                            help="Specified yield strength of reinforcement used for the selected design checks."
                        )

                    st.markdown('<div class="aci-section-subtitle">Bar sets</div>', unsafe_allow_html=True)
                    bars_str = st.text_input(
                        "Main bars (comma-separated, e.g., #5,#6,#8 or 16,20,25)",
                        value=bars_str,
                        help="List the developed bar sizes used in the check. Use US bar marks (#) or SI diameters (mm) consistent with Units."
                    )
                    stirrup_str = st.text_input(
                        "Transverse bars (stirrups/ties) (comma-separated)",
                        value=stirrup_str,
                        help="List transverse reinforcement bar sizes potentially used for confinement / Ktr calculations (if applicable)."
                    )
                    stirrup_options = parse_bars_mm(stirrup_str) or [8.0]
                    stirrup_index = min(range(len(stirrup_options)), key=lambda i: abs(float(stirrup_options[i]) - float(stirrup_db_mm)))
                    edu_inputs["stirrup_db_mm"] = float(st.selectbox(
                        "Selected stirrup/tie db (mm)",
                        stirrup_options,
                        index=stirrup_index,
                        help="Graphics-only selected closed stirrup/tie diameter for the Cross Section SVG.",
                    ))

                    # Development method (ℓd) — placed after main + transverse bar sizes (per user UX requirement)
                    ld_method = st.selectbox(
                        "Development method (ℓd)",
                        method_options,
                        index=0,
                        help="General: ACI Eq. 25.4.2.4a. Simplified: ACI Table 25.4.2.3 (only if conditions are satisfied).",
                        key="ld_method",
                    )

                    # -----------------------------
                    # 25.2 – Cover & spacing
                    # -----------------------------
                with st.expander("25.2) Clear cover and spacing", expanded=True):
                    st.markdown(
                        """
                        <style>
                        .aci-252-label {
                            min-height: 1.85rem;
                            display: flex;
                            align-items: flex-end;
                            margin-bottom: 0.2rem;
                            color: #dbe7f2;
                            font-size: 0.82rem;
                            font-weight: 650;
                            line-height: 1.15;
                        }
                        </style>
                        """,
                        unsafe_allow_html=True,
                    )
                    with st.container(border=True):
                        c0, c1, c2, c3, c4 = st.columns([1, 1, 1, 1, 1], gap="small")
                        with c0:
                            st.markdown("<div class='aci-252-label'>Beam width b (mm)</div>", unsafe_allow_html=True)
                            edu_inputs["beam_b_mm"] = st.number_input(
                                "Beam width b (mm)",
                                min_value=1.0,
                                value=float(beam_b_mm),
                                step=10.0,
                                help="Graphics-only cross-section width used to scale the SVG.",
                                label_visibility="collapsed",
                            )
                        with c1:
                            st.markdown("<div class='aci-252-label'>Beam depth h (mm)</div>", unsafe_allow_html=True)
                            edu_inputs["beam_h_mm"] = st.number_input(
                                "Beam depth h (mm)",
                                min_value=1.0,
                                value=float(beam_h_mm),
                                step=10.0,
                                help="Graphics-only cross-section depth used to scale the SVG.",
                                label_visibility="collapsed",
                            )
                        with c2:
                            st.markdown(f"<div class='aci-252-label'>Clear cover ({unit_len})</div>", unsafe_allow_html=True)
                            edu_inputs["cover_d"] = st.number_input(
                                f"Clear cover ({unit_len})",
                                value=cover_d,
                                help="Clear distance from the bar surface to the concrete tension face.",
                                label_visibility="collapsed",
                            )
                        with c3:
                            st.markdown(f"<div class='aci-252-label'>s_clear ({unit_len})</div>", unsafe_allow_html=True)
                            edu_inputs["spacing_d"] = st.number_input(
                                f"s_clear ({unit_len})",
                                value=spacing_d,
                                help="Clear spacing between adjacent bars.",
                                label_visibility="collapsed",
                            )
                        with c4:
                            st.markdown(f"<div class='aci-252-label'>s_o ({unit_len})</div>", unsafe_allow_html=True)
                            edu_inputs["so_d"] = st.number_input(
                                f"s_o ({unit_len})",
                                value=so_d,
                                help="Enter this value if another bar layer or adjacent surface exists.",
                                label_visibility="collapsed",
                            )

                    # -----------------------------
                    # 25.3 – Hook geometry (angle, confinement)
                    # -----------------------------
                with st.expander("25.3) Hook geometry and confinement", expanded=False):
                    with st.container(border=True):
                        edu_inputs["hook_angle"] = st.selectbox("Hook angle", [90, 135, 180], index=0)
                        edu_inputs["hook_conf_mode"] = st.selectbox(
                            "Hook confinement mode",
                            ["Manual", "Auto (Needs verification)"],
                            index=0,
                            help="If 'Manual', use the checkbox below. If 'Auto', the app uses s_hook ≤ 3·db as criterion."
                        )
                        edu_inputs["hook_conf_manual"] = st.checkbox("Hook confinement provided (manual)", value=False)
                        edu_inputs["s_hook_d"] = st.number_input(
                            f"Hook tie spacing s_hook ({unit_len})",
                            min_value=0.0,
                            value=0.0,
                            help="Spacing of transverse reinforcement relevant to the hooked bar region."
                        )

                    # -----------------------------
                    # 25.4 – Modification factors (epoxy, top-bar, Ktr, λ)
                    # -----------------------------
                with st.expander("25.4) Modification factors (epoxy, top-bar, Ktr, λ)", expanded=False):
                    with st.container(border=True):
                        col_m1, col_m2 = st.columns(2)
                        with col_m1:
                            edu_inputs["epoxy"] = st.toggle("Epoxy coated bars?", value=False, help="ψe factor")
                            edu_inputs["top_bar"] = st.toggle("Top-bar condition?", value=False, help="ψt = 1.3 if True, else 1.0")
                        with col_m2:
                            lam_case = st.selectbox(
                                "Concrete unit weight (λ)",
                                ["Normalweight (λ = 1.0)", "Lightweight concrete (λ = 0.75)"],
                                index=0,
                                help="Lightweight concrete reduces bond strength."
                            )
                            edu_inputs["lam"] = 1.00 if lam_case.startswith("Normalweight") else 0.75

                        st.markdown("**Transverse reinforcement for Ktr (General method only)**")
                        st.caption("Ktr inputs are controlled only in Geometry Main tab. This panel is read-only to keep calculation and drawing consistent.")
                        ktr_read = st.session_state.get("calc_inputs", {}) or {}
                        edu_inputs["Atr"] = float(ktr_read.get("Atr_mm2", 0.0) or 0.0)
                        edu_inputs["s_tr"] = from_si_len_to_display(float(ktr_read.get("s_tr_mm", 0.0) or 0.0), units)
                        edu_inputs["n_long"] = int(ktr_read.get("n_long", 1) or 1)
                        k1, k2, k3, k4 = st.columns(4)
                        k1.metric("A_tr", f"{edu_inputs['Atr']:.2f} mm²")
                        k2.metric("s", f"{edu_inputs['s_tr']:.2f} {unit_len}")
                        k3.metric("n", f"{edu_inputs['n_long']}")
                        k4.metric("Ktr", f"{ktr_aci(edu_inputs['Atr'], to_si_len_from_display(float(edu_inputs['s_tr']), units), int(edu_inputs['n_long'])):.2f} mm")

                st.session_state.calc_inputs["special_seismic_region"] = bool(special_seismic_region)
                main_list = parse_bars_mm(bars_str)
                stir_list = parse_bars_mm(stirrup_str)
                st.session_state.main_list = main_list
                st.session_state.stir_list = stir_list
                if user_mode.startswith("ðŸŽ“"):
                    live_beam_b_mm = edu_inputs.get("beam_b_mm", beam_b_mm)
                    live_beam_h_mm = edu_inputs.get("beam_h_mm", beam_h_mm)
                    live_stirrup_db_mm = edu_inputs.get("stirrup_db_mm", stirrup_db_mm)
                    live_cover_d = edu_inputs.get("cover_d", cover_d)
                    live_spacing_d = edu_inputs.get("spacing_d", spacing_d)
                    live_so_d = edu_inputs.get("so_d", so_d)
                    live_epoxy = edu_inputs.get("epoxy", epoxy)
                    live_top_bar = edu_inputs.get("top_bar", top_bar)
                else:
                    live_beam_b_mm = beam_b_mm
                    live_beam_h_mm = beam_h_mm
                    live_stirrup_db_mm = stirrup_db_mm
                    live_cover_d = cover_d
                    live_spacing_d = spacing_d
                    live_so_d = so_d
                    live_epoxy = epoxy
                    live_top_bar = top_bar
                live_geometry_inputs = dict(st.session_state.get("calc_inputs", {}) or {})
                live_geometry_inputs.update({
                    "beam_b_mm": float(live_beam_b_mm),
                    "beam_h_mm": float(live_beam_h_mm),
                    "stirrup_db_mm": float(live_stirrup_db_mm),
                    "cover_d": live_cover_d,
                    "spacing_d": live_spacing_d,
                    "so_d": live_so_d,
                    "cover_mm": to_si_len_from_display(live_cover_d, units),
                    "spacing_mm": to_si_len_from_display(live_spacing_d, units),
                    "so_mm": to_si_len_from_display(live_so_d, units),
                    "db_mm": float(main_list[0]) if (main_list and isinstance(main_list[0], (int, float))) else None,
                    "epoxy": bool(live_epoxy),
                    "top_bar": bool(live_top_bar),
                    "units": units,
                })
                st.session_state["geometry_live_inputs"] = live_geometry_inputs
                # ---- INPUT SNAPSHOT (for reporting / downstream blocks) ----
                # Store base inputs in session_state.calc_inputs to avoid undefined locals later.
                db_mm = float(main_list[0]) if (main_list and isinstance(main_list[0], (int, float))) else None


            submitted = st.button("Calculate / Update", key="btn_calc_update", on_click=_on_submit_main_form)

            # Persist calculation snapshot so results remain visible after the submit-run.
            if submitted:
                # In educational mode, read values from edu_inputs.
                if user_mode.startswith("🎓"):
                    beam_b_mm = edu_inputs["beam_b_mm"]
                    beam_h_mm = edu_inputs["beam_h_mm"]
                    stirrup_db_mm = edu_inputs["stirrup_db_mm"]
                    cover_d = edu_inputs["cover_d"]
                    spacing_d = edu_inputs["spacing_d"]
                    so_d = edu_inputs["so_d"]
                    epoxy = edu_inputs["epoxy"]
                    top_bar = edu_inputs["top_bar"]
                    lam = edu_inputs["lam"]
                    hook_angle = edu_inputs["hook_angle"]
                    hook_conf_mode = edu_inputs["hook_conf_mode"]
                    hook_conf_manual = edu_inputs["hook_conf_manual"]
                    s_hook_d = edu_inputs["s_hook_d"]
                    # FIX12N: Ktr values are not taken from educational-mode widgets.
                    # Single source of truth is Geometry Main tab / calc_inputs.
                    _ktr_source_submit = st.session_state.get("calc_inputs", {}) or {}
                    Atr = float(_ktr_source_submit.get("Atr_mm2", 0.0) or 0.0)
                    s_tr = from_si_len_to_display(float(_ktr_source_submit.get("s_tr_mm", 0.0) or 0.0), units)
                    n_long = int(_ktr_source_submit.get("n_long", 1) or 1)
                # Unit conversions
                cover = to_si_len_from_display(cover_d, units)
                spacing = to_si_len_from_display(spacing_d, units)
                so = to_si_len_from_display(so_d, units)
                fc = to_si_stress_from_display(fc_d, units)
                fy = to_si_stress_from_display(fy_d, units)
                db = db_mm if (db_mm is not None) else 0.0

                st.session_state.calc_inputs.update({
                    'fc_d': fc_d,
                    'fy_d': fy_d,
                    'cover_d': cover_d,
                    'spacing_d': spacing_d,
                    'so_d': so_d,
                    'beam_b_mm': float(beam_b_mm),
                    'beam_h_mm': float(beam_h_mm),
                    'stirrup_db_mm': float(stirrup_db_mm),
                    'lam': lam,
                    'epoxy': epoxy,
                    'units': units,
                    'edition': edition,
                    'user_mode': user_mode,
                    'ld_method': ld_method,
                    'db_mm': db_mm,
                    'cover_mm': cover,
                    'spacing_mm': spacing,
                    'so_mm': so,
                    'fc_mpa': fc,
                    'fy_mpa': fy,
                    'hook_conf_mode': hook_conf_mode,
                    'hook_conf_manual': hook_conf_manual,
                    's_hook_d': s_hook_d,
                    's_hook_mm': to_si_len_from_display(s_hook_d, units),
                    'Atr_mm2': float(Atr),
                    's_tr_mm': to_si_len_from_display(float(s_tr), units),
                    'n_long': int(n_long),
                    'top_bar': top_bar,
                })

                # Normalize snapshot to guarantee a complete key-set for reports/exports
                st.session_state.calc_inputs = normalize_calc_inputs(st.session_state.calc_inputs, enforce_required=True)
                st.session_state.calc_inputs = derive_inputs(st.session_state.calc_inputs)

                st.session_state["calc_ready"] = True
            else:
                # Reuse last snapshot so outputs stay on screen even without re-submitting.
                snap = st.session_state.get("calc_inputs", {}) or {}
                if not snap or snap.get("cover_mm") is None:
                    st.session_state["calc_ready"] = False
                else:
                    cover = float(snap["cover_mm"])
                    spacing = float(snap["spacing_mm"])
                    so = float(snap["so_mm"])
                    fc = float(snap["fc_mpa"])
                    fy = float(snap["fy_mpa"])
                    _ml = st.session_state.get("main_list", []) or []
                    if snap.get("db_mm") is not None:
                        db = float(snap["db_mm"])
                    elif _ml:
                        db = float(_ml[0])
                    else:
                        db = 0.0
                    st.session_state["calc_ready"] = True

            # ----- Sync Atr, s_tr, n_long from the locked snapshot for all modes -----
            # Streamlit reruns after tab/report interactions. Without this sync, non-educational
            # modes fell back to the local defaults Atr=0, s_tr=0, n_long=1 and overwrote Ktr.
            snap_vals = st.session_state.get("calc_inputs", {})
            if snap_vals:
                Atr = float(snap_vals.get("Atr_mm2", Atr) or 0.0)
                s_tr_mm_val = float(snap_vals.get("s_tr_mm", 0.0) or 0.0)
                s_tr = from_si_len_to_display(s_tr_mm_val, units)
                n_long = int(snap_vals.get("n_long", n_long) or 1)
            # -------------------------------------------------------------

            # -----------------------------------------------------------
            snap = st.session_state.get("calc_inputs", {}) or {}
            has_snapshot = snap.get("cover_mm") is not None and snap.get("fc_mpa") is not None
            should_run = bool(submitted) or bool(has_snapshot)
            if not should_run:
                st.info("Set inputs and press 'Calculate / Update' to run calculations.")
                st.session_state["calc_triggered"] = False
                return
            # -----------------------------------------------------------
            unit_tag = "mm" if units == "SI" else "cm"
            step = 50.0 if units == "SI" else 5.0

            def disp_round(mm: float) -> float:
                v = from_si_len_to_display(mm, units)
                if not apply_round:
                    return round(v, 1)
                if rounding_mode == "UP to 50":
                    return math.ceil(v / step) * step
                return round(v / step) * step

            def get_diff_str(gen: float, simp: float) -> str:
                diff = gen - simp
                d = disp_round(diff)
                if diff < -1.0:
                    return f"-{abs(d)}"
                elif diff > 1.0:
                    return f"+{d}"
                return "-"

            res_rows: List[Dict[str, Any]] = []
            bosb_main_rows: List[Dict[str, Any]] = []

            main_list = st.session_state.get("main_list", None)
            if not main_list:
                try:
                    main_list = parse_bars_mm(bars_str)
                except Exception:
                    main_list = []
            st.session_state.main_list = main_list

            stir_list = st.session_state.get("stir_list", None)
            if not stir_list:
                try:
                    stir_list = parse_bars_mm(stirrup_str)
                except Exception:
                    stir_list = []
            st.session_state.stir_list = stir_list

            for db in main_list:
                psi_s = psi_s_tension(db)
                psi_e_val = psi_e_tension(epoxy, cover, spacing, db)
                if str(user_mode).startswith("🎓"):
                    edu_local = st.session_state.get("edu_modifiers", {})
                    if isinstance(edu_local, dict) and edu_local:
                        try:
                            tbl_25425 = get_ch25_tables_for_edition(edition)["25.4.2.5"]
                            ps_mode = str(edu_local.get("psi_s_mode", "Auto (from db)"))
                            if ps_mode.startswith("Force small"):
                                psi_s = float(tbl_25425["psi_s"]["no19_and_smaller_or_wire"])
                            elif ps_mode.startswith("Force large"):
                                psi_s = float(tbl_25425["psi_s"]["no22_and_larger"])
                            pe_mode = str(edu_local.get("psi_e_mode", "Auto (from epoxy + cover/spacing)"))
                            if epoxy and pe_mode.startswith("Force epoxy-other"):
                                psi_e_val = float(tbl_25425["psi_e"]["epoxy_other"])
                            elif epoxy and pe_mode.startswith("Force epoxy-low"):
                                psi_e_val = float(tbl_25425["psi_e"]["epoxy_low_cover_or_spacing"])
                        except Exception:
                            pass
                try:
                    # Use the same locked snapshot used by reports/exports.
                    # This prevents Ktr from being reset to zero on Streamlit reruns.
                    _ktr_snap = st.session_state.get("calc_inputs", {}) or {}
                    Atr_mm2 = float(_ktr_snap.get("Atr_mm2", Atr) or 0.0)
                    s_tr_mm = float(_ktr_snap.get("s_tr_mm", to_si_len_from_display(float(s_tr), units)) or 0.0)
                    n_long_eff = int(_ktr_snap.get("n_long", n_long) or 0)
                    Ktr_base = ktr_aci(Atr_mm2, s_tr_mm, n_long_eff)
                except Exception:
                    Ktr_base = 0.0

                cb_val, cb_gov = cb_from_geom(cover, spacing, so, db)

                st.session_state.calc_inputs["cb_mm"] = cb_val
                st.session_state.calc_inputs["cb_gov"] = cb_gov
                try:
                    st.session_state.calc_inputs["Ktr_base"] = Ktr_base
                except Exception:
                    pass
                good_geom = (spacing >= db) and (cover >= db)

                gen_T = ld_general_mm(fc, fy, db, lam, psi_t_T, psi_e_val, psi_s, cb_val, Ktr_base, psi_g_val)
                gen_B = ld_general_mm(fc, fy, db, lam, psi_t_B, psi_e_val, psi_s, cb_val, Ktr_base, psi_g_val)

                st.session_state.calc_inputs["beta_raw"] = gen_T.get("beta_raw", None)
                st.session_state.calc_inputs["beta_eff"] = gen_T.get("beta", None)
                st.session_state.calc_inputs["beta_is_capped"] = gen_T.get("is_capped", None)

                simp_T = ld_simplified_mm(fc, fy, db, lam, psi_t_T, psi_e_val, float(psi_g_val), db <= 19.1, good_geom)
                simp_B = ld_simplified_mm(fc, fy, db, lam, psi_t_B, psi_e_val, float(psi_g_val), db <= 19.1, good_geom)

                ls_Simp_T = max(st.session_state.alpha_used * simp_T["ld"], LD_MIN_MM)
                ls_Simp_B = max(st.session_state.alpha_used * simp_B["ld"], LD_MIN_MM)
                ls_Gen_T = max(st.session_state.alpha_used * gen_T["ld"], LD_MIN_MM)
                ls_Gen_B = max(st.session_state.alpha_used * gen_B["ld"], LD_MIN_MM)

                is_simplified = str(ld_method).startswith("Simplified method")
                is_both = str(ld_method).startswith("Both methods")

                if is_both:
                    ld_final_T = max(simp_T["ld"], gen_T["ld"])
                    ls_final_T = max(ls_Simp_T, ls_Gen_T)
                    ld_final_B = max(simp_B["ld"], gen_B["ld"])
                    ls_final_B = max(ls_Simp_B, ls_Gen_B)
                    gov_T = "General" if gen_T["ld"] >= simp_T["ld"] else "Simplified"
                    gov_B = "General" if gen_B["ld"] >= simp_B["ld"] else "Simplified"
                else:
                    ld_final_T = simp_T["ld"] if is_simplified else gen_T["ld"]
                    ls_final_T = ls_Simp_T if is_simplified else ls_Gen_T
                    ld_final_B = simp_B["ld"] if is_simplified else gen_B["ld"]
                    ls_final_B = ls_Simp_B if is_simplified else ls_Gen_B
                    gov_T = "Simplified" if is_simplified else "General"
                    gov_B = "Simplified" if is_simplified else "General"

                splice_status = "CALCULATED"
                splice_note = ""
                if bool(special_seismic_region):
                    splice_status = "NOT REPORTED"
                    splice_note = ("Tension lap splices may be restricted in special seismic detailing regions; "
                                  "refer to ACI 318 Chapter 18.")

                hook_status = "CALCULATED"
                hook_note = ""
                if bool(special_seismic_region):
                    hook_status = "NOT REPORTED"
                    hook_note = ("Standard hooks may be restricted in special seismic detailing regions; "
                                "refer to ACI 318 Chapter 18.")
                    hk = {"ldh": float("nan")}
                else:
                    hk = ldh_hook_mm(fc, fy, db, lam, epoxy, edition, has_conf, cover_ok, False, hook_angle)
                ldt_vals = ldt_headed_mm(fc, fy, db, epoxy, edition, psi_p_h, psi_o_h, psi_c_h) if enable_headed else {}
                geo = calculate_exact_hook_geometry(db, "main")
                ldh_disp = "—" if hook_status == "NOT REPORTED" else disp_round(hk["ldh"])
                psi_r_local = 0.75 if has_conf else 1.0
                ldc_sel = ldc_compression_mm(fc, fy, db, lam, psi_r_local, edition=edition)
                ldc_disp = disp_round(ldc_sel["ldc"])
                lst_gov = max(float(ls_final_T), float(ls_final_B))
                lsc_sel = lap_splice_length_compression_mm(fc, fy, db, lst_gov, edition=edition)
                lsc_val = lsc_sel.get("lsc")
                lsc_disp = disp_round(lsc_val) if isinstance(lsc_val, (int, float)) else str(lsc_val)

                row: Dict[str, Any] = {
                    "db": round(db, 1),
                    "Class": (class_val or st.session_state.get("class_val", "Class B")),
                    f"ℓd_T ({unit_tag})": disp_round(ld_final_T),
                    f"ℓd_B ({unit_tag})": disp_round(ld_final_B),
                    f"ℓdc ({unit_tag})": ldc_disp,
                    f"ℓs_T ({unit_tag})": disp_round(ls_final_T),
                    f"ℓs_B ({unit_tag})": disp_round(ls_final_B),
                    f"ℓs_c ({unit_tag})": lsc_disp,
                    f"ℓdh ({unit_tag})": ldh_disp,
                    "Hook status": hook_status,
                    "Hook note/ref": hook_note,
                    "Splice status": splice_status,
                    "Splice note/ref": splice_note,
                }
                row["cb"] = round(from_si_len_to_display(cb_val, units), 1)
                row["gov"] = cb_gov
                row["(c_b+K_tr)/d_b"] = round(gen_T["beta_raw"], 2)
                row["cap @ 2.5?"] = "YES" if gen_T["is_capped"] else "-"

                if "compare" in str(ld_method):
                    row["Simp_ℓd_T"] = disp_round(simp_T["ld"])
                    row["Gen_ℓd_T"] = disp_round(gen_T["ld"])
                    row["Δ_ℓd_T"] = get_diff_str(gen_T["ld"], simp_T["ld"])
                    row["Simp_ℓs_T"] = disp_round(ls_Simp_T)
                    row["Gen_ℓs_T"] = disp_round(ls_Gen_T)
                    row["Δ_ℓs_T"] = get_diff_str(ls_Gen_T, ls_Simp_T)
                    row["Simp_ℓd_B"] = disp_round(simp_B["ld"])
                    row["Gen_ℓd_B"] = disp_round(gen_B["ld"])
                    row["Δ_ℓd_B"] = get_diff_str(gen_B["ld"], simp_B["ld"])
                    row["Simp_ℓs_B"] = disp_round(ls_Simp_B)
                    row["Gen_ℓs_B"] = disp_round(ls_Gen_B)
                    row["Δ_ℓs_B"] = get_diff_str(ls_Gen_B, ls_Simp_B)
                elif str(ld_method) == "Both":
                    row[f"ℓd_T_Simp ({unit_tag})"] = disp_round(simp_T["ld"])
                    row[f"ℓd_T_Gen ({unit_tag})"] = disp_round(gen_T["ld"])
                    row[f"ℓd_T_Final ({unit_tag})"] = disp_round(ld_final_T)
                    row["Gov_T"] = gov_T
                    row[f"ℓs_T_Simp ({unit_tag})"] = disp_round(ls_Simp_T)
                    row[f"ℓs_T_Gen ({unit_tag})"] = disp_round(ls_Gen_T)
                    row[f"ℓs_T_Final ({unit_tag})"] = disp_round(ls_final_T)
                    row[f"ℓd_B_Simp ({unit_tag})"] = disp_round(simp_B["ld"])
                    row[f"ℓd_B_Gen ({unit_tag})"] = disp_round(gen_B["ld"])
                    row[f"ℓd_B_Final ({unit_tag})"] = disp_round(ld_final_B)
                    row["Gov_B"] = gov_B
                    row[f"ℓs_B_Simp ({unit_tag})"] = disp_round(ls_Simp_B)
                    row[f"ℓs_B_Gen ({unit_tag})"] = disp_round(ls_Gen_B)
                    row[f"ℓs_B_Final ({unit_tag})"] = disp_round(ls_final_B)
                elif "General" in str(ld_method):
                    row[f"ℓd_T ({unit_tag})"] = disp_round(ld_final_T)
                    row[f"ℓs_T ({unit_tag})"] = disp_round(ls_final_T)
                    row[f"ℓd_B ({unit_tag})"] = disp_round(ld_final_B)
                    row[f"ℓs_B ({unit_tag})"] = disp_round(ls_final_B)
                else:
                    for _k in ("cb", "gov", "conf.term", "capped", "(c_b+K_tr)/d_b", "cap @ 2.5?"):
                        row.pop(_k, None)
                    row[f"ℓd_T ({unit_tag})"] = disp_round(ld_final_T)
                    row[f"ℓs_T ({unit_tag})"] = disp_round(ls_final_T)
                    row[f"ℓd_B ({unit_tag})"] = disp_round(ld_final_B)
                    row[f"ℓs_B ({unit_tag})"] = disp_round(ls_final_B)

                if splice_status == "NOT REPORTED":
                    for _k in list(row.keys()):
                        if "ℓs" in str(_k):
                            row[_k] = "—"
                if enable_headed:
                    row[f"ℓdt ({unit_tag})"] = disp_round(ldt_vals["ldt"])
                res_rows.append(row)

                bosb_main_rows.append({
                    "db": db,
                    "Bend D": int(geo["D_mm"]),
                    "L1 (90°)": disp_round(geo["L1_total"]),
                    "Ext 90": round(geo["L1_ext"]),
                    "L2 (135°)": disp_round(geo["L2_total"]),
                    "Ext 135": round(geo["L2_ext"]),
                    "L3 (180°)": disp_round(geo["L3_total"]),
                    "Ext 180": round(geo["L3_ext"]),
                })

            bosb_stir_rows: List[Dict[str, Any]] = []
            for db in stir_list:
                geo = calculate_exact_hook_geometry(db, "stirrup")
                bosb_stir_rows.append({
                    "db": db,
                    "Bend D": int(geo["D_mm"]),
                    "L1 (90°)": disp_round(geo["L1_total"]),
                    "Ext 90": round(geo["L1_ext"]),
                    "L2 (135°)": disp_round(geo["L2_total"]),
                    "Ext 135": round(geo["L2_ext"]),
                    "L3 (180°)": disp_round(geo["L3_total"]),
                    "Ext 180": round(geo["L3_ext"]),
                })

            st.session_state.df_res = pd.DataFrame(res_rows)
            st.session_state.qa_messages = build_engineering_qa_messages(
                st.session_state.get("calc_inputs", {}) or {},
                st.session_state.df_res,
                units,
                edition,
            )
            st.session_state.df_bosb_main = pd.DataFrame(bosb_main_rows)
            st.session_state.df_bosb_stir = pd.DataFrame(bosb_stir_rows)
            st.session_state.data_calculated = True

            try:
                st.session_state.excel_blob = generate_excel_blob(
                    st.session_state.df_res,
                    st.session_state.df_bosb_main,
                    st.session_state.df_bosb_stir,
                )
                _run_id = st.session_state.get("calc_run_id", BUILD_ID)
                _inputs_for_name = st.session_state.get("calc_inputs", {}) or {}
                st.session_state["excel_filename"] = make_artifact_filename("Excel", "xlsx", _inputs_for_name, _run_id)
            except Exception as e:
                st.session_state.excel_blob = None
                try:
                    st.session_state.export_errors["excel"] = f"{type(e).__name__}: {e}"
                except Exception:
                    pass
                st.warning(f"Excel export is not available ({type(e).__name__}: {e}). Install 'xlsxwriter' or 'openpyxl' to enable Excel download.")

            try:
                _inputs_for_report = st.session_state.get("calc_inputs", {}) or {}
                _df_for_report = st.session_state.get("df_res", None)
                if _df_for_report is not None:
                    st.session_state.html_report = generate_html_report(_inputs_for_report, _df_for_report)
                    _run_id = st.session_state.get("calc_run_id", BUILD_ID)
                    st.session_state.html_report_filename = make_artifact_filename("Report", "html", _inputs_for_report, _run_id)
                else:
                    st.session_state.html_report = None
                    st.session_state.html_report_filename = None
            except Exception as e:
                st.session_state.html_report = None
                st.session_state.html_report_filename = None
                try:
                    st.session_state.export_errors["html"] = f"{type(e).__name__}: {e}"
                except Exception:
                    pass



    # (نتایج خلاصه اکنون در نوار بالای فرم نمایش داده می‌شود — _render_results_topbar)

    # =======================================================================
    #  OUTPUT TABS — full width, below the inputs
    # =======================================================================
    st.divider()
    if st.session_state.data_calculated:
            df_res = st.session_state.df_res
            st.subheader(f"Detailed results ({units})")
            render_professional_run_header(
                edition,
                units,
                st.session_state.get("calc_inputs", {}) or {},
                st.session_state.get("qa_messages", []),
            )
            is_edu = user_mode.startswith("🎓")
            tab_names = ["**Lengths**"]
            if is_edu:
                tab_names += ["**Detailed Report**", "**Chart Analysis**"]
            tab_names += [
                "**Numbers Table**",
                "**Engineering QA**",
                "**Calc Transparency**",
                "**Geometry Main**",
                "**Geometry Stirrup**",
                "**Excel**",
                "**Print Report**",
            ]
            tabs = st.tabs(tab_names)

            def _tab(label: str):
                try:
                    return tabs[tab_names.index(label)]
                except Exception:
                    return None

            t_len = _tab("**Lengths**")
            t_num = _tab("**Numbers Table**")
            t_qa = _tab("**Engineering QA**")
            t_calc = _tab("**Calc Transparency**")
            t_rpt = _tab("**Detailed Report**") if is_edu else None
            t_cht = _tab("**Chart Analysis**") if is_edu else None
            t_geo_m = _tab("**Geometry Main**")
            t_geo_s = _tab("**Geometry Stirrup**")
            t_xls = _tab("**Excel**")
            t_prt = _tab("**Print Report**")

            def _summary_df(df: 'pd.DataFrame') -> 'pd.DataFrame':
                if df is None or getattr(df, 'empty', True):
                    return df
                keep = []
                for c in df.columns:
                    cs = str(c)
                    if cs in ('db', 'Class', 'Hook status', 'Hook note/ref', 'Splice status', 'Splice note/ref'):
                        keep.append(c)
                        continue
                    if ('ℓd' in cs) or ('ℓs' in cs) or ('ℓdh' in cs) or ('ℓdc' in cs) or ('ℓdt' in cs):
                        if cs.startswith(('Simp_', 'Gen_', 'Δ_')):
                            continue
                        keep.append(c)
                if keep:
                    keep = [c for c in keep if c in df.columns]
                    return df.loc[:, keep]
                return df

            with t_len:
                # FIX12B_LENGTHS_TABLE_ONLY_NOTE
                st.caption("Table-only view: this tab intentionally shows the full multi-bar length table only. Single-bar details and β/Ktr/cb breakdown are available in Detailed Report, Chart Analysis, Geometry Main, Calc Transparency, and Engineering QA.")
                st.dataframe(_arrow_safe_df(_summary_df(st.session_state.df_res)), hide_index=True, width=1600)

            if t_num is not None:
                with t_num:
                    st.markdown("### Numbers table (edition-aware)")
                    st.markdown(build_aci_parameter_reference_md(edition, units))
                    ci = st.session_state.get("calc_inputs", {}) or {}
                    if isinstance(ci, dict) and len(ci) > 0:
                        rows = [{"Parameter": k, "Last used value": str(v)} for k, v in ci.items()]
                        st.dataframe(_arrow_safe_df(pd.DataFrame(rows)), width='stretch', hide_index=True)
                    else:
                        st.info("No calculation snapshot is available yet.")


            # FIX11A_RENDER_CALC_TRANSPARENCY_BLOCK
            if 't_calc' in locals() and t_calc is not None:
                with t_calc:
                    render_calc_transparency_report(
                        st.session_state.get("calc_inputs", {}) or {},
                        st.session_state.get("df_res", None),
                        st.session_state.get("qa_messages", []),
                        units,
                        edition,
                    )

            if t_qa is not None:
                with t_qa:
                    st.markdown("### Engineering QA / Warning Panel")
                    st.caption("Centralized QA messages generated from the same calculation snapshot and result table.")
                    qa_msgs = st.session_state.get("qa_messages", [])
                    render_engineering_qa_panel(qa_msgs)
                    st.markdown("---")
                    st.markdown("#### Input Snapshot Used by QA")
                    ci = st.session_state.get("calc_inputs", {}) or {}
                    if isinstance(ci, dict) and ci:
                        st.dataframe(
                            _arrow_safe_df(pd.DataFrame([{"Parameter": k, "Value": str(v)} for k, v in ci.items()])),
                            width="stretch",
                            hide_index=True,
                        )
                    else:
                        st.info("No calculation input snapshot is available yet.")

            if is_edu and t_rpt is not None and st.session_state.main_list:
                with t_rpt:
                    st.markdown("### Detailed calculation report")
                    compare_mode = False
                    show_diff_only = True
                    append_diff = bool(st.session_state.get("ch25_append_full_edition_comparison", False))
                    other_edition = "ACI318-2019" if edition == "ACI318-2025" else "ACI318-2025"

                    # No duplicate input expander here – inputs are now in the main area.
                    # Only show the report content.

                    _main_list = st.session_state.get("main_list", []) or []
                    if not _main_list:
                        st.warning("No main bar list is available yet.")
                        st.stop()
                    det_db = shared_single_bar_selectbox("Select bar diameter for report (mm)", _main_list, key="selected_bar_detail_report_mm")
                    if det_db is not None:
                        st.caption(f"Synced selected bar across single-bar tabs: {float(det_db):.0f} mm")
                    if det_db is None:
                        st.warning("Please select a bar diameter for report.")
                        st.stop()
                    inputs = normalize_calc_inputs(st.session_state.get("calc_inputs", {}), enforce_required=True)
                    st.session_state.calc_inputs = inputs
                    fc = to_si_stress_from_display(inputs["fc_d"], inputs["units"])
                    fy = to_si_stress_from_display(inputs["fy_d"], inputs["units"])
                    psi_g_val = float(inputs.get("psi_g_val", 1.0))
                    cover = to_si_len_from_display(inputs["cover_d"], inputs["units"])
                    spacing = to_si_len_from_display(inputs["spacing_d"], inputs["units"])
                    so = to_si_len_from_display(inputs["so_d"], inputs["units"])
                    lam = inputs["lam"]
                    epoxy = inputs["epoxy"]
                    psi_g_val = inputs.get("psi_g_val", 1.0)
                    # Report must use the same Ktr source as the calculation snapshot.
                    Ktr_base = float(inputs.get("Ktr_base") or 0.0)
                    if Ktr_base <= 0.0:
                        Ktr_base = ktr_aci(
                            float(inputs.get("Atr_mm2", 0.0) or 0.0),
                            float(inputs.get("s_tr_mm", 0.0) or 0.0),
                            int(inputs.get("n_long", 1) or 1),
                        )

                    d_psi_s = psi_s_tension(det_db)
                    d_psi_e = psi_e_tension(epoxy, cover, spacing, det_db)
                    d_cb, d_gov = cb_from_geom(cover, spacing, so, det_db)
                    d_good = (spacing >= det_db) and (cover >= det_db)
                    d_gen = ld_general_mm(fc, fy, det_db, lam, psi_t_T, d_psi_e, d_psi_s, d_cb, Ktr_base, psi_g_val)
                    d_simp = ld_simplified_mm(fc, fy, det_db, lam, psi_t_T, d_psi_e, float(psi_g_val), det_db <= 19.1, d_good)

                    if not ((locals().get('compare_mode', False)) and (locals().get('show_diff_only', False))):
                        st.markdown(f"#### 1. Straight development length ℓd for bar {det_db:.0f} mm")
                        st.write("**Used parameters (ACI-based):**")
                        psi_g_src = str(inputs.get("psi_g_source") or "")
                        psi_g_note_latex = ""
                        if "nearest Grade" in psi_g_src:
                            psi_g_note_latex = r"\;(\text{nearest})"
                        st.latex(
                            rf"\psi_t={psi_t_T:.2f},\; "
                            rf"\psi_e={d_psi_e:.2f},\; "
                            rf"\psi_s={d_psi_s:.2f},\; "
                            rf"\psi_g={psi_g_val:.2f}{psi_g_note_latex},\; "
                            rf"\lambda={lam:.2f},\; "
                            rf"f'_c={inputs['fc_d']:.1f},\; f_y={inputs['fy_d']:.1f}"
                        )

                        ld_method_local = str(inputs.get('ld_method', ld_method))
                        is_compare_mode = ('compare' in ld_method_local.lower()) or ('both' in ld_method_local.lower())
                        is_simplified_mode = ('simplified' in ld_method_local.lower()) and (not is_compare_mode)

                        def _render_simplified_block():
                            st.markdown("**1.2 Simplified ACI formula (Table 25.4.2.3)**")
                            if d_good:
                                st.write("Spacing and cover satisfy **good bond conditions**")
                                st.markdown(f"**Denominator factor (α):** {d_simp['denom_factor']:.2f}")
                            else:
                                st.write("Spacing and/or cover do **not** satisfy good bond conditions.")
                                st.markdown(f"**Denominator factor (α):** {d_simp['denom_factor']:.2f}")
                            _num_txt = "" if abs(d_simp.get("num_factor", 1.0) - 1.0) < 1e-9 else r"1.5\, "
                            st.latex(
                                r"\ell_d = \dfrac{" + _num_txt + r"f_y\,\psi_t\psi_e\psi_g}{"
                                + f"{d_simp['denom_factor']:.1f}" + r"\,\lambda\,\sqrt{f'_c}}\, d_b"
                            )
                            st.markdown(
                                "Numeric substitution:"
                                + (f" ℓd,raw = (1.5 × {fy:.1f} / ({d_simp['denom_factor']:.2f} × {lam:.2f} × √{fc:.1f})) × "
                                   if abs(d_simp.get('num_factor', 1.0) - 1.0) > 1e-9
                                   else f" ℓd,raw = ({fy:.1f} / ({d_simp['denom_factor']:.2f} × {lam:.2f} × √{fc:.1f})) × ")
                                + f"({psi_t_T:.2f} × {d_psi_e:.2f} × {psi_g_val:.2f}) × {det_db:.1f} = {d_simp['raw']:.1f} mm"
                            )
                            st.write(f"Apply ACI minimum 300 mm → ℓd = {d_simp['ld']:.1f} mm")

                        def _render_general_block():
                            st.markdown("**1.1 General ACI formula (Eq. (25.4.2.4a))**")
                            st.latex(
                                r"\ell_d = \dfrac{f_y\,\psi_t\,\psi_e\,\psi_s\,\psi_g}{1.1\,\lambda\,\sqrt{f'_c}}"
                                r"\,\dfrac{d_b}{(c_b + K_{tr})/d_b}"
                            )
                            st.markdown(
                                "Numeric substitution:"
                                f" ℓd,raw = ({fy:.1f} / (1.1 × {lam:.2f} × √{fc:.1f})) × "
                                f"(({psi_t_T:.2f} × {d_psi_e:.2f} × {d_psi_s:.2f} × {psi_g_val:.2f}) / (({d_gen['cb']:.1f} + {d_gen['Ktr']:.1f}) / {det_db:.1f})) × "
                                f"{det_db:.1f} = {d_gen['raw']:.1f} mm"
                            )
                            st.markdown("**Ktr (Transverse Reinforcement Index):**")
                            Atr_val = float(inputs.get("Atr_mm2", 0.0) or inputs.get("Atr", 0.0))
                            s_val = float(inputs.get("s_tr_mm", 0.0) or inputs.get("s_tr", 0.0))
                            n_val = int(inputs.get("n_long", 1))
                            st.latex(r"K_{tr} = \frac{40\,A_{tr}}{s\,n}")
                            if Atr_val > 0 and s_val > 0 and n_val > 0:
                                st.markdown(
                                    f"With input values:\n"
                                    f"- $A_{{tr}} = {Atr_val:.1f}$ mm²\n"
                                    f"- $s = {s_val:.1f}$ mm\n"
                                    f"- $n = {n_val}$\n\n"
                                    f"Calculation:\n"
                                    f"$K_{{tr}} = 40 \\times {Atr_val:.1f} \\,/\\, ({s_val:.1f} \\times {n_val}) = {Ktr_base:.2f}$"
                                )
                            else:
                                st.markdown("Transverse reinforcement data (Atr, s, n) not provided; $K_{tr} = 0$ is assumed.")
                            st.caption("*(The constant 40 has units 1/mm and is applicable only in SI.)*")
                            st.write(f"β = (c_b + Ktr)/db = ({d_gen['cb']:.1f} + {d_gen['Ktr']:.1f}) / {det_db:.1f} = {d_gen['beta_raw']:.3f}")
                            if d_gen.get("is_capped"):
                                st.warning("β was capped to 2.5 per ACI limit.")
                            if str(user_mode).startswith("🎓"):
                                st.markdown('<div class="aci-section-title">β breakdown – (c_b + Ktr) / d_b</div>', unsafe_allow_html=True)
                                with st.container(border=True):
                                    _db = float(det_db) if det_db is not None else None
                                    _cover = cover if cover is not None else None
                                    _spacing = spacing if spacing is not None else None
                                    _so = so if so is not None else None
                                    _Ktr = float(d_gen.get("Ktr", 0.0) or 0.0)
                                    cb_candidates = {}
                                    if _cover is not None:
                                        cb_candidates["cover"] = float(_cover)
                                    if _spacing is not None:
                                        cb_candidates["spacing/2"] = float(_spacing) / 2.0
                                    if _so is not None and float(_so) > 0.0:
                                        cb_candidates["s_o/2"] = float(_so) / 2.0
                                    cb_calc = None
                                    cb_key = None
                                    if cb_candidates:
                                        cb_key = min(cb_candidates, key=lambda k: cb_candidates[k])
                                        cb_calc = cb_candidates[cb_key]
                                    actual_cb = d_cb
                                    beta_raw_local = d_gen.get("beta_raw", None)
                                    beta_eff_local = d_gen.get("beta_eff", None)
                                    beta_cap = 2.5
                                    beta_is_capped = bool(d_gen.get("is_capped")) if "is_capped" in d_gen else bool(inputs.get("beta_is_capped"))
                                    st.markdown('<div dir="ltr">', unsafe_allow_html=True)
                                    st.markdown(
                                        """- **β** is the ratio of the “geometric concrete confinement/anchorage capacity around the bar” to the bar diameter.
- Increasing **c_b** (effective concrete cover/spacing) and/or **K<sub>tr</sub>** (transverse reinforcement confinement) generally increases β.
- In this tool, **c_b** is taken as the **minimum** of the available options (cover, half clear spacing, and if provided, s<sub>o</sub>/2).""",
                                        unsafe_allow_html=True,
                                    )
                                    st.markdown("</div>", unsafe_allow_html=True)
                                    st.write("**Step 1 — c_b candidates:**")
                                    if cb_candidates:
                                        for k, v in cb_candidates.items():
                                            st.write(f"- {k} = {v:.1f} mm")
                                        st.write(f"✅ Governing c_b = {actual_cb:.1f} mm  ({d_gov})")
                                    else:
                                        st.warning("c_b candidates are not available (missing cover/spacing inputs).")
                                    st.write("**Step 2 — Ktr:**")
                                    st.write(f"Ktr = {(_Ktr):.1f} mm")
                                    st.write("**Step 3 — β:**")
                                    if _db and actual_cb is not None:
                                        beta_raw_calc = (actual_cb + _Ktr) / _db
                                        st.latex(rf"\beta_{{raw}} = \dfrac{{c_b + K_{{tr}}}}{{d_b}} = \dfrac{{{actual_cb:.1f} + {_Ktr:.1f}}}{{{_db:.1f}}} = {beta_raw_calc:.3f}")
                                    if beta_raw_local is not None:
                                        st.write(f"β(raw) used in calculation = {float(beta_raw_local):.3f}")
                                    if beta_eff_local is not None:
                                        st.write(f"β(effective) used in calculation = {float(beta_eff_local):.3f}")
                                    if beta_is_capped:
                                        st.warning(f"β was capped: β(effective) = min(β(raw), {beta_cap})")
                            st.write(f"Apply ACI minimum 300 mm → ℓd = {d_gen['ld']:.1f} mm")

                        if is_compare_mode and ((not compare_mode) or (compare_mode and not show_diff_only)):
                            st.markdown("##### 1.1–1.2 Both methods side-by-side (General vs Simplified)")
                            col_g, col_s = st.columns(2)
                            with col_g:
                                st.markdown("**General method**")
                                _render_general_block()
                            with col_s:
                                st.markdown("**Simplified method**")
                                _render_simplified_block()
                        elif is_simplified_mode:
                            _render_simplified_block()
                        else:
                            _render_general_block()

                        st.divider()
                    st.markdown("#### 2. Hooked bar development length")
                    has_conf = inputs.get("has_conf", False)
                    cover_ok = cover >= det_db and spacing >= det_db
                    if False:
                        pass  # removed compare UI
                    else:
                        hk_res = ldh_hook_mm(fc, fy, det_db, lam, epoxy, edition, has_conf, cover_ok, False, inputs.get("hook_angle", 90))
                        st.markdown(f"**{'ACI 318-2019' if edition=='ACI318-2019' else 'ACI 318-2025'} – §25.4.3.1(a)**")
                        if edition == "ACI318-2019":
                            st.latex(r"\ell_{dh} = \max\left[\left(\dfrac{f_y \psi_e \psi_r \psi_o \psi_c}{23 \lambda \sqrt{f'_c}}\right) d_b^{1.5},\; 8d_b,\;150\text{ mm}\right]")
                        else:
                            C_used = float(hk_res.get('C_used', 0.0))
                            exp_db = float(hk_res.get('db_exponent', 1.0))
                            exp_tex = '' if abs(exp_db - 1.0) < 1e-9 else f'^{{{exp_db:.1f}}}'
                            st.latex(rf"\ell_{{dh}} = \max\left[\left(\dfrac{{f_y \psi_e \psi_s \psi_{{cc}} \psi_r}}{{{C_used:.2f} \lambda \sqrt{{f'_c}}}}\right) d_b{exp_tex},\; 8d_b,\;152\text{{ mm}}\right]")
                        if edition == 'ACI318-2019':
                            st.markdown(
                                'Numeric substitution:'
                                f" ℓdh,raw = ( {fy:.1f} × {hk_res['psi_e']:.2f} × "
                                f"{hk_res['psi_r']:.2f} × {hk_res['psi_o']:.2f} × "
                                f"{hk_res['psi_c']:.2f} / "
                                f"({hk_res['C_used']:.0f} × {lam:.2f} × √{fc:.1f}) ) × "
                                f"{det_db:.1f}^1.5 = {hk_res['base']:.1f} mm"
                            )
                        else:
                            st.markdown(
                                'Numeric substitution:'
                                f" ℓdh,raw = ( {fy:.1f} × {hk_res['psi_e']:.2f} × "
                                f"{hk_res['psi_s']:.2f} × {hk_res['psi_cc']:.2f} × {hk_res['psi_r']:.2f} / "
                                f"({hk_res['C_used']:.2f} × {lam:.2f} × √{fc:.1f}) ) × "
                                f"{det_db:.1f} = {hk_res['base']:.1f} mm"
                            )
                        st.markdown(f"8/12 db term = {hk_res['angle_min']:.0f} mm")
                    st.latex(r"\ell_{s} = \alpha \, \ell_{d},\quad \alpha = 1.0 \text{ (Class A)},\; 1.3 \text{ (Class B)}")
                    alpha_used = float(st.session_state.get('alpha_used', 1.0))
                    st.markdown(f"Selected splice class: **Class {st.session_state.get('class_val', 'A')}** → α = {alpha_used:.2f}")
                    st.markdown(f"- ℓs, General = α × ℓd,General = {alpha_used:.2f} × {d_gen['ld']:.0f} = {alpha_used * d_gen['ld']:.0f} mm")
                    st.markdown(f"- ℓs, Simplified = α × ℓd,Simplified = {alpha_used:.2f} × {d_simp['ld']:.0f} = {alpha_used * d_simp['ld']:.0f} mm")

                    if (not compare_mode) or (compare_mode and not show_diff_only):
                        st.divider()
                        st.markdown("#### 4. Main bar hook geometry (L1 / L2 / L3)")
                    g_main = calculate_exact_hook_geometry(det_db, "main")
                    kd_main = g_main["bend_factor"]
                    Dm = g_main["D_mm"]
                    L1_ext = g_main["L1_ext"]
                    L1_tot = g_main["L1_total"]
                    L2_ext = g_main["L2_ext"]
                    L2_tot = g_main["L2_total"]
                    L3_ext = g_main["L3_ext"]
                    L3_tot = g_main["L3_total"]
                    st.markdown("This section computes bend diameter **D** and total hook lengths **L1**, **L2**, **L3** for 90°, 135°, and 180° hooks for main bars.")
                    st.markdown("**4.1 Bend diameter D (ACI 25.3.1, Table 25.3.1)**")
                    st.latex(r"D = k_d\, d_b")
                    st.markdown(f"For db = {det_db:.0f} mm, table gives k_d = {kd_main:.0f} for main bars.")
                    st.markdown(f"Therefore: D = k_d × db = {kd_main:.0f} × {det_db:.0f} = {Dm:.0f} mm.")
                    st.markdown("**4.2 Hook extensions for main bars**")
                    st.markdown(f"90° main bar hook: L_ext,90 = {L1_ext:.0f} mm; L_90 = D/2 + db + L_ext,90 = {Dm/2:.0f} + {det_db:.0f} + {L1_ext:.0f} = {L1_tot:.0f} mm.")
                    st.markdown(f"135° main bar hook: L_ext,135 = {L2_ext:.0f} mm; L_135 = D/2 + db + L_ext,135 = {Dm/2:.0f} + {det_db:.0f} + {L2_ext:.0f} = {L2_tot:.0f} mm.")
                    st.markdown(f"180° main bar hook: L_ext,180 = {L3_ext:.0f} mm; L_180 = D/2 + db + L_ext,180 = {Dm/2:.0f} + {det_db:.0f} + {L3_ext:.0f} = {L3_tot:.0f} mm.")
                    if (not compare_mode) or (compare_mode and not show_diff_only):
                        st.divider()
                        st.markdown("#### 5. Stirrup / tie hook geometry")
                    g_stir = calculate_exact_hook_geometry(det_db, "stirrup")
                    kd_stir = g_stir["bend_factor"]
                    Dst = g_stir["D_mm"]
                    sL1_ext = g_stir["L1_ext"]
                    sL1_tot = g_stir["L1_total"]
                    sL2_ext = g_stir["L2_ext"]
                    sL2_tot = g_stir["L2_total"]
                    sL3_ext = g_stir["L3_ext"]
                    sL3_tot = g_stir["L3_total"]
                    st.markdown("For stirrups and ties, ACI Table 25.3.2 provides bend diameters and hook extensions based on bar size group.")
                    st.markdown("**5.1 Bend diameter for stirrups/ties (Table 25.3.2)**")
                    st.latex(r"D_{\text{stir}} = k_d\, d_b")
                    st.markdown(f"For db = {det_db:.0f} mm, table gives k_d = {kd_stir:.0f} for stirrups/ties.")
                    st.markdown(f"Therefore: D_stir = k_d × db = {kd_stir:.0f} × {det_db:.0f} = {Dst:.0f} mm.")
                    st.markdown("**5.2 Hook extensions for stirrups/ties (Table 25.3.2)**")
                    is_small_stir = det_db <= 16.0
                    if is_small_stir:
                        st.write("- 90° stirrup hook (No.10–16): L_ext,90 = max(6 db, 75 mm).")
                    else:
                        st.write("- 90° stirrup hook (No.19–25): L_ext,90 = 12 db.")
                    st.markdown(f"90° stirrup hook: L_ext,90 = {sL1_ext:.0f} mm; L_90 = {Dst/2:.0f} + {det_db:.0f} + {sL1_ext:.0f} = {sL1_tot:.0f} mm.")
                    st.markdown(f"135° stirrup hook: L_ext,135 = {sL2_ext:.0f} mm; L_135 = {Dst/2:.0f} + {det_db:.0f} + {sL2_ext:.0f} = {sL2_tot:.0f} mm.")
                    st.markdown(f"180° stirrup hook: L_ext,180 = {sL3_ext:.0f} mm; L_180 = {Dst/2:.0f} + {det_db:.0f} + {sL3_ext:.0f} = {sL3_tot:.0f} mm.")

                    if locals().get('append_diff', False):
                        st.divider()
                        st.markdown("### Appendix — Full edition comparison (computed)")
                        st.caption(f"Selected: {edition}  |  Compared against: {other_edition}  |  Items shown: full side-by-side comparison")
                        has_conf_local = inputs.get("has_conf", False)
                        cover_ok_local = (cover >= det_db) and (spacing >= det_db)
                        hk_2019 = ldh_hook_mm(fc, fy, det_db, lam, epoxy, "ACI318-2019", has_conf_local, cover_ok_local, False, inputs.get("hook_angle", 90))
                        hk_2025 = ldh_hook_mm(fc, fy, det_db, lam, epoxy, "ACI318-2025", has_conf_local, cover_ok_local, False, inputs.get("hook_angle", 90))
                        hk_rows = [
                            _compare_row_all("Denominator coefficient C", hk_2019.get("C"), hk_2025.get("C")),
                            _compare_row_all("Equation family", hk_2019.get("equation"), hk_2025.get("equation")),
                            _compare_row_all("Computed ℓdh (before minimums), mm", f"{hk_2019.get('base'):.1f}", f"{hk_2025.get('base'):.1f}"),
                            _compare_row_all("Final ℓdh, mm", f"{hk_2019.get('ldh'):.0f} ({hk_2019.get('governing')})", f"{hk_2025.get('ldh'):.0f} ({hk_2025.get('governing')})"),
                        ]
                        render_diff_only_table("Hooked bar development length ℓdh (ACI §25.4.3.1)", hk_rows, silent_if_empty=False)
                        if enable_headed:
                            hd_sel = ldt_headed_mm(fc, fy, det_db, epoxy, edition, inputs.get("psi_p_h", 1.0), inputs.get("psi_o_h", 1.0), inputs.get("psi_c_h", 1.0))
                            hd_oth = ldt_headed_mm(fc, fy, det_db, epoxy, other_edition, inputs.get("psi_p_h", 1.0), inputs.get("psi_o_h", 1.0), inputs.get("psi_c_h", 1.0))
                            hd_rows = [
                                _compare_row_all("Denominator coefficient C", hd_sel.get("C_used"), hd_oth.get("C_used")),
                                _compare_row_all("Computed ℓdt, mm", f"{hd_sel.get('ldt'):.0f}", f"{hd_oth.get('ldt'):.0f}"),
                            ]
                            render_diff_only_table("Headed bar development length ℓdt (ACI §25.4.4.1)", hd_rows, silent_if_empty=False)
                        psi_r_local = 0.75 if has_conf_local else 1.0
                        c_sel = ldc_compression_mm(fc, fy, det_db, lam, psi_r_local, edition=edition)
                        c_oth = ldc_compression_mm(fc, fy, det_db, lam, psi_r_local, edition=other_edition)
                        c_rows = [
                            _compare_row_all("Coefficient in term (a)", f"{c_sel.get('c_a'):.6f}", f"{c_oth.get('c_a'):.6f}"),
                            _compare_row_all("Term (a), mm", f"{c_sel.get('term_a'):.1f}", f"{c_oth.get('term_a'):.1f}"),
                            _compare_row_all("Final ℓdc, mm", f"{c_sel.get('ldc'):.0f}", f"{c_oth.get('ldc'):.0f}"),
                        ]
                        render_diff_only_table("Compression development length ℓdc (ACI §25.4.9.2)", c_rows, silent_if_empty=False)

            if is_edu and t_cht is not None and st.session_state.main_list:
                with t_cht:
                    y_axis_opt = st.selectbox("Chart Y-axis variable", [
                        "Development Length (Ld Top)", "Development Length (Ld Bottom)",
                        "Splice Length (Ls Top)", "Splice Length (Ls Bottom)",
                        "Hook Length (Ldh)", "Headed Length (Ldt)",
                    ])
                    chart_db = shared_single_bar_selectbox("Select bar diameter for chart (mm)", st.session_state.main_list, key="selected_bar_chart_mm")
                    if chart_db is not None:
                        st.caption(f"Synced selected bar across single-bar tabs: {float(chart_db):.0f} mm")
                    inputs = normalize_calc_inputs(st.session_state.get("calc_inputs", {}), enforce_required=True)
                    st.session_state.calc_inputs = inputs
                    fc = to_si_stress_from_display(inputs["fc_d"], inputs["units"])
                    fy = to_si_stress_from_display(inputs["fy_d"], inputs["units"])
                    cover = to_si_len_from_display(inputs["cover_d"], inputs["units"])
                    lam = inputs["lam"]
                    spacings = list(range(20, 150, 5))
                    y_simp, y_gen = [], []
                    const_hook = ldh_hook_mm(fc, fy, chart_db, lam, inputs["epoxy"], inputs.get("edition", edition), inputs.get("has_conf", False), inputs.get("cover_ok", False), False, inputs.get("hook_angle", 90))["ldh"]
                    const_headed = ldt_headed_mm(fc, fy, chart_db, inputs["epoxy"], inputs.get("edition", edition), inputs.get("psi_p_h", 1.0), inputs.get("psi_o_h", 1.0), inputs.get("psi_c_h", 1.0))["ldt"]
                    for sp in spacings:
                        if "Hook" in y_axis_opt:
                            y_simp.append(const_hook); y_gen.append(const_hook)
                        elif "Headed" in y_axis_opt:
                            y_simp.append(const_headed); y_gen.append(const_headed)
                        else:
                            sp_val = sp if inputs["units"] == "SI" else sp * 10
                            c_sim = cb_from_geom(cover, sp_val, 0, chart_db)
                            good_g = (sp_val >= chart_db and cover >= chart_db)
                            p_psi_t = psi_t_B if "Bottom" in y_axis_opt else psi_t_T
                            p_alpha = st.session_state.alpha_used if "Splice" in y_axis_opt else 1.0
                            s_res = ld_simplified_mm(fc, fy, chart_db, lam, p_psi_t, inputs["epoxy"] and 1.5 or 1.0, 1.0, chart_db <= 19.1, good_g)
                            g_res = ld_general_mm(fc, fy, chart_db, lam, p_psi_t, inputs["epoxy"] and 1.5 or 1.0, 0.8 if chart_db <= 19 else 1.0, c_sim[0], 0.0, psi_g_val)
                            y_simp.append(s_res["ld"] * p_alpha)
                            y_gen.append(g_res["ld"] * p_alpha)
                    fig, ax = plt.subplots(figsize=(9, 4))
                    ax.plot(spacings, y_simp, linestyle="--", linewidth=2, label="Simplified")
                    if "Hook" in y_axis_opt or "Headed" in y_axis_opt:
                        ax.plot(spacings, y_gen, linewidth=2, label="Constant")
                    else:
                        ax.plot(spacings, y_gen, linewidth=2, label="General")
                    ax.set_xlabel("Clear spacing (mm)")
                    ax.set_ylabel(f"{y_axis_opt} (mm)")
                    ax.set_title(f"Sensitivity vs spacing – bar {chart_db} mm")
                    ax.grid(True, linestyle=":")
                    ax.legend()
                    st.pyplot(fig)

            with t_geo_m:
                geometry_panel_inputs = dict(st.session_state.get("calc_inputs", {}) or {})
                geometry_panel_inputs.update(st.session_state.get("geometry_live_inputs", {}) or {})
                render_live_geometry_panel(
                    geometry_panel_inputs,
                    st.session_state.get("df_res", None),
                    st.session_state.get("qa_messages", []),
                    units,
                    edition,
                )
                st.markdown("---")
                st.markdown("### Bend Geometry Table — Main Bars")
                df_main = st.session_state.get("df_bosb_main", None)
                if df_main is None or getattr(df_main, "empty", False):
                    st.info("Geometry-Main table is not available yet. Run a calculation first.")
                else:
                    st.dataframe(df_main, hide_index=True, width=1500)
            with t_geo_s:
                df_stir = st.session_state.get("df_bosb_stir", None)
                if df_stir is None or getattr(df_stir, "empty", False):
                    st.info("Geometry-Stirrup table is not available yet. Run a calculation first.")
                else:
                    st.dataframe(df_stir, hide_index=True, width=1500)
            with t_xls:
                excel_blob = st.session_state.get('excel_blob', None)
                excel_name = st.session_state.get('excel_filename', 'ACI_Rebar_Results.xlsx')
                if not excel_blob:
                    st.info('Excel file is not available yet. Run Calculate / Update first.')
                else:
                    st.download_button(label='Download Excel workbook', data=excel_blob, file_name=excel_name, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key='dl_excel_tab')
            with t_prt:
                inputs = normalize_calc_inputs(st.session_state.get("calc_inputs", {}) or {}, enforce_required=False)
                st.session_state.calc_inputs = inputs
                df_res = st.session_state.get("df_res", None)
                if inputs is None or df_res is None:
                    st.info("HTML report is not available yet. Run a calculation first.")
                else:
                    is_edu_local = bool(inputs.get("is_edu", False)) or str(inputs.get("user_mode", "")).startswith("🎓")
                    if is_edu_local:
                        st.markdown("### 🎓 Educational HTML report")
                        file_name = "ACI_Report_v61_EDU.html"
                    else:
                        st.markdown("### Professional HTML report")
                        file_name = "ACI_Report_v61.html"
                    html_rep = st.session_state.get("html_report", None)
                    html_name = st.session_state.get("html_report_filename", None) or file_name
                    if not html_rep:
                        try:
                            html_rep = generate_html_report(inputs, df_res)
                            st.session_state.html_report = html_rep
                            st.session_state.html_report_filename = html_name
                        except Exception as e:
                            st.session_state.export_errors = st.session_state.get("export_errors", {}) or {}
                            st.session_state.export_errors["html"] = f"{type(e).__name__}: {e}"
                            st.error(f"HTML report generation failed: {type(e).__name__}: {e}")
                            html_rep = None
                    if html_rep:
                        st.download_button(label="Download HTML report", data=html_rep, file_name=html_name, mime="text/html", key="dl_html_tab")
                    else:
                        st.info("HTML report is not available yet. Run Calculate / Update first.")




if __name__ == "__main__":
    main()
