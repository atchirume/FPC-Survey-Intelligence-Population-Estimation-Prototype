# ============================================================
# FPC SURVEY INTELLIGENCE & POPULATION ESTIMATION PLATFORM
# Streamlit application for Foreign Private Capital Surveys
# ============================================================
#
# Run from Terminal or Jupyter:
#
# IMPORTANT:
# - Native/typed PDFs are extracted with PyMuPDF.
# - Scanned PDFs are OCR-processed with Tesseract.
# - OCR/manual-review flags are retained because handwriting can be
#   misread by OCR and should not be treated as automatically verified.
# ============================================================

import io
import os
import re
import sys
import shutil
import subprocess
import tempfile
import pickle
from pathlib import Path
from copy import deepcopy
from datetime import datetime

# -----------------------------
# Package bootstrap
# -----------------------------
def ensure_package(import_name, pip_name=None):
    pip_name = pip_name or import_name
    try:
        __import__(import_name)
        return True
    except ImportError:
        try:
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", pip_name],
                stdout=subprocess.DEVNULL
            )
            __import__(import_name)
            return True
        except Exception:
            return False

REQUIRED = {
    "streamlit": "streamlit>=1.32,<2",
    "pandas": "pandas>=2.0",
    "numpy": "numpy>=1.24",
    "openpyxl": "openpyxl>=3.1",
    "fitz": "pymupdf>=1.24",
    "PIL": "Pillow<11",
    "plotly": "plotly>=5.18",
    "pytesseract": "pytesseract>=0.3.10",
    "docx": "python-docx>=1.1",
    "pptx": "python-pptx>=0.6.23",
    "reportlab": "reportlab>=4.0",
}

for imp, pip_name in REQUIRED.items():
    ensure_package(imp, pip_name)

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import fitz
from PIL import Image
import pytesseract
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# -----------------------------
# Page configuration
# -----------------------------
st.set_page_config(
    page_title="FPC Survey Intelligence & Population Estimation",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)


st.info(
    "Public prototype: this application is intended for research, analytical and demonstration purposes. "
    "It is not an officially authorised institutional system and should not be interpreted as representing "
    "the policies, decisions, systems or official statistics of any government or other institution."
)

# -----------------------------
# CSS
# -----------------------------
st.markdown("""
<style>
:root {
    --app-blue: #1F4E79;
    --app-blue-dark: #163A5C;
    --app-blue-mid: #2F75B5;
    --app-gold: #2A9D8F;
    --app-gold-dark: #1F7A6E;
    --app-pale: #F5F8FC;
    --app-border: #D8E1EC;
    --app-text: #16324A;
}

/* Overall application */
.stApp {
    background: linear-gradient(180deg, #F7F9FC 0%, #FFFFFF 42%);
    color: var(--app-text);
}
.main .block-container {
    padding-top: 1.15rem;
    padding-bottom: 3rem;
    max-width: 1500px;
}

/* Institutional top accent */
.block-container::before {
    content: "";
    display: block;
    height: 5px;
    border-radius: 0 0 8px 8px;
    background: linear-gradient(90deg, var(--app-blue) 0%, var(--app-blue) 72%, var(--app-gold) 72%, var(--app-gold) 100%);
    margin: -1.15rem -1rem 1.15rem -1rem;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, var(--app-blue-dark) 0%, var(--app-blue) 55%, #245A7A 100%);
}
section[data-testid="stSidebar"] > div {
    padding-top: 0.75rem;
}
section[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}
section[data-testid="stSidebar"] .stCaption,
section[data-testid="stSidebar"] small {
    color: #DCEAF6 !important;
}
section[data-testid="stSidebar"] hr {
    border-color: rgba(242,178,51,0.55);
}
section[data-testid="stSidebar"] .stRadio > label {
    font-weight: 700;
    letter-spacing: 0.02em;
}
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] {
    gap: 0.15rem;
}
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label {
    border-radius: 9px;
    padding: 0.45rem 0.5rem;
}
section[data-testid="stSidebar"] .stRadio div[role="radiogroup"] > label:hover {
    background: rgba(242,178,51,0.14);
}
section[data-testid="stSidebar"] .stMetric {
    background: rgba(255,255,255,0.08);
    border: 1px solid rgba(255,255,255,0.14);
    border-radius: 12px;
    padding: 0.35rem 0.55rem;
}
section[data-testid="stSidebar"] .stButton > button {
    background: rgba(255,255,255,0.08);
    border: 1px solid rgba(242,178,51,0.65);
    color: #FFFFFF !important;
}

/* Institutional brand lockup */
.app-lockup {
    display: flex;
    align-items: center;
    gap: 0.75rem;
    padding: 0.45rem 0 0.8rem 0;
}
.app-lockup img {
    width: 62px;
    height: 62px;
    object-fit: contain;
    flex: 0 0 auto;
}
.app-lockup .bank-name {
    font-size: 0.88rem;
    font-weight: 800;
    line-height: 1.15;
    letter-spacing: 0.035em;
}
.app-lockup .system-name {
    color: var(--app-gold) !important;
    font-size: 0.75rem;
    font-weight: 700;
    line-height: 1.2;
    margin-top: 0.2rem;
}

/* Main hero */
.app-hero {
    background: linear-gradient(135deg, #163A5C 0%, #1F4E79 62%, #0A5B8D 100%);
    border-radius: 20px;
    padding: 1.65rem 1.75rem;
    color: white;
    box-shadow: 0 12px 35px rgba(0,59,111,0.18);
    border: 1px solid rgba(242,178,51,0.32);
    margin-bottom: 1.25rem;
    position: relative;
    overflow: hidden;
}
.app-hero::after {
    content: "";
    position: absolute;
    right: -70px;
    top: -90px;
    width: 250px;
    height: 250px;
    border: 28px solid rgba(242,178,51,0.15);
    border-radius: 50%;
}
.app-hero .eyebrow {
    color: #2A9D8F;
    font-size: 0.82rem;
    font-weight: 800;
    letter-spacing: 0.14em;
    text-transform: uppercase;
}
.app-hero h1 {
    color: white;
    margin: 0.25rem 0 0.25rem 0;
    font-size: 2.25rem;
    line-height: 1.08;
}
.app-hero p {
    color: #E7F1F8;
    margin: 0;
    font-size: 1.02rem;
}

/* Cards */
.app-card {
    background: #FFFFFF;
    border: 1px solid var(--app-border);
    border-top: 4px solid var(--app-blue);
    border-radius: 14px;
    padding: 1rem 1.1rem;
    min-height: 120px;
    box-shadow: 0 5px 18px rgba(22,50,74,0.06);
}
.app-card.gold {
    border-top-color: var(--app-gold);
}
.app-card h4 {
    margin: 0 0 0.35rem 0;
    color: var(--app-blue);
}
.app-card p {
    margin: 0;
    color: #5C6B78;
    font-size: 0.9rem;
}

/* Section headings */
h1, h2, h3 {
    color: var(--app-blue);
}
h2 {
    border-left: 5px solid var(--app-gold);
    padding-left: 0.65rem;
}

/* Metrics */
div[data-testid="stMetric"] {
    background: white;
    border: 1px solid var(--app-border);
    border-radius: 12px;
    padding: 0.7rem 0.85rem;
    box-shadow: 0 4px 15px rgba(22,50,74,0.05);
}

/* Buttons */
.stButton > button,
.stDownloadButton > button {
    border-radius: 9px !important;
    border: 1px solid var(--app-blue) !important;
    font-weight: 700 !important;
}
.stButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"] {
    background: var(--app-blue) !important;
    color: white !important;
}
.stButton > button:hover,
.stDownloadButton > button:hover {
    border-color: var(--app-gold-dark) !important;
    box-shadow: 0 4px 12px rgba(0,59,111,0.15);
}

/* Inputs */
.stTextInput input:focus,
.stSelectbox div[data-baseweb="select"]:focus-within,
.stFileUploader section:focus-within {
    border-color: var(--app-gold) !important;
    box-shadow: 0 0 0 1px var(--app-gold) !important;
}

/* File uploader */
[data-testid="stFileUploader"] section {
    border: 2px dashed #9DB5C9;
    background: #F8FBFE;
    border-radius: 14px;
}
[data-testid="stFileUploader"] section:hover {
    border-color: var(--app-gold);
    background: #FFFDF7;
}

/* Dataframes */
[data-testid="stDataFrame"] {
    border: 1px solid var(--app-border);
    border-radius: 10px;
    overflow: hidden;
}

/* Status banners */
.app-security {
    background: linear-gradient(90deg, #FFF8E7, #FFFDF7);
    border: 1px solid #F0D58A;
    border-left: 5px solid var(--app-gold);
    border-radius: 12px;
    padding: 0.85rem 1rem;
    color: #5C4611;
}
.app-info {
    background: #F2F8FD;
    border: 1px solid #C8DDED;
    border-left: 5px solid var(--app-blue);
    border-radius: 12px;
    padding: 0.85rem 1rem;
}

/* About sections */
.about-box {
    background: white;
    border: 1px solid var(--app-border);
    border-radius: 14px;
    padding: 1rem 1.15rem;
    box-shadow: 0 4px 15px rgba(22,50,74,0.04);
}
.about-box strong {
    color: var(--app-blue);
}

/* Footer */
.app-footer {
    margin-top: 2rem;
    padding: 1.1rem 1.2rem;
    background: var(--app-blue-dark);
    border-radius: 14px;
    color: #DCEAF6;
    border-top: 4px solid var(--app-gold);
    text-align: center;
    font-size: 0.82rem;
}
.app-footer strong {
    color: white;
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# APPLICATION STATE, PERSISTENCE AND DATA WORKFLOW
# ============================================================

APP_VERSION = "V4.1 Public Prototype • BPM6 BOP-IIP + Survey Estimation + Accounting Validation"
# V4.1 adds a persistent accounting-validation audit layer while remaining compatible
# with the existing V4.0/V3.4 persisted survey datasets.
APP_STATE_VERSION = "FPC_V4_1_ACCOUNTING_VALIDATION_BPM6_BOP_IIP_20260922"

TABLE_KEYS = [
    "company_master", "financials", "industrial", "shareholding", "equity",
    "profits", "asset_profits", "non_equity_liabilities",
    "non_equity_assets", "services", "comments", "qa", "data_quality",
    "bop_iip_detail", "bop_iip_summary", "iip_reconciliation", "bop_coverage", "bpm6_item_catalog",
    "estimation_microdata", "population_estimates", "population_breakdowns", "estimation_qa",
    "accounting_validation",
    "raw_text"
]

DEFAULT_STATE = {
    "records": [],
    **{k: pd.DataFrame() for k in TABLE_KEYS},
    "cleaned_tables": {},
    "cleaning_log": pd.DataFrame(),
    "sampling_frame": pd.DataFrame(),
    "estimation_config": {},
    "estimation_metadata": {},
    "estimation_ready": False,
    "estimation_run": None,
    "estimation_linkage": pd.DataFrame(),
    "accounting_validation": pd.DataFrame(),
    "accounting_validation_status": "NOT_RUN",
    "accounting_validation_run": None,
    "accounting_validation_override": False,
    "processed": False,
    "cleaned": False,
    "last_run": None,
    "last_cleaned": None,
    "dataset_id": None,
}

PERSISTENCE_FILE = Path(__file__).with_name(".fpc_v4_0_fpc_survey_intelligence_dataset.pkl")
LEGACY_PERSISTENCE_FILE = Path(__file__).with_name(".fpc_v3_4_bop_iip_dataset.pkl")


def _blank_state():
    return {k: deepcopy(v) for k, v in DEFAULT_STATE.items()}


def _normalise_company_names(df):
    if not isinstance(df, pd.DataFrame):
        return pd.DataFrame(columns=["company_name"])
    out = df.copy()
    if "company_name" not in out.columns:
        # Recover from common alternative labels if an old export used one.
        alternatives = [
            "Company Name", "company", "Company", "name", "Name",
            "companyname", "Company_Name"
        ]
        found = next((c for c in alternatives if c in out.columns), None)
        if found:
            out = out.rename(columns={found: "company_name"})
    if "company_name" not in out.columns:
        out["company_name"] = pd.Series(dtype="string")
    out["company_name"] = (
        out["company_name"].astype("string").str.replace(r"\s+", " ", regex=True).str.strip()
    )
    return out


def _recover_company_master_from_tables(tables, records=None):
    """Return a structurally valid Company Master.

    Older builds sometimes persisted a one-column Company Master containing only
    company_name. That is not sufficient for the explorer or analytics. Prefer a
    complete table, otherwise rebuild from the per-company records.
    """
    cm = tables.get("company_master")
    if isinstance(cm, pd.DataFrame) and "company_name" in cm.columns:
        cm = _normalise_company_names(cm)
        meaningful_cols = [c for c in cm.columns if c != "company_name"]
        if meaningful_cols and not cm.empty:
            return cm

    if isinstance(records, list) and records:
        company_rows = []
        for rec in records:
            if isinstance(rec, dict) and isinstance(rec.get("company"), dict):
                company_rows.append(rec["company"])
        if company_rows:
            rebuilt = pd.DataFrame(company_rows)
            if "company_name" in rebuilt.columns:
                return _normalise_company_names(rebuilt)

    # Last-resort recovery from any canonical table. Keep all company names but
    # do not fabricate missing company-level fields.
    names = []
    for key in [
        "financials", "industrial", "shareholding", "equity", "profits",
        "asset_profits", "non_equity_liabilities", "non_equity_assets",
        "services", "comments", "qa", "raw_text"
    ]:
        df = tables.get(key)
        if isinstance(df, pd.DataFrame) and "company_name" in df.columns:
            vals = df["company_name"].astype("string").str.strip()
            names.extend([x for x in vals.dropna().tolist() if str(x).strip()])

    names = list(dict.fromkeys(names))
    if names:
        return pd.DataFrame({"company_name": names})
    return pd.DataFrame(columns=["company_name"])


def _validate_table(key, value):
    if not isinstance(value, pd.DataFrame):
        return False
    if key == "company_master" and "company_name" not in value.columns:
        return False
    return True


def _canonicalise_tables(tables, records=None):
    out = {}
    for key in TABLE_KEYS:
        value = tables.get(key, pd.DataFrame())
        out[key] = value.copy() if isinstance(value, pd.DataFrame) else pd.DataFrame()

    out["company_master"] = _recover_company_master_from_tables(out, records=records)
    return out

def _apply_tables(tables, records=None):
    tables = _canonicalise_tables(tables, records=records)
    for key in TABLE_KEYS:
        st.session_state[key] = tables[key]
    st.session_state["processed"] = dataset_company_count(tables) > 0
    return st.session_state["processed"]


def _build_persistence_payload():
    tables = {k: st.session_state.get(k, pd.DataFrame()) for k in TABLE_KEYS}
    tables = _canonicalise_tables(tables, records=st.session_state.get("records", []))
    return {
        "state_version": APP_STATE_VERSION,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "dataset_id": st.session_state.get("dataset_id"),
        "last_run": st.session_state.get("last_run"),
        "last_cleaned": st.session_state.get("last_cleaned"),
        "processed": bool(st.session_state.get("processed")),
        "cleaned": bool(st.session_state.get("cleaned")),
        "tables": tables,
        "cleaned_tables": {
            k: v for k, v in st.session_state.get("cleaned_tables", {}).items()
            if isinstance(v, pd.DataFrame)
        },
        "cleaning_log": st.session_state.get("cleaning_log", pd.DataFrame()),
        "sampling_frame": st.session_state.get("sampling_frame", pd.DataFrame()),
        "estimation_config": st.session_state.get("estimation_config", {}),
        "estimation_metadata": st.session_state.get("estimation_metadata", {}),
        "estimation_ready": bool(st.session_state.get("estimation_ready", False)),
        "estimation_run": st.session_state.get("estimation_run"),
        "population_estimates": st.session_state.get("population_estimates", pd.DataFrame()),
        "population_breakdowns": st.session_state.get("population_breakdowns", pd.DataFrame()),
        "estimation_qa": st.session_state.get("estimation_qa", pd.DataFrame()),
        "estimation_microdata": st.session_state.get("estimation_microdata", pd.DataFrame()),
        "estimation_linkage": st.session_state.get("estimation_linkage", pd.DataFrame()),
        "accounting_validation": st.session_state.get("accounting_validation", pd.DataFrame()),
        "accounting_validation_status": st.session_state.get("accounting_validation_status", "NOT_RUN"),
        "accounting_validation_run": st.session_state.get("accounting_validation_run"),
        "accounting_validation_override": bool(st.session_state.get("accounting_validation_override", False)),
        # Keep records as a recovery source, but the canonical tables above
        # are what the application reads.
        "records": st.session_state.get("records", []),
    }


def save_persisted_state():
    try:
        payload = _build_persistence_payload()
        tmp = PERSISTENCE_FILE.with_suffix(".tmp")
        with open(tmp, "wb") as fh:
            pickle.dump(payload, fh, protocol=pickle.HIGHEST_PROTOCOL)
        tmp.replace(PERSISTENCE_FILE)
        return True, ""
    except Exception as exc:
        return False, str(exc)


def _quarantine_persistence(reason="legacy"):
    """Move an incompatible/corrupt persistence file aside instead of reusing it."""
    try:
        if PERSISTENCE_FILE.exists():
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            target = PERSISTENCE_FILE.with_name(
                f"{PERSISTENCE_FILE.stem}.{reason}.{stamp}{PERSISTENCE_FILE.suffix}"
            )
            PERSISTENCE_FILE.replace(target)
            return str(target)
    except Exception:
        pass
    return ""


def load_persisted_state():
    """Load a complete canonical dataset only after all parser functions exist."""
    persistence_source = PERSISTENCE_FILE if PERSISTENCE_FILE.exists() else LEGACY_PERSISTENCE_FILE
    if not persistence_source.exists():
        return False, ""

    try:
        with open(persistence_source, "rb") as fh:
            payload = pickle.load(fh)

        if not isinstance(payload, dict):
            return False, "The saved dataset is not a valid application dataset."

        # V4 can safely migrate the V3.4/V3.5 canonical survey tables.
        # Population-estimation fields are initialised empty when they were not
        # present in the legacy persistence payload.
        saved_version = payload.get("state_version")
        compatible_versions = {
            APP_STATE_VERSION,
            "FPC_V4_0_BPM6_BOP_IIP_SURVEY_ESTIMATION_20260921",
            "FPC_V3_4_BPM6_BOP_IIP_20260921",
        }
        if saved_version and saved_version not in compatible_versions:
            quarantined = _quarantine_persistence("legacy")
            detail = f" The old file was isolated as {quarantined}." if quarantined else ""
            return False, (
                f"A previous dataset from {saved_version} was found and was not restored."
                + detail + " Re-process the PDFs to create a clean canonical dataset."
            )

        # V3 native format.
        if isinstance(payload.get("tables"), dict):
            records = payload.get("records", [])
            tables = _canonicalise_tables(payload["tables"], records=records)

            if company_count(tables["company_master"]) == 0:
                return False, "The saved dataset contains no recoverable company names."

            _apply_tables(tables)
            cleaned = payload.get("cleaned_tables", {})
            st.session_state["cleaned_tables"] = {
                k: v.copy() for k, v in cleaned.items()
                if isinstance(v, pd.DataFrame)
            }
            st.session_state["cleaning_log"] = (
                payload.get("cleaning_log", pd.DataFrame())
                if isinstance(payload.get("cleaning_log", pd.DataFrame()), pd.DataFrame)
                else pd.DataFrame()
            )
            st.session_state["sampling_frame"] = (
                payload.get("sampling_frame", pd.DataFrame())
                if isinstance(payload.get("sampling_frame", pd.DataFrame()), pd.DataFrame)
                else pd.DataFrame()
            )
            st.session_state["estimation_config"] = payload.get("estimation_config", {}) or {}
            st.session_state["estimation_metadata"] = payload.get("estimation_metadata", {}) or {}
            st.session_state["estimation_ready"] = bool(payload.get("estimation_ready", False))
            st.session_state["estimation_run"] = payload.get("estimation_run")
            for _ek in ["population_estimates", "population_breakdowns", "estimation_qa", "estimation_microdata", "estimation_linkage"]:
                _ev = payload.get(_ek, pd.DataFrame())
                st.session_state[_ek] = _ev.copy() if isinstance(_ev, pd.DataFrame) else pd.DataFrame()
            _av = payload.get("accounting_validation", pd.DataFrame())
            st.session_state["accounting_validation"] = _av.copy() if isinstance(_av, pd.DataFrame) else pd.DataFrame()
            st.session_state["accounting_validation_status"] = payload.get("accounting_validation_status", "NOT_RUN")
            st.session_state["accounting_validation_run"] = payload.get("accounting_validation_run")
            st.session_state["accounting_validation_override"] = bool(payload.get("accounting_validation_override", False))
            st.session_state["records"] = payload.get("records", [])
            st.session_state["last_run"] = payload.get("last_run")
            st.session_state["last_cleaned"] = payload.get("last_cleaned")
            st.session_state["dataset_id"] = payload.get("dataset_id")
            st.session_state["cleaned"] = bool(
                st.session_state["cleaned_tables"]
            )
            st.session_state["processed"] = True
            if persistence_source != PERSISTENCE_FILE:
                save_persisted_state()
            return True, ""

        # Legacy V2 recovery.
        legacy_tables = {
            k: payload.get(k, pd.DataFrame()) for k in TABLE_KEYS
        }
        legacy_tables = _canonicalise_tables(legacy_tables, records=payload.get("records", []))
        if company_count(legacy_tables["company_master"]) > 0:
            _apply_tables(legacy_tables)
            _av = payload.get("accounting_validation", pd.DataFrame())
            st.session_state["accounting_validation"] = _av.copy() if isinstance(_av, pd.DataFrame) else pd.DataFrame()
            st.session_state["accounting_validation_status"] = payload.get("accounting_validation_status", "NOT_RUN")
            st.session_state["accounting_validation_run"] = payload.get("accounting_validation_run")
            st.session_state["accounting_validation_override"] = bool(payload.get("accounting_validation_override", False))
            st.session_state["records"] = payload.get("records", [])
            st.session_state["last_run"] = payload.get("last_run")
            st.session_state["dataset_id"] = payload.get(
                "dataset_id",
                datetime.now().strftime("%Y%m%d%H%M%S")
            )
            st.session_state["cleaned_tables"] = {}
            st.session_state["cleaned"] = False
            st.session_state["accounting_validation"] = pd.DataFrame()
            st.session_state["accounting_validation_status"] = "NOT_RUN"
            st.session_state["accounting_validation_run"] = None
            st.session_state["accounting_validation_override"] = False
            save_persisted_state()
            return True, ""

        return False, "The previous dataset could not be reconstructed."
    except Exception as exc:
        quarantined = _quarantine_persistence("corrupt")
        detail = f" Persistence was isolated as {quarantined}." if quarantined else ""
        return False, str(exc) + detail


def company_count(df):
    if not isinstance(df, pd.DataFrame) or df.empty:
        return 0
    if "company_name" not in df.columns:
        return 0
    names = (
        df["company_name"].astype("string")
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
        .dropna()
    )
    return int(names.nunique())


def dataset_company_count(tables=None):
    """Count companies from any valid canonical table, not only Company Master."""
    if tables is None:
        tables = {k: st.session_state.get(k, pd.DataFrame()) for k in TABLE_KEYS}
    names = set()
    for df in tables.values():
        if not isinstance(df, pd.DataFrame) or df.empty or "company_name" not in df.columns:
            continue
        vals = (
            df["company_name"].astype("string")
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
            .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
            .dropna()
        )
        names.update(str(x) for x in vals.tolist() if str(x).strip())
    return len(names)


def initialise_state():
    if st.session_state.get("_fpc_state_version") != APP_STATE_VERSION:
        for k, v in DEFAULT_STATE.items():
            st.session_state[k] = deepcopy(v)
        st.session_state["_fpc_state_version"] = APP_STATE_VERSION

        restored, err = load_persisted_state()
        st.session_state["_fpc_restore_error"] = err
        st.session_state["_fpc_restored"] = restored
    else:
        for k, v in DEFAULT_STATE.items():
            if k not in st.session_state:
                st.session_state[k] = deepcopy(v)

        # If session state is empty, attempt a restore.
        if not st.session_state.get("processed", False):
            restored, err = load_persisted_state()
            st.session_state["_fpc_restore_error"] = err
            st.session_state["_fpc_restored"] = restored


# ============================================================
# PDF EXTRACTION / PARSING ENGINE
# ============================================================

# -----------------------------
# Utilities
# -----------------------------
def clean_text(x):
    if x is None:
        return ""
    x = str(x).replace("\x00", " ")
    x = x.replace("￾", "-").replace("–", "-").replace("—", "-")
    x = re.sub(r"[ \t]+", " ", x)
    x = re.sub(r"\n{3,}", "\n\n", x)
    return x.strip()

def normalise_line(line):
    line = clean_text(line)
    line = re.sub(r"\s+", " ", line)
    return line.strip()

def safe_filename(name):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name)

def money_to_float(value):
    if value is None:
        return np.nan
    s = str(value).strip()
    if not s:
        return np.nan
    if s.upper() in {"N/A", "NA", "N.A", "N.A.", "-", "—", "NOT APPLICABLE", "NIL"}:
        return np.nan

    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True

    # Remove currency labels and common OCR artifacts
    s = re.sub(r"(?i)US\$|USD|\$", "", s)
    s = s.replace(",", "").replace(" ", "")
    s = s.strip("()")

    # Keep only plausible numeric characters
    s = re.sub(r"[^0-9.\-]", "", s)

    if not s or s in {"-", ".", "-."}:
        return np.nan

    try:
        v = float(s)
        return -abs(v) if negative else v
    except Exception:
        return np.nan

def extract_first(patterns, text, flags=re.I | re.S):
    for p in patterns:
        m = re.search(p, text, flags)
        if m:
            val = m.group(1).strip()
            val = re.sub(r"\s+", " ", val)
            return val
    return ""

def find_value_after_label(lines, labels, max_ahead=2):
    labels_lower = [x.lower() for x in labels]
    for i, line in enumerate(lines):
        low = line.lower()
        for label in labels_lower:
            if label in low:
                # same-line after colon
                if ":" in line:
                    val = line.split(":", 1)[1].strip()
                    if val and val.lower() != label:
                        return val
                # text after label
                pos = low.find(label)
                after = line[pos + len(label):].strip(" :-")
                if after:
                    return after
                # next lines
                for j in range(1, max_ahead + 1):
                    if i + j < len(lines):
                        nxt = lines[i + j].strip()
                        if nxt and not nxt.lower().startswith(("date", "website", "district")):
                            return nxt
    return ""

def extract_all_numbers(text):
    # Captures values such as 1,657,618; (2,819,915); 203,341.56
    pat = r"(?<![A-Za-z0-9])(?:\(?-?\s*\$?\s*\d{1,3}(?:,\d{3})+(?:\.\d+)?\)?|\(?-?\s*\d+(?:\.\d+)?\)?)(?![A-Za-z0-9])"
    return re.findall(pat, text)

def numeric_candidates(line):
    vals = []
    for x in extract_all_numbers(line):
        v = money_to_float(x)
        if not pd.isna(v):
            vals.append(v)
    return vals

TESSERACT_PATH = None

def detect_tesseract():
    """Detect a working Tesseract executable without ever raising at import time."""
    candidates = [
        os.environ.get("TESSERACT_CMD"),
        shutil.which("tesseract"),
        "/opt/homebrew/bin/tesseract",
        "/usr/local/bin/tesseract",
        "/usr/bin/tesseract",
    ]
    seen = set()
    for candidate in candidates:
        if not candidate:
            continue
        candidate = os.path.expanduser(str(candidate))
        if candidate in seen or not os.path.isfile(candidate) or not os.access(candidate, os.X_OK):
            continue
        seen.add(candidate)
        try:
            subprocess.run([candidate, "--version"], check=True,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=5)
            pytesseract.pytesseract.tesseract_cmd = candidate
            return candidate
        except Exception:
            continue
    return None

try:
    TESSERACT_PATH = detect_tesseract()
except Exception:
    TESSERACT_PATH = None

# -----------------------------
# PDF extraction
# -----------------------------
def detect_table_xbounds(drawings, page_width):
    """Infer questionnaire table column boundaries from PDF cell rectangles."""
    xs=[]
    for d in drawings or []:
        r=d.get("rect")
        if r is None:
            continue
        # Table cells have repeated rectangle borders. Ignore small text boxes.
        if r.y0 < page_width and r.y1 > r.y0:
            xs.extend([float(r.x0), float(r.x1)])
    if not xs:
        return None
    xs=sorted(xs)
    clusters=[]
    for x in xs:
        if not clusters or abs(x-clusters[-1][-1])>4.0:
            clusters.append([x])
        else:
            clusters[-1].append(x)
    centers=[sum(c)/len(c) for c in clusters]
    # Select a long sequence of boundaries. The FPC non-equity table has 12
    # columns: type, source, relationship, maturity, A, B, C, D1, D2, D3, E, G.
    candidates=[x for x in centers if 40 <= x <= page_width-20]
    if len(candidates) >= 13:
        # The first and last major boundaries normally bracket the whole table.
        # Keep the densest ordered sequence by selecting boundaries separated by
        # at least 20 points.
        selected=[]
        for x in candidates:
            if not selected or x-selected[-1] >= 20:
                selected.append(x)
        if len(selected) >= 13:
            return selected
    return None



def extract_pdf(uploaded_file, dpi=220, ocr_if_needed=True):
    """
    Extract a questionnaire while retaining three parallel representations:

    1. native page text;
    2. word coordinates;
    3. PyMuPDF table geometry/extractions.

    The third representation is important because FPC questionnaires are
    highly structured forms. Sequential text extraction can lose blank cells
    and consequently shift financial values into the wrong equity/service row.
    """
    pdf_bytes = uploaded_file.getvalue()
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    page_texts = []
    page_words = []
    page_table_xbounds = []
    page_sizes = []
    page_tables = []
    ocr_pages = []
    ocr_confidences = []
    page_meta = []

    for page_no, page in enumerate(doc, start=1):
        txt = clean_text(page.get_text("text"))
        words = page.get_text("words")
        page_words.append(words)
        page_table_xbounds.append(
            detect_table_xbounds(page.get_drawings(), page.rect.width)
        )
        page_sizes.append((float(page.rect.width), float(page.rect.height)))

        # Keep table extraction independent of OCR. Typed questionnaires have
        # real PDF vector/text tables and these are materially safer than
        # reconstructing rows from a flattened text stream.
        extracted_tables = []
        try:
            finder = page.find_tables()
            for table in getattr(finder, "tables", []) or []:
                try:
                    extracted_tables.append(table.extract())
                except Exception:
                    continue
        except Exception:
            extracted_tables = []
        page_tables.append(extracted_tables)

        native_chars = len(re.sub(r"\s+", "", txt))
        used_ocr = False
        ocr_confidence = np.nan

        if ocr_if_needed and native_chars < 80:
            if TESSERACT_PATH:
                try:
                    pix = page.get_pixmap(dpi=dpi, alpha=False)
                    img = Image.open(io.BytesIO(pix.tobytes("png")))
                    txt = clean_text(
                        pytesseract.image_to_string(img, config="--psm 6")
                    )
                    try:
                        ocr_data = pytesseract.image_to_data(
                            img,
                            config="--psm 6",
                            output_type=pytesseract.Output.DICT,
                        )
                        confs = []
                        for c in ocr_data.get("conf", []):
                            try:
                                v = float(c)
                                if v >= 0:
                                    confs.append(v)
                            except Exception:
                                pass
                        if confs:
                            ocr_confidence = float(np.mean(confs))
                    except Exception:
                        pass
                    used_ocr = True
                    ocr_pages.append(page_no)
                    ocr_confidences.append({
                        "page": page_no,
                        "confidence": ocr_confidence
                    })
                except Exception:
                    txt = ""
            else:
                txt = ""

        page_texts.append(txt)
        page_meta.append({
            "page": page_no,
            "native_characters": native_chars,
            "ocr_used": used_ocr,
            "ocr_confidence": ocr_confidence,
            "text_characters": len(txt),
            "table_count": len(extracted_tables),
        })

    return {
        "filename": uploaded_file.name,
        "text": "\n".join(page_texts),
        "page_texts": page_texts,
        "page_words": page_words,
        "page_table_xbounds": page_table_xbounds,
        "page_sizes": page_sizes,
        "page_tables": page_tables,
        "ocr_pages": ocr_pages,
        "ocr_confidences": ocr_confidences,
        "page_meta": page_meta,
        "page_count": len(doc),
        "tesseract_available": bool(TESSERACT_PATH),
    }

# -----------------------------

# -----------------------------
# Questionnaire parser V2
# -----------------------------
MISSING_TOKENS = {
    "", "N/A", "NA", "N.A", "N.A.", "NONE", "NIL", "NILL", "N?A",
    "N//A", "N./A", "N./A.", "N?A", "NOT APPLICABLE"
}
ZERO_LIKE_TOKENS = {"-", "--"}

def clean_text(x):
    if x is None:
        return ""
    x = str(x).replace("\x00", " ")
    x = x.replace("￾", "-").replace("–", "-").replace("—", "-")
    x = re.sub(r"[ \t]+", " ", x)
    x = re.sub(r"\n{3,}", "\n\n", x)
    return x.strip()

def normalise_line(line):
    line = clean_text(line)
    return re.sub(r"\s+", " ", line).strip()

def safe_filename(name):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name)

def cell_state(raw):
    """Return raw cell, numeric value where explicit, and semantic status."""
    if raw is None:
        return {"raw": "", "value": np.nan, "status": "BLANK"}
    s = str(raw).strip()
    if not s:
        return {"raw": "", "value": np.nan, "status": "BLANK"}
    su = re.sub(r"\s+", " ", s).upper()
    if su in MISSING_TOKENS - {""}:
        return {"raw": s, "value": np.nan, "status": "N/A"}
    if su in ZERO_LIKE_TOKENS:
        # Keep dash as a dash. Do not silently convert it to zero.
        return {"raw": s, "value": np.nan, "status": "DASH"}
    v = money_to_float(s)
    if not pd.isna(v):
        return {"raw": s, "value": v, "status": "VALUE"}
    return {"raw": s, "value": np.nan, "status": "TEXT"}


def clean_table_cell(value):
    """Normalise a PyMuPDF table cell without destroying semantic text."""
    if value is None:
        return ""
    x = str(value).replace("\x00", " ").replace("\n", " ")
    x = re.sub(r"\s+", " ", x).strip()
    # PDF line wrapping often splits a number across a line boundary:
    # 45,374,814 .28 -> 45,374,814.28
    x = re.sub(r"(?<=\d)\s+(?=[.,]\d)", "", x)
    x = re.sub(r"(?<=,)\s+(?=\d)", "", x)
    return x.strip()


def table_row_clean(row):
    return [clean_table_cell(x) for x in row]


def table_text(table):
    if not table:
        return ""
    return " ".join(
        clean_table_cell(x)
        for row in table
        for x in row
        if clean_table_cell(x)
    )


def is_section_table(table, required_terms):
    txt = table_text(table).lower()
    return all(term.lower() in txt for term in required_terms)


def find_table_extract(page_tables, required_terms, preferred_min_cols=1):
    """Find the most plausible extracted table for a section."""
    candidates = []
    for page_no, tables in enumerate(page_tables or [], start=1):
        for table in tables or []:
            if not table:
                continue
            ncols = max((len(r) for r in table), default=0)
            if ncols < preferred_min_cols:
                continue
            txt = table_text(table).lower()
            score = sum(term.lower() in txt for term in required_terms)
            if score:
                candidates.append((score, ncols, page_no, table))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return candidates[0][2], candidates[0][3]


def find_all_tables(page_tables, required_terms):
    out = []
    for page_no, tables in enumerate(page_tables or [], start=1):
        for table in tables or []:
            txt = table_text(table).lower()
            if all(term.lower() in txt for term in required_terms):
                out.append((page_no, table))
    return out


def semantic_status_for_table_cell(raw):
    """Same semantics as cell_state, but after PDF-table whitespace repair."""
    return cell_state(clean_table_cell(raw))


def numeric_with_zero_fallback(state):
    """
    Return numeric value plus whether zero fallback was used.

    In the survey, N/A, '-' and blank are not equivalent source states. For
    reconciliation only, N/A and '-' may be treated as zero when the source
    clearly provides the other required balance cells. The original raw/status
    values are always retained.
    """
    if state is None:
        return np.nan, False
    if state["status"] == "VALUE":
        return state["value"], False
    if state["status"] in {"N/A", "DASH"}:
        return 0.0, True
    return np.nan, False


def canonical_company_name_from_filename(filename):
    """Derive a clean company name only when the PDF text layer is unusable."""
    stem = Path(str(filename)).stem
    x = stem
    patterns = [
        r"\s*-\s*Survey on Foreign Private Capital Flows\s+Stocks\s*-\s*Questionnaire.*$",
        r"\s*-\s*Survey on Foreign Private Capital Flows and Stocks.*$",
        r"\s+FPC Questionnaire.*$",
        r"\s*-\s*FPC Questionnaire.*$",
        r"\s+Questionnaire.*$",
    ]
    for p in patterns:
        x = re.sub(p, "", x, flags=re.I)
    x = re.sub(r"\(\d+\)$", "", x).strip(" -_")
    x = re.sub(r"\s+", " ", x).strip()
    return x or stem

def money_to_float(value):
    if value is None:
        return np.nan
    s = str(value).strip()
    if not s:
        return np.nan
    su = re.sub(r"\s+", " ", s).upper()
    if su in MISSING_TOKENS or su in ZERO_LIKE_TOKENS:
        return np.nan
    negative = s.startswith("(") and s.endswith(")")
    s = re.sub(r"(?i)US\$|USS\$|USD|\$", "", s)
    s = s.replace(",", "").replace(" ", "")
    s = s.strip("()")
    s = re.sub(r"[^0-9.\-]", "", s)
    if not s or s in {"-", ".", "-."}:
        return np.nan
    try:
        v = float(s)
        return -abs(v) if negative else v
    except Exception:
        return np.nan

VALUE_RE = re.compile(
    r"(?:\((?:US\$|USS\$|USD|US|\$)?-?\s*\d[\d,]*(?:\.\d+)?\)|(?:US\$|USS\$|USD|US|\$)?\(?-?\s*\d[\d,]*(?:\.\d+)?\)?)%?"
)

def value_tokens_from_lines(lines, include_missing=True):
    """
    Extract table-cell-like tokens from a block while ignoring headings.
    Only accepts lines that are themselves a cell value, not prose containing numbers.
    """
    out = []
    for line in lines:
        s = normalise_line(line)
        if not s:
            continue
        # Remove footnote markers.
        s = re.sub(r"^\*+", "", s).strip()
        if include_missing and re.fullmatch(
            r"(?i)(N/?A|N\.?A\.?|NONE|NIL|NILL|N\?A|N//A|-|--)", s
        ):
            out.append(s)
            continue
        if re.fullmatch(r"(?:\((?:US\$|USS\$|USD|US|\$)?-?\s*\d[\d,]*(?:\.\d+)?\)|(?:US\$|USS\$|USD|US|\$)?\(?-?\s*\d[\d,]*(?:\.\d+)?\)?)", s):
            out.append(s)
            continue
        # A line can contain only a value plus a simple footnote.
        m = re.fullmatch(
            r"((?:US\$|USS\$|USD|\$)?\(?-?\s*\d[\d,]*(?:\.\d+)?\)?)[*a-zA-Z0-9]*",
            s
        )
        if m:
            out.append(m.group(1))
    return out

def section_between(text, start_pattern, end_patterns):
    m = re.search(start_pattern, text, re.I | re.S)
    if not m:
        return ""
    start = m.end()
    end = len(text)
    tail = text[start:]
    positions = []
    for p in end_patterns:
        mm = re.search(p, tail, re.I | re.S)
        if mm:
            positions.append(mm.start())
    if positions:
        end = start + min(positions)
    return text[start:end]

def section_page(page_texts, pattern):
    for i, t in enumerate(page_texts, start=1):
        if re.search(pattern, t or "", re.I | re.S):
            return i
    return np.nan

def extract_labeled_line(text, label, stop_labels=None, allow_next_line=True):
    stop_labels = stop_labels or []
    pattern = r"(?<![A-Za-z0-9])" + re.escape(label) + r"(?![A-Za-z0-9])[ \t]*:?[ \t]*(.*)"
    m = re.search(pattern, text, re.I)
    if not m:
        return ""
    value = m.group(1).strip()

    def trim(v):
        if not v:
            return ""
        # Stop at another labelled field appearing on the same line.
        for stop in stop_labels:
            sm = re.search(r"(?:\s+|^)" + re.escape(stop) + r"\s*:", v, re.I)
            if sm:
                v = v[:sm.start()].strip()
        return v

    value = trim(value)
    starts_with_stop = any(
        re.match(r"^" + re.escape(stop) + r"\s*:", value, re.I)
        for stop in stop_labels
    )
    if value and not starts_with_stop and not re.fullmatch(r"\(?\s*dd\s*/\s*mm\s*/\s*yyyy\s*\)?", value, re.I):
        return value

    if allow_next_line:
        lines = text[m.end():].splitlines()
        for nxt in lines[:4]:
            nxt = normalise_line(nxt)
            if not nxt:
                continue
            # Do not consume the next labelled field.
            known_labels = {
                "district", "area", "street", "website", "date of establishment",
                "date of commencing operations", "previous name of the company (if any)",
                "reporting period", "position", "mobile", "mob", "e-mail", "email",
                "location of establishment (region)", "location of establishment"
            }
            if re.match(r"^[A-Za-z][A-Za-z0-9 /&().'-]*\s*:", nxt) or nxt.lower().rstrip(":") in known_labels:
                break
            if nxt.lower().startswith(("questionnaire", "part ", "1:", "2:", "3:")):
                break
            return trim(nxt)
    return ""

def parse_company_info(text):
    """Parse the company-information block with same-line label awareness."""
    company = extract_labeled_line(text, "Company name", ["Company Address"])
    info_block = section_between(
        text, r"Company Address\s*:",
        [r"Particulars of the person completing this questionnaire"]
    )
    if not info_block:
        info_block = text

    address = extract_labeled_line(text, "Company Address", ["Tel", "Website", "District"])
    telephone = extract_labeled_line(info_block, "Tel", ["Website", "District"])
    website = extract_labeled_line(info_block, "Website", ["District"])

    # Many questionnaires put Address, Tel and Website on one physical line.
    if "Tel:" in address or re.search(r"\bTel\s*:", address, re.I):
        if not telephone:
            m = re.search(r"\bTel\s*:\s*(.*?)(?=\s+Website\s*:|\s+District\s*:|$)", address, re.I)
            if m:
                telephone = m.group(1).strip()
        address = re.split(r"\s+Tel\s*:", address, flags=re.I)[0].strip()
    if "Website:" in address or re.search(r"\bWebsite\s*:", address, re.I):
        if not website:
            m = re.search(r"\bWebsite\s*:\s*(.*?)(?=\s+District\s*:|$)", address, re.I)
            if m:
                website = m.group(1).strip()
        address = re.split(r"\s+Website\s*:", address, flags=re.I)[0].strip()
    if not website:
        website = re.sub(r"^Website\s*:\s*", "", website, flags=re.I).strip()

    district = extract_labeled_line(info_block, "District", ["Area", "Street", "Location of Establishment"])
    area = extract_labeled_line(info_block, "Area", ["Street", "Location of Establishment"])
    street = extract_labeled_line(info_block, "Street", ["Location of Establishment"])

    region = extract_labeled_line(
        info_block, "Location of Establishment (Region)",
        ["Date of Establishment", "Date of Commencing Operations"]
    )
    if not region:
        region = extract_labeled_line(
            info_block, "Location of Establishment",
            ["Date of Establishment", "Date of Commencing Operations"]
        )
    establishment = extract_labeled_line(
        text, "Date of Establishment", ["Date of Commencing Operations"]
    )
    commencement = extract_labeled_line(
        info_block, "Date of Commencing Operations", ["Previous name of the Company (if any)", "Reporting Period"]
    )
    previous = extract_labeled_line(
        info_block, "Previous name of the Company (if any)", ["Date completed", "Reporting Period"]
    )
    completed = extract_labeled_line(info_block, "Date completed", ["Reporting Period"])
    reporting = extract_labeled_line(info_block, "Reporting Period")

    respondent_block = section_between(
        text, r"Particulars of the person completing this questionnaire",
        [r"Particulars of an alternative person to be contacted"]
    )
    respondent = extract_labeled_line(respondent_block, "Name", ["Position"])
    position = extract_labeled_line(respondent_block, "Position", ["Mobile", "Mob", "E-mail", "Email"])
    mobile = extract_labeled_line(respondent_block, "Mobile", ["Mob", "E-mail", "Email"])
    if not mobile:
        mobile = extract_labeled_line(respondent_block, "Mob", ["E-mail", "Email"])
    email = extract_labeled_line(respondent_block, "E-mail", ["Email"])
    if not email:
        email = extract_labeled_line(respondent_block, "Email")

    alt_block = section_between(
        text, r"Particulars of an alternative person to be contacted",
        [r"COMPANY AFFILIATES"]
    )
    alt_name = extract_labeled_line(alt_block, "Name", ["Position"])
    alt_position = extract_labeled_line(alt_block, "Position", ["Mob", "Mobile", "E-mail", "Email"])
    alt_mobile = extract_labeled_line(alt_block, "Mob", ["E-mail", "Email"])
    if not alt_mobile:
        alt_mobile = extract_labeled_line(alt_block, "Mobile", ["E-mail", "Email"])
    alt_email = extract_labeled_line(alt_block, "E-mail", ["Email"])
    if not alt_email:
        alt_email = extract_labeled_line(alt_block, "Email")

    affiliates = section_between(
        text, r"COMPANY AFFILIATES",
        [r"QUESTIONNAIRE FOR THE SURVEY", r"QUESTIONNAIRE FOR THE"]
    )
    subsidiaries = extract_labeled_line(affiliates, "2.1 Does your company have any subsidiaries")
    consolidated = extract_labeled_line(affiliates, "2.2 If yes, are you supplying consolidated information")

    # The template footnote marker "1" can appear inside the question text.
    # Use a targeted Yes/No extraction as a fallback so the answer itself is
    # not polluted by the full question wording.
    if not subsidiaries:
        m = re.search(r"2\.1.*?\b(Yes|No)\b", affiliates, re.I | re.S)
        subsidiaries = m.group(1) if m else ""
    if not consolidated:
        m = re.search(r"2\.2.*?\b(Yes|No)\b", affiliates, re.I | re.S)
        consolidated = m.group(1) if m else ""

    # Keep only the answer when the extracted value still contains part of the
    # question. The last Yes/No is the marked response in this form.
    for _key, _value in [("subsidiaries", subsidiaries), ("consolidated", consolidated)]:
        answers = re.findall(r"\b(Yes|No)\b", str(_value), re.I)
        if answers:
            if _key == "subsidiaries":
                subsidiaries = answers[-1]
            else:
                consolidated = answers[-1]

    # Remove accidental neighbouring labels from values.
    split_labels = r"\s+(?:Date of Establishment|Date of Commencing Operations|Previous name of the Company \(if any\)|Reporting Period)\s*:"
    region = re.split(split_labels, region, flags=re.I)[0].strip()
    establishment = re.split(split_labels, establishment, flags=re.I)[0].strip()
    commencement = re.split(split_labels, commencement, flags=re.I)[0].strip()
    completed = re.split(split_labels, completed, flags=re.I)[0].strip()

    return {
        "company_name": company,
        "company_address": address,
        "telephone": telephone,
        "website": website,
        "district": district,
        "area": area,
        "street": street,
        "establishment_region": region,
        "date_established": establishment,
        "date_commenced_operations": commencement,
        "previous_company_name": previous,
        "date_completed": completed,
        "reporting_period": reporting,
        "respondent_name": respondent,
        "respondent_position": position,
        "respondent_mobile": mobile,
        "respondent_email": email,
        "alternative_contact_name": alt_name,
        "alternative_contact_position": alt_position,
        "alternative_contact_mobile": alt_mobile,
        "alternative_contact_email": alt_email,
        "subsidiaries_in_zimbabwe": subsidiaries,
        "consolidated_group_information": consolidated,
    }

def parse_industrial(text, company, page_texts):
    sec = section_between(
        text,
        r"1:\s*INDUSTRIAL CLASSIFICATION",
        [r"PART 1:\s*FOREIGN LIABILITIES"]
    )
    if not sec:
        return pd.DataFrame()

    # Split rows 1-5. This is based on the questionnaire's actual table layout.
    sec2 = re.sub(r"\s+(?=[1-5]\.\s)", "\n", sec)
    blocks = re.split(r"(?m)^(?=[1-5]\.\s)", sec2)
    rows = []
    for block in blocks:
        b = block.strip()
        if not re.match(r"^[1-5]\.", b):
            continue
        lines = [normalise_line(x) for x in b.splitlines() if normalise_line(x)]
        row_no = re.match(r"^([1-5])\.", lines[0]).group(1)
        first = re.sub(r"^[1-5]\.\s*", "", lines[0]).strip()
        if first:
            data_lines = [first] + lines[1:]
        else:
            data_lines = lines[1:]
        if not data_lines:
            continue
        pct_idx = next(
            (i for i, x in enumerate(data_lines) if re.fullmatch(r"\d+(?:\.\d+)?%", x)),
            None
        )
        if pct_idx is None:
            continue
        sector = data_lines[0]
        description = " ".join(data_lines[1:pct_idx]).strip()
        pct = data_lines[pct_idx]
        rows.append({
            "company_name": company,
            "industry_row": int(row_no),
            "sector_industrial_classification": sector,
            "economic_activity_description": description,
            "investment_contribution_pct": money_to_float(pct.replace("%","")),
            "investment_contribution_raw": pct,
            "source_page": section_page(page_texts, r"1:\s*INDUSTRIAL CLASSIFICATION")
        })
    return pd.DataFrame(rows)

def parse_shareholding(text, company, page_texts):
    sec = section_between(
        text,
        r"1\.1:\s*SHAREHOLDING STRUCTURE",
        [r"1\.2:\s*EQUITY INVESTMENT"]
    )
    if not sec:
        return pd.DataFrame()

    sec2 = re.sub(r"\s+(?=[1-5]\.?\s)", "\n", sec)
    blocks = re.split(r"(?m)^(?=[1-5]\.?\s)", sec2)

    country_names = {
        "ZIMBABWE", "UK", "UNITED KINGDOM", "SOUTH AFRICA", "ZAMBIA",
        "MOZAMBIQUE", "BOTSWANA", "MALAWI", "MAURITIUS", "CHINA",
        "UAE", "UNITED ARAB EMIRATES", "USA", "UNITED STATES",
        "REST OF WORLD", "NAMIBIA", "TANZANIA", "KENYA", "INDIA"
    }

    rows = []
    for block in blocks:
        b = block.strip()
        if not re.match(r"^[1-5]\.?\s*", b):
            continue

        lines = [normalise_line(x) for x in b.splitlines() if normalise_line(x)]
        mrow = re.match(r"^([1-5])\.?\s*(.*)$", lines[0])
        if not mrow:
            continue

        row_no = int(mrow.group(1))
        first = mrow.group(2).strip()
        vals = ([first] if first else []) + lines[1:]
        if not vals:
            continue

        # Numeric cells may or may not contain a % sign in the PDF.
        pct_positions = []
        for i, x in enumerate(vals):
            if re.fullmatch(r"(?:\d+(?:\.\d+)?%|\d+\.\d+)", x):
                pct_positions.append(i)

        if not pct_positions:
            continue

        p1i = pct_positions[0]
        p2i = pct_positions[1] if len(pct_positions) > 1 else None

        pre = vals[:p1i]
        country = ""
        shareholder = ""

        if pre:
            first_pre = pre[0].strip()
            if first_pre.upper() in country_names:
                country = first_pre
                shareholder = " ".join(pre[1:]).strip()
            else:
                shareholder = " ".join(pre).strip()

        pct1 = vals[p1i]
        pct2 = vals[p2i] if p2i is not None else ""

        # Everything between the two percentages is the first relationship;
        # everything after the second percentage is the second relationship.
        between = vals[p1i + 1:p2i] if p2i is not None else vals[p1i + 1:]
        after2 = vals[p2i + 1:] if p2i is not None else []

        rel1 = " ".join(between).strip()
        rel2 = " ".join(after2).strip()

        # Remove trailing page number.
        rel2 = re.sub(r"\b(?:7|8|9|10|11|12)\b$", "", rel2).strip()

        rows.append({
            "company_name": company,
            "shareholder_row": row_no,
            "source_country_or_organisation": country,
            "shareholder_name": shareholder,
            "percentage_shareholding_previous": money_to_float(pct1.replace("%","")),
            "percentage_shareholding_previous_raw": pct1,
            "relationship_previous": rel1,
            "percentage_shareholding_current": money_to_float(pct2.replace("%","")) if pct2 else np.nan,
            "percentage_shareholding_current_raw": pct2,
            "relationship_current": rel2,
            "source_page": section_page(page_texts, r"1\.1:\s*SHAREHOLDING STRUCTURE")
        })

    return pd.DataFrame(rows)

EQUITY_ROW_PATTERNS = [
    ("Paid-up Share Capital", r"Paid-up Share Capital"),
    ("Share Premium", r"Share Premium"),
    ("Reserves", r"Reserves\s*\(Capital,\s*Statutory,\s*revaluation,\s*&?\s*Other\)"),
    ("Other Equity", r"Other Equity"),
    ("Accumulated Retained Earnings/Loss", r"Accumulated\s+Retained Earnings/Loss"),
]

def parse_table_cell_sequence(block):
    lines = [normalise_line(x) for x in block.splitlines() if normalise_line(x)]
    # Remove footnote and table heading fragments.
    vals = value_tokens_from_lines(lines, include_missing=True)
    return [cell_state(v) for v in vals]

def assign_equity_cells(cells):
    """
    Map the questionnaire's A, B, C, D1, D2, D3, E cells.
    Because blank cells disappear in extracted PDF text, mapping is based
    on the observed number/order of explicit table cells.
    """
    result = {
        "closing_2024": np.nan,
        "closing_2024_raw": "",
        "closing_2024_status": "",
        "purchase_increase_2025": np.nan,
        "purchase_increase_2025_raw": "",
        "purchase_increase_2025_status": "",
        "sales_decrease_2025": np.nan,
        "sales_decrease_2025_raw": "",
        "sales_decrease_2025_status": "",
        "other_price": np.nan,
        "other_price_raw": "",
        "other_price_status": "",
        "other_exchange_rate": np.nan,
        "other_exchange_rate_raw": "",
        "other_exchange_rate_status": "",
        "other_volume": np.nan,
        "other_volume_raw": "",
        "other_volume_status": "",
        "closing_2025": np.nan,
        "closing_2025_raw": "",
        "closing_2025_status": "",
    }
    n = len(cells)

    def put(name, c):
        result[name] = c["value"]
        result[name + "_raw"] = c["raw"]
        result[name + "_status"] = c["status"]

    if n == 0:
        return result
    if n == 1:
        put("closing_2024", cells[0])
    elif n == 2:
        put("closing_2024", cells[0]); put("closing_2025", cells[1])
    elif n == 3:
        # Common pattern: A, B/C/D blank, E; the middle explicit token is
        # usually B when a dash was entered.
        put("closing_2024", cells[0]); put("purchase_increase_2025", cells[1]); put("closing_2025", cells[2])
    elif n == 4:
        # Common: A, B, C, E
        for name, c in zip(
            ["closing_2024", "purchase_increase_2025", "sales_decrease_2025", "closing_2025"],
            cells
        ):
            put(name, c)
    elif n == 5:
        # Common: A, B, C, D1, E
        for name, c in zip(
            ["closing_2024", "purchase_increase_2025", "sales_decrease_2025", "other_price", "closing_2025"],
            cells
        ):
            put(name, c)
    else:
        names = [
            "closing_2024", "purchase_increase_2025", "sales_decrease_2025",
            "other_price", "other_exchange_rate", "other_volume", "closing_2025"
        ]
        for name, c in zip(names, cells[:7]):
            put(name, c)
    return result


def _equity_row_from_table(row, company, source_page):
    """
    Parse one equity table data row using the actual questionnaire columns.

    Standard PyMuPDF extraction of the FPC equity table is:
    0 label | 1 A | 2 B | 3 C | 4 D1 | 5 spacer | 6 D2 | 7 D3 | 8 spacer | 9 E
    """
    r = list(row) + [None] * max(0, 10 - len(row))
    label = clean_table_cell(r[0])
    if not label:
        return None

    label_low = re.sub(r"\s+", " ", label).lower()
    matched = None
    for canonical, pattern in EQUITY_ROW_PATTERNS:
        if re.search(pattern, label_low, re.I):
            matched = canonical
            break
    if not matched:
        return None

    indices = {
        "closing_2024": 1,
        "purchase_increase_2025": 2,
        "sales_decrease_2025": 3,
        "other_price": 4,
        "other_exchange_rate": 6,
        "other_volume": 7,
        "closing_2025": 9,
    }

    out = {
        "company_name": company,
        "equity_type": matched,
        "source_page": source_page,
    }

    for field, idx in indices.items():
        state = semantic_status_for_table_cell(r[idx] if idx < len(r) else "")
        out[field] = state["value"]
        out[field + "_raw"] = state["raw"]
        out[field + "_status"] = state["status"]

    # Reconciliation is transparent. We calculate an implied D only when all
    # required balances are explicit or explicitly N/A/dash. The output also
    # records whether zero fallback was required.
    states = {
        k: semantic_status_for_table_cell(r[idx] if idx < len(r) else "")
        for k, idx in indices.items()
    }
    a, a_fb = numeric_with_zero_fallback(states["closing_2024"])
    b, b_fb = numeric_with_zero_fallback(states["purchase_increase_2025"])
    c, c_fb = numeric_with_zero_fallback(states["sales_decrease_2025"])
    e, e_fb = numeric_with_zero_fallback(states["closing_2025"])

    if all(not pd.isna(v) for v in [a, b, c, e]):
        out["implied_other_changes"] = float(e - (a + b - c))
        out["implied_other_changes_basis"] = (
            "Explicit A/B/C/E"
            if not any([a_fb, b_fb, c_fb, e_fb])
            else "A/B/C/E with N/A or dash treated as zero for reconciliation"
        )
    else:
        out["implied_other_changes"] = np.nan
        out["implied_other_changes_basis"] = "Insufficient explicit balance cells"

    return out


def parse_equity_table(text, company, page_texts, page_tables=None):
    """
    Parse Section 1.2 using PDF table geometry first.

    This is a critical correction over earlier versions: blank cells are not
    allowed to disappear and shift subsequent values into another equity row.
    Each row is anchored by its printed equity-type label.
    """
    if page_tables:
        candidates = find_all_tables(
            page_tables,
            ["closing balance", "purchase", "sales", "share premium"]
        )
        for page_no, table in candidates:
            rows = []
            for raw_row in table:
                parsed = _equity_row_from_table(raw_row, company, page_no)
                if parsed:
                    rows.append(parsed)
            if rows:
                return pd.DataFrame(rows)

    # Text fallback for scanned/irregular PDFs.
    sec = section_between(
        text,
        r"1\.2:\s*EQUITY INVESTMENT IN YOUR COMPANY",
        [r"1\.3:\s*PROFITS, DIVIDENDS AND RETAINED EARNINGS"]
    )
    if not sec:
        return pd.DataFrame()

    starts = []
    for label, pat in EQUITY_ROW_PATTERNS:
        m = re.search(pat, sec, re.I)
        if m:
            starts.append((m.start(), label, m.end()))
    starts.sort()

    rows = []
    for j, (pos, label, label_end) in enumerate(starts):
        end = starts[j + 1][0] if j + 1 < len(starts) else len(sec)
        block = sec[label_end:end]
        cells = [cell_state(x) for x in VALUE_RE.findall(block)]
        # VALUE_RE may pick table years/headings in unusual OCR. Keep only the
        # first seven plausible cells after the equity label.
        cells = cells[:7]

        mapped = assign_equity_cells(cells)
        mapped.update({
            "company_name": company,
            "equity_type": label,
            "source_page": section_page(page_texts, r"1\.2:\s*EQUITY INVESTMENT")
        })

        states = {
            k: cell_state(mapped.get(k + "_raw", ""))
            for k in [
                "closing_2024", "purchase_increase_2025",
                "sales_decrease_2025", "closing_2025"
            ]
        }
        a, af = numeric_with_zero_fallback(states["closing_2024"])
        b, bf = numeric_with_zero_fallback(states["purchase_increase_2025"])
        c, cf = numeric_with_zero_fallback(states["sales_decrease_2025"])
        e, ef = numeric_with_zero_fallback(states["closing_2025"])
        if all(not pd.isna(v) for v in [a, b, c, e]):
            mapped["implied_other_changes"] = float(e - (a + b - c))
            mapped["implied_other_changes_basis"] = (
                "Explicit A/B/C/E"
                if not any([af, bf, cf, ef])
                else "A/B/C/E with N/A or dash treated as zero for reconciliation"
            )
        else:
            mapped["implied_other_changes"] = np.nan
            mapped["implied_other_changes_basis"] = "Insufficient explicit balance cells"

        rows.append(mapped)

    return pd.DataFrame(rows)


def _profit_row_from_table(row, company, source_page, assets=False):
    r = [clean_table_cell(x) for x in row]
    if len(r) < 4:
        return None

    # Profit tables have a numeric/semantic row underneath a multi-line header.
    # Identify it by the first four cells rather than by row number.
    first_four = r[:4]
    if all(x == "" for x in first_four):
        return None
    if re.search(r"(?i)net profit|dividends|official use|during", " ".join(first_four)):
        return None

    states = [semantic_status_for_table_cell(x) for x in first_four]
    if not any(st["status"] in {"VALUE", "N/A", "DASH"} for st in states):
        return None

    names = [
        "net_profit_after_tax",
        "dividends_declared",
        "dividends_paid_or_remitted",
        "retained_earnings_reported",
    ]
    out = {
        "company_name": company,
        "source_page": source_page,
    }
    for name, state in zip(names, states):
        out[name] = state["value"]
        out[name + "_raw"] = state["raw"]
        out[name + "_status"] = state["status"]

    # For the official retained-earnings formula, N/A and '-' can be treated
    # as zero for reconciliation, but this is explicitly flagged.
    a, af = numeric_with_zero_fallback(states[0])
    b, bf = numeric_with_zero_fallback(states[1])
    d = states[3]["value"]

    if not pd.isna(a) and not pd.isna(b):
        out["retained_earnings_calculated"] = float(a - b)
        out["retained_earnings_calculation_basis"] = (
            "A-B"
            if not (af or bf)
            else "A-B with N/A or dash treated as zero"
        )
    else:
        out["retained_earnings_calculated"] = np.nan
        out["retained_earnings_calculation_basis"] = "Insufficient A/B values"

    out["retained_earnings_difference"] = (
        float(d - out["retained_earnings_calculated"])
        if not pd.isna(d) and not pd.isna(out["retained_earnings_calculated"])
        else np.nan
    )
    return out


def parse_profit_cells(text, company, page_texts, assets=False, page_tables=None):
    start_pat = (
        r"2\.2:\s*PROFITS, DIVIDENDS AND RETAINED EARNINGS"
        if assets
        else r"1\.3:\s*PROFITS, DIVIDENDS AND RETAINED EARNINGS"
    )
    end_pats = (
        [r"2\.4:\s*NON-EQUITY ASSETS", r"PART 3:"]
        if assets
        else [r"1\.4:\s*NON-EQUITY LIABILITIES", r"PART 2:"]
    )

    if page_tables:
        candidates = find_all_tables(
            page_tables,
            ["net profit", "dividends"]
        )
        for page_no, table in candidates:
            # Exclude the equity table because it also contains a "purchase"
            # header and can otherwise be a false candidate.
            txt = table_text(table).lower()
            if "net profit" not in txt:
                continue
            parsed_rows = []
            for row in table:
                parsed = _profit_row_from_table(
                    row, company, page_no, assets=assets
                )
                if parsed:
                    parsed_rows.append(parsed)
            if parsed_rows:
                # A page can contain several visually blank rows. The actual
                # data row is normally the one with the strongest number of
                # populated semantic cells.
                parsed_rows.sort(
                    key=lambda z: sum(
                        not pd.isna(z.get(c))
                        for c in [
                            "net_profit_after_tax",
                            "dividends_declared",
                            "dividends_paid_or_remitted",
                            "retained_earnings_reported",
                        ]
                    ),
                    reverse=True
                )
                return pd.DataFrame([parsed_rows[0]])

    # Text fallback.
    sec = section_between(text, start_pat, end_pats)
    if not sec:
        return pd.DataFrame()

    lines = [normalise_line(x) for x in sec.splitlines() if normalise_line(x)]
    d_idx = next(
        (i for i, x in enumerate(lines) if re.search(r"Official Use Only", x, re.I)),
        None
    )
    candidate_lines = lines[d_idx + 1:] if d_idx is not None else lines
    cells = [cell_state(x) for x in value_tokens_from_lines(candidate_lines, True)]
    cells = [
        c for c in cells
        if not (c["status"] == "VALUE" and c["value"] in {8, 9, 10, 11, 12})
    ]

    names = [
        "net_profit_after_tax",
        "dividends_declared",
        "dividends_paid_or_remitted",
        "retained_earnings_reported",
    ]
    row = {"company_name": company, "source_page": section_page(page_texts, start_pat)}
    for i, name in enumerate(names):
        c = cells[i] if i < len(cells) else {
            "raw": "", "value": np.nan, "status": "BLANK"
        }
        row[name] = c["value"]
        row[name + "_raw"] = c["raw"]
        row[name + "_status"] = c["status"]

    a, af = numeric_with_zero_fallback(cell_state(row["net_profit_after_tax_raw"]))
    b, bf = numeric_with_zero_fallback(cell_state(row["dividends_declared_raw"]))
    if not pd.isna(a) and not pd.isna(b):
        row["retained_earnings_calculated"] = a - b
        row["retained_earnings_calculation_basis"] = (
            "A-B" if not (af or bf) else "A-B with N/A or dash treated as zero"
        )
    else:
        row["retained_earnings_calculated"] = np.nan
        row["retained_earnings_calculation_basis"] = "Insufficient A/B values"

    d = row["retained_earnings_reported"]
    row["retained_earnings_difference"] = (
        d - row["retained_earnings_calculated"]
        if not pd.isna(d) and not pd.isna(row["retained_earnings_calculated"])
        else np.nan
    )
    return pd.DataFrame([row])

def _parse_non_equity_page_coordinate(words, page_width, page_height, company, assets=False, xbounds=None):
    """Extract one fixed-layout non-equity table page using word coordinates."""
    if not words:
        return []

    col_names = ["instrument", "country", "relationship", "maturity", "A", "B", "C", "D1", "D2", "D3", "E", "G"]
    if xbounds and len(xbounds) >= 13:
        bounds = list(xbounds[:13])
    else:
        bounds = [0, 155, 215, 255, 330, 398, 452, 505, 545, 600, 640, 705, page_width + 1]
    instrument_names = [
        "Loans (Including Financial Leases, Repos)",
        "Debt securities (Including Money Market Instruments, Bonds and notes).",
        "Suppliers/Trade Credits & Advances",
        "Life & Non-Life Insurance Technical Reserves",
        "Pension Entitlements/Claims",
        "Other Accounts Payable",
    ]
    instrument_starts = {
        "loans": instrument_names[0],
        "debt": instrument_names[1],
        "suppliers": instrument_names[2],
        "life": instrument_names[3],
        "pension": instrument_names[4],
        "other accounts": instrument_names[5],
    }

    clean_words = []
    for w in words:
        x0,y0,x1,y1,txt = w[:5]
        txt = normalise_line(txt)
        if not txt:
            continue
        yc=(float(y0)+float(y1))/2
        if yc < 145 or yc > page_height - 65:
            continue
        clean_words.append((float(x0),float(x1),yc,txt))

    def col_for_x(x):
        for i in range(len(bounds)-1):
            if bounds[i] <= x < bounds[i+1]:
                return col_names[i]
        return "G"

    # Primary anchors come from financial cells. These distinguish adjacent
    # questionnaire rows even when text in a narrow source cell wraps.
    financial_cols = {"A","B","C","D","E","G"}
    anchor_ys=[]
    for x0,x1,yc,txt in clean_words:
        if col_for_x((x0+x1)/2) in financial_cols:
            # Only words that look like a value/text entry, not header fragments.
            if txt:
                anchor_ys.append(yc)

    def cluster(values, tol=3.5):
        vals=sorted(values)
        out=[]
        for v in vals:
            if not out or abs(v-out[-1])>tol:
                out.append(v)
            else:
                out[-1]=(out[-1]+v)/2
        return out

    anchors=cluster(anchor_ys,3.5)
    if not anchors:
        secondary=[]
        for x0,x1,yc,txt in clean_words:
            col=col_for_x((x0+x1)/2)
            if col in {"country","relationship","maturity"}:
                secondary.append(yc)
        anchors=cluster(secondary,3.5)

    # If a page has no anchors at all, nothing can be safely extracted.
    if not anchors:
        return []

    grouped={i:[] for i in range(len(anchors))}
    for word in clean_words:
        x0,x1,yc,txt=word
        col=col_for_x((x0+x1)/2)
        idx=min(range(len(anchors)), key=lambda i:abs(anchors[i]-yc))
        distance=abs(anchors[idx]-yc)

        # Wrapped text in the source-country column often extends below the
        # numeric row anchor. Parenthetical continuation words belong to the
        # preceding source cell, not the next country row.
        if col == "country" and idx > 0:
            prev_idx=idx-1
            prev_text=" ".join(w[3] for w in grouped.get(prev_idx, []) if col_for_x((w[0]+w[1])/2)=="country")
            if (txt.startswith("(") or txt in {"materials)", "Africa", "Burley)"}) and prev_text and abs(anchors[prev_idx]-yc) <= 22:
                grouped[prev_idx].append(word)
                continue

        if distance <= 10.5:
            grouped[idx].append(word)

    def parse_cell(txt):
        txt=normalise_line(txt)
        if not txt:
            return {"raw":"","value":np.nan,"status":"BLANK"}
        return cell_state(txt)

    rows=[]
    current_instrument=""
    previous_country=""

    for idx in sorted(grouped):
        row_words=sorted(grouped[idx],key=lambda z:(z[2],z[0]))
        if not row_words:
            continue
        cells={c:[] for c in col_names}
        d_sub={"D1":[],"D2":[],"D3":[]}
        for x0,x1,yc,txt in row_words:
            xc=(x0+x1)/2
            col=col_for_x(xc)
            cells[col].append(txt)
        cells={k:" ".join(v).strip() for k,v in cells.items()}
        inst_text=cells["instrument"]
        inst_low=inst_text.lower()
        for start_text,full_name in instrument_starts.items():
            if inst_low.startswith(start_text):
                current_instrument=full_name
                break
        if not current_instrument:
            continue
        if any(x in inst_low for x in ["type of loan","closing balance","official use"]):
            continue

        country=cells["country"]
        relationship=cells["relationship"]
        maturity=cells["maturity"]

        # Source-country cells in the questionnaire can wrap onto 2–3 physical
        # lines (e.g. "South Africa (Wrapping materials)"). Keep the source
        # wording intact rather than allowing the continuation to leak into the
        # next logical row.
        country = re.sub(r"\s+", " ", country).strip()
        relationship = re.sub(r"\s+", " ", relationship).strip()
        maturity = re.sub(r"\s+", " ", maturity).strip()
        parsed={k:parse_cell(cells[k]) for k in ["A","B","C","E","G"]}
        for k in ["D1","D2","D3"]:
            parsed[k]=parse_cell(cells[k])

        data_cells=[parsed[k] for k in ["A","B","C","D1","D2","D3","E","G"]]
        has_value=any(c["status"]=="VALUE" for c in data_cells)
        has_na=any(c["status"]=="N/A" for c in data_cells)

        # Preserve textual entries such as None/Paid up when no numeric value
        # exists, but do not mistake them for financial amounts.
        if country.upper() in {"AFRICA","WORLD"} and previous_country and not has_value:
            country=f"{previous_country} {country}".strip()
            if rows and rows[-1].get("instrument")==current_instrument:
                rows[-1]["country_or_organisation"]=country
                rows[-1]["raw_line"]=(rows[-1].get("raw_line","")+" "+cells["country"]).strip()
                previous_country=country
                continue

        if country and country.upper() not in {"N/A","NONE"}:
            previous_country=country

        has_context=bool(country or relationship or maturity or cells["A"] or cells["E"])
        if not has_value and not has_na and not has_context:
            continue

        has_text_financial = any(
            c["status"] not in {"VALUE", "N/A", "BLANK", "DASH"}
            for c in data_cells
        )
        status = (
            "VALUE" if has_value else
            "REVIEW" if has_text_financial else
            "N/A" if has_na or country.upper() in {"N/A","NONE"} else
            "BLANK"
        )
        raw_values=[]; numeric_values=[]
        for k in ["A","B","C","D1","D2","D3","E","G"]:
            c=parsed[k]
            if c["raw"]: raw_values.append(f"{k}={c['raw']}")
            if not pd.isna(c["value"]): numeric_values.append(f"{k}={c['value']}")

        raw_line=" | ".join(x for x in [inst_text,country,relationship,maturity,cells["A"],cells["B"],cells["C"]," ".join(cells[k] for k in ["D1","D2","D3"] if cells[k]),cells["E"],cells["G"]] if x)
        rows.append({
            "company_name":company,"instrument":current_instrument,
            "country_or_organisation":country,"relationship":relationship,"maturity":maturity,
            "closing_balance_2024":parsed["A"]["value"],"closing_balance_2024_raw":parsed["A"]["raw"],
            "amount_received_2025":parsed["B"]["value"],"amount_received_2025_raw":parsed["B"]["raw"],
            "principal_repayment_2025":parsed["C"]["value"],"principal_repayment_2025_raw":parsed["C"]["raw"],
            "other_price":parsed["D1"]["value"],"other_price_raw":parsed["D1"]["raw"],
            "other_exchange_rate":parsed["D2"]["value"],"other_exchange_rate_raw":parsed["D2"]["raw"],
            "other_volume":parsed["D3"]["value"],"other_volume_raw":parsed["D3"]["raw"],
            "closing_balance_2025":parsed["E"]["value"],"closing_balance_2025_raw":parsed["E"]["raw"],
            "interest_paid_2025":parsed["G"]["value"],"interest_paid_2025_raw":parsed["G"]["raw"],
            "raw_line":raw_line,"numeric_values_raw":" | ".join(raw_values),
            "numeric_values":" | ".join(numeric_values),"status":status,
        })
    # Repair a small class of wrapped source-country cells. This is intentionally
    # conservative and only acts on parenthetical continuation text.
    for i in range(1, len(rows)):
        prev = rows[i-1]
        cur = rows[i]
        if prev.get("instrument") == cur.get("instrument"):
            c = str(cur.get("country_or_organisation") or "").strip()
            if c.startswith("(") and ")" in c:
                prev["country_or_organisation"] = (
                    str(prev.get("country_or_organisation") or "").strip() + " " + c
                ).strip()
                # If the current row carries no financial values, it is only a
                # wrapped text fragment and can be removed.
                if cur.get("status") == "BLANK":
                    rows[i] = None
    rows = [r for r in rows if r is not None]

    # Merge vertically wrapped numeric cells that belong to the same logical
    # row. Some questionnaire PDFs split values such as 45,374,814.28 into
    # two text lines. A fragment row has no country/relationship/maturity and
    # therefore can safely be merged into the preceding row.
    merged=[]
    for r in rows:
        context = any(str(r.get(k) or "").strip() for k in ["country_or_organisation","relationship","maturity"])
        value_keys=["closing_balance_2024","amount_received_2025","principal_repayment_2025","other_price","other_exchange_rate","other_volume","closing_balance_2025","interest_paid_2025"]
        has_num=any(not pd.isna(r.get(k)) for k in value_keys)
        raw_fragments = [str(r.get(k+"_raw") or "").strip() for k in value_keys]
        is_numeric_fragment = any(x.startswith(".") or x.endswith(",") or x in {"000.00","261.04","24"} for x in raw_fragments if x)
        if merged and r.get("instrument")==merged[-1].get("instrument") and has_num and not context and is_numeric_fragment:
            prev=merged[-1]
            for k in value_keys:
                cur_raw=str(r.get(k+"_raw") or "").strip()
                if not cur_raw:
                    continue
                prev_raw=str(prev.get(k+"_raw") or "").strip()
                if (cur_raw.startswith(".") or cur_raw in {"000.00","261.04"} or (prev_raw.endswith(".") and cur_raw.isdigit())) and prev_raw:
                    combined=prev_raw+cur_raw
                    try:
                        prev[k]=money_to_float(combined)
                    except Exception:
                        pass
                    prev[k+"_raw"]=combined
                elif cur_raw.endswith(",") and prev_raw:
                    prev[k+"_raw"]=(prev_raw+cur_raw).strip()
                    try:
                        prev[k]=money_to_float(prev[k+"_raw"])
                    except Exception:
                        pass
                elif pd.isna(prev.get(k)) and not pd.isna(r.get(k)):
                    prev[k]=r[k]
                    prev[k+"_raw"]=cur_raw
            prev["raw_line"]=(str(prev.get("raw_line") or "")+" "+str(r.get("raw_line") or "")).strip()
            prev["numeric_values_raw"]=(str(prev.get("numeric_values_raw") or "")+" | "+str(r.get("numeric_values_raw") or "")).strip(" |")
            prev["numeric_values"]=(str(prev.get("numeric_values") or "")+" | "+str(r.get("numeric_values") or "")).strip(" |")
            continue
        merged.append(r)
    return merged



def _canonical_instrument_label(value):
    x = re.sub(r"\s+", " ", clean_table_cell(value)).strip()
    low = x.lower()
    if "loans" in low and "financial leases" in low:
        return "Loans (Including Financial Leases, Repos)"
    if "debt securities" in low:
        return "Debt securities (Including Money Market Instruments, Bonds and notes)."
    if "suppliers" in low and "trade" in low:
        return "Suppliers/Trade Credits & Advances"
    if "life" in low and "insurance" in low:
        return "Life & Non-Life Insurance Technical Reserves"
    if "pension" in low:
        return "Pension Entitlements/Claims"
    if "other accounts" in low and "payable" in low:
        return "Other Accounts Payable"
    return x



def _infer_non_equity_columns(table):
    """Infer semantic columns from the questionnaire's multi-row header."""
    n = max((len(r) for r in table), default=0)
    headers = [""] * n
    for c in range(n):
        pieces = []
        for r in table[:10]:
            if c < len(r) and clean_table_cell(r[c]):
                pieces.append(clean_table_cell(r[c]))
        headers[c] = " ".join(pieces).lower()

    def first_where(predicate):
        for c, h in enumerate(headers):
            if predicate(h):
                return c
        return None

    cols = {
        "instrument": 0,
        "country": 1,
        "relationship": 2,
        "maturity": 3,
        "A": first_where(
            lambda h: "closing balance" in h and "2024" in h
        ),
        "B": first_where(lambda h: "amount received" in h),
        "C": first_where(
            lambda h: "principal" in h and "repayme" in h
        ),
        "E": first_where(
            lambda h: "closing balance" in h and "2025" in h
        ),
        "G": first_where(lambda h: "interest paid" in h),
        "D1": first_where(
            lambda h: re.search(r"\bd1\b", h) and "price" in h
        ),
        "D2": first_where(
            lambda h: re.search(r"\bd2\b", h) and "exchange" in h
        ),
        "D3": first_where(
            lambda h: re.search(r"\bd3\b", h) and "volume" in h
        ),
    }

    # In several questionnaire PDFs the A header is in a merged cell one
    # column to the right of the actual data cell. If the inferred A column
    # contains no row data but its immediate predecessor does, shift A left.
    if cols["A"] is not None and cols["A"] > 0:
        a_col = cols["A"]
        candidate_score = sum(
            bool(clean_table_cell(r[a_col]))
            for r in table[10:]
            if a_col < len(r)
        )
        left_score = sum(
            bool(clean_table_cell(r[a_col - 1]))
            for r in table[10:]
            if a_col - 1 < len(r)
        )
        if left_score > candidate_score:
            cols["A"] = a_col - 1

    # E can also be duplicated by a merged header. Choose the candidate with
    # the strongest population below the header.
    e_candidates = [
        c for c, h in enumerate(headers)
        if "closing balance" in h and "2025" in h
    ]
    if len(e_candidates) > 1:
        cols["E"] = max(
            e_candidates,
            key=lambda c: sum(
                bool(clean_table_cell(r[c]))
                for r in table[10:] if c < len(r)
            )
        )

    return cols

def _parse_non_equity_table_geometry(table, company, source_page, assets=False):
    if not table:
        return []

    cols = _infer_non_equity_columns(table)
    required = ["A", "B", "C", "E", "G"]
    if not all(cols.get(k) is not None for k in required):
        return []

    rows = []
    current_instrument = ""
    for raw_row in table[1:]:
        max_col = max([c for c in cols.values() if c is not None], default=0)
        row = list(raw_row) + [None] * max(0, max_col + 1 - len(raw_row))
        first = clean_table_cell(row[cols["instrument"]])
        if first:
            canonical = _canonical_instrument_label(first)
            if canonical:
                current_instrument = canonical

        # Ignore repeated multi-line header fragments.
        if not current_instrument:
            continue

        country = clean_table_cell(row[cols["country"]])
        relationship = clean_table_cell(row[cols["relationship"]])
        maturity = clean_table_cell(row[cols["maturity"]])

        # Normalise common PDF truncations.
        relationship = re.sub(r"^Othe$", "Other", relationship, flags=re.I)
        relationship = re.sub(r"^Parent\s+Compan[y|y]?$", "Parent Company", relationship, flags=re.I)
        if country.upper() == "SOUTH":
            country = "South Africa"
        elif country.upper() == "SOUTH AFRICA":
            country = "South Africa"
        elif country.upper() == "ZAMBIA":
            country = "Zambia"

        field_map = {
            "closing_balance_2024": "A",
            "amount_received_2025": "B",
            "principal_repayment_2025": "C",
            "other_price": "D1",
            "other_exchange_rate": "D2",
            "other_volume": "D3",
            "closing_balance_2025": "E",
            "interest_paid_2025": "G",
        }

        out = {
            "company_name": company,
            "instrument": current_instrument,
            "country_or_organisation": country,
            "relationship": relationship,
            "maturity": maturity,
            "source_page": source_page,
        }

        statuses = []
        financial_raw = []
        has_any_content = bool(country or relationship or maturity or first)

        for field, col_key in field_map.items():
            actual_col = cols.get(col_key)
            if actual_col is None or actual_col >= len(row):
                state = {"value": np.nan, "raw": "", "status": "BLANK"}
            else:
                state = semantic_status_for_table_cell(row[actual_col])
            out[field] = state["value"]
            out[field + "_raw"] = state["raw"]
            statuses.append(state["status"])
            if state["raw"]:
                financial_raw.append(f"{field}={state['raw']}")
                has_any_content = True

        # A purely empty continuation row is not a record.
        if not has_any_content:
            continue

        # Rows that contain only the printed instrument heading and no response
        # are not useful records. Explicit N/A is retained.
        nonblank_financial = any(
            st in {"VALUE", "N/A", "DASH", "TEXT"} for st in statuses
        )
        if not nonblank_financial and not (country or relationship or maturity):
            continue

        status = "VALUE" if any(st == "VALUE" for st in statuses) else (
            "REVIEW" if any(st == "TEXT" for st in statuses) else "N/A"
        )

        out["raw_line"] = " | ".join(
            [x for x in [current_instrument, country, relationship, maturity] if x]
            + financial_raw
        )
        out["numeric_values_raw"] = " | ".join(financial_raw)
        nums = []
        for field in field_map:
            if not pd.isna(out[field]):
                nums.append(f"{field}={out[field]}")
        out["numeric_values"] = " | ".join(nums)
        out["status"] = status

        # If the source contains a suspicious text entry such as "Paid up" in
        # a financial column, make it explicit for review rather than coercing it.
        if any(st == "TEXT" for st in statuses):
            out["status"] = "REVIEW"

        rows.append(out)

    return rows


def parse_non_equity_table(
    text, company, page_texts, assets=False,
    page_words=None, page_table_xbounds=None, page_sizes=None, page_tables=None
):
    """
    Parse non-equity liabilities/assets from actual PDF tables first.

    The table parser preserves wrapped numbers and continuation rows, and it
    does not interpret page numbers as financial values.
    """
    heading = (
        r"2\.4:\s*NON-EQUITY ASSETS"
        if assets else r"1\.4:\s*NON-EQUITY LIABILITIES"
    )

    if page_tables:
        candidates = find_all_tables(
            page_tables,
            ["type of loan", "closing balance"]
        )
        # The table may span a page break; choose the table on the page that
        # actually contains the section heading.
        target_page = section_page(page_texts, heading)
        ordered = sorted(
            candidates,
            key=lambda x: (0 if x[0] == target_page else 1, x[0])
        )
        for page_no, table in ordered:
            rows = _parse_non_equity_table_geometry(
                table, company, page_no, assets=assets
            )
            if rows:
                return pd.DataFrame(rows)

    # Coordinate fallback retained for PDFs where table extraction fails.
    if page_words:
        for page_no, page_text in enumerate(page_texts, start=1):
            if re.search(heading, page_text or "", re.I):
                words = page_words[page_no - 1] if page_no - 1 < len(page_words) else []
                xbounds = (
                    page_table_xbounds[page_no - 1]
                    if page_table_xbounds and page_no - 1 < len(page_table_xbounds)
                    else None
                )
                if page_sizes and page_no - 1 < len(page_sizes):
                    page_width, page_height = page_sizes[page_no - 1]
                else:
                    page_width = max([float(w[2]) for w in words], default=612.0)
                    page_height = max([float(w[3]) for w in words], default=600.0) + 20.0
                rows = _parse_non_equity_page_coordinate(
                    words,
                    page_width,
                    page_height,
                    company,
                    assets=assets,
                    xbounds=xbounds,
                )
                if rows:
                    df = pd.DataFrame(rows)
                    df["source_page"] = page_no
                    return df
                break

    # Final conservative fallback: return only a section-level review record,
    # not fabricated financial numbers.
    sec = section_between(text, heading, [r"PART 2:", r"PART 3:"])
    if not sec:
        return pd.DataFrame()
    return pd.DataFrame([{
        "company_name": company,
        "instrument": "",
        "country_or_organisation": "",
        "relationship": "",
        "maturity": "",
        "closing_balance_2024": np.nan,
        "closing_balance_2024_raw": "",
        "amount_received_2025": np.nan,
        "amount_received_2025_raw": "",
        "principal_repayment_2025": np.nan,
        "principal_repayment_2025_raw": "",
        "other_price": np.nan,
        "other_price_raw": "",
        "other_exchange_rate": np.nan,
        "other_exchange_rate_raw": "",
        "other_volume": np.nan,
        "other_volume_raw": "",
        "closing_balance_2025": np.nan,
        "closing_balance_2025_raw": "",
        "interest_paid_2025": np.nan,
        "interest_paid_2025_raw": "",
        "raw_line": "Section detected but table geometry could not be safely extracted",
        "numeric_values_raw": "",
        "numeric_values": "",
        "status": "REVIEW",
        "source_page": section_page(page_texts, heading),
    }])

SERVICE_CODES = list("abcdefghijklm")

SERVICE_CODES = list("abcdefghijklm")

SERVICE_LABELS = {
    "a": "Manufacturing services on physical inputs owned by others",
    "b": "Maintenance and repair services n.i.e.",
    "c": "Computer services",
    "d": "Information services",
    "e": "Insurance premiums",
    "f": "Insurance claims",
    "g": "Pension premiums",
    "h": "Pension claims",
    "i": "Commissions, fees and other charges by financial institutions (excluding interest)",
    "j": "Charges for the use of intellectual property",
    "k": "Research and development services",
    "l": "Professional and management consulting services",
    "m": "Construction services",
}


def _service_row(company, code, label, exp_state, imp_state,
                 unallocated=None, source_page=np.nan):
    unallocated = unallocated or {"raw": "", "value": np.nan, "status": "BLANK"}
    return {
        "company_name": company,
        "service_code": code,
        "service_type": label,
        "exports": exp_state["value"],
        "exports_raw": exp_state["raw"],
        "exports_status": exp_state["status"],
        "imports": imp_state["value"],
        "imports_raw": imp_state["raw"],
        "imports_status": imp_state["status"],
        "unallocated_value": unallocated["value"],
        "unallocated_value_raw": unallocated["raw"],
        "unallocated_value_status": unallocated["status"],
        "source_page": source_page,
    }


def parse_services(text, company, page_texts, page_tables=None):
    """
    Parse Section 3.1 using the actual three-column service table where
    available. This removes the previous ambiguity where a single numeric
    value on a flattened text line could not safely be assigned to exports or
    imports.
    """
    rows = []

    # First collect proper 3-column tables. Some questionnaires split the
    # service table across pages; each page can therefore contribute rows.
    if page_tables:
        for page_no, tables in enumerate(page_tables, start=1):
            for table in tables or []:
                if not is_section_table(table, ["type of service", "exports", "imports"]):
                    continue
                for raw_row in table[1:]:
                    r = table_row_clean(raw_row)
                    if not r:
                        continue
                    m = re.match(r"^\(([a-m])\)\s*(.*)$", r[0], re.I | re.S)
                    if not m:
                        continue
                    code = m.group(1).lower()
                    label = re.sub(r"\s+", " ", m.group(2)).strip()
                    label = label or SERVICE_LABELS.get(code, "")
                    exp_raw = r[1] if len(r) > 1 else ""
                    imp_raw = r[2] if len(r) > 2 else ""
                    exp = semantic_status_for_table_cell(exp_raw)
                    imp = semantic_status_for_table_cell(imp_raw)
                    rows.append(
                        _service_row(
                            company,
                            code,
                            label,
                            exp,
                            imp,
                            source_page=page_no,
                        )
                    )

    # Deduplicate service codes when a table extractor returns the same page
    # twice. A later row is retained only if it carries more information.
    if rows:
        by_code = {}
        for row in rows:
            code = row["service_code"]
            score = sum(
                row[k] not in {"", None} and not pd.isna(row[k])
                for k in [
                    "exports", "imports",
                    "exports_raw", "imports_raw"
                ]
            )
            old = by_code.get(code)
            old_score = (
                sum(
                    old[k] not in {"", None} and not pd.isna(old[k])
                    for k in ["exports", "imports", "exports_raw", "imports_raw"]
                )
                if old else -1
            )
            if old is None or score > old_score:
                by_code[code] = row
        # Fill missing service codes from the canonical questionnaire list.
        for code in SERVICE_CODES:
            if code not in by_code:
                by_code[code] = _service_row(
                    company,
                    code,
                    SERVICE_LABELS[code],
                    {"raw": "", "value": np.nan, "status": "BLANK"},
                    {"raw": "", "value": np.nan, "status": "BLANK"},
                    source_page=section_page(
                        page_texts, r"3\.1\s*SERVICE IMPORTS AND EXPORTS"
                    ),
                )
        return pd.DataFrame(
            [by_code[c] for c in SERVICE_CODES]
        )

    # Text fallback. Unlike earlier versions, values can occur on the same
    # line as the service description (e.g. "0 391,680").
    sec = section_between(
        text,
        r"3\.1\s*SERVICE IMPORTS AND EXPORTS",
        [r"NB:\s*PLEASE ATTACH", r"THANK YOU"]
    )
    if not sec:
        return pd.DataFrame()

    sec = re.sub(r"\n\s*\d+\s*$", "", sec, flags=re.M)
    parts = re.split(r"(?m)(?=^\([a-m]\)\s*)", sec, flags=re.I)
    source_page = section_page(
        page_texts, r"3\.1\s*SERVICE IMPORTS AND EXPORTS"
    )

    for part in parts:
        m = re.match(r"^\(([a-m])\)\s*(.*)$", part.strip(), re.I | re.S)
        if not m:
            continue
        code = m.group(1).lower()
        block = m.group(2)
        label = SERVICE_LABELS.get(code, "")
        # Strip known table headings/years and page numbers.
        cleaned = re.sub(r"\b2025\b", " ", block)
        cleaned = re.sub(r"\b(?:8|9|10|11|12)\b(?=\s*$)", " ", cleaned)
        token_pattern = (
            r"(?i)(?<![A-Za-z0-9])"
            r"(?:N/?A|N\.?A\.?|NONE|NIL|NILL|N\?A|N//A|N\./A|N\./A\.|--|-|"
            r"\((?:US\$|USS\$|USD|US|\$)?-?\s*\d[\d,]*(?:\.\d+)?\)|"
            r"(?:US\$|USS\$|USD|US|\$)?\(?-?\s*\d[\d,]*(?:\.\d+)?\)?)"
            r"(?![A-Za-z0-9])"
        )
        tokens = re.findall(token_pattern, cleaned)
        cells = [cell_state(x) for x in tokens]

        # The text fallback can accidentally capture tokens from the next
        # service/section. Restrict to at most two source cells.
        cells = cells[:2]
        if len(cells) >= 2:
            exp, imp = cells[0], cells[1]
            unallocated = None
        elif len(cells) == 1:
            exp = {"raw": "", "value": np.nan, "status": "UNRESOLVED"}
            imp = {"raw": "", "value": np.nan, "status": "UNRESOLVED"}
            unallocated = cells[0]
        else:
            exp = {"raw": "", "value": np.nan, "status": "BLANK"}
            imp = {"raw": "", "value": np.nan, "status": "BLANK"}
            unallocated = None

        rows.append(
            _service_row(
                company,
                code,
                label,
                exp,
                imp,
                unallocated=unallocated,
                source_page=source_page,
            )
        )

    return pd.DataFrame(rows)

def parse_comments(text, company, page_texts):
    m = re.search(
        r"GENERAL COMMENTS(.*?)(?=THANK YOU FOR YOUR COOPERATION|$)",
        text, re.I | re.S
    )
    if not m:
        return pd.DataFrame()
    comment = clean_text(m.group(1))
    comment = re.sub(r"Please if you have any comments or feedback related to the institution policies and operations, kindly indicate in space below", "", comment, flags=re.I)
    comment = re.sub(r"NB:\s*PLEASE ATTACH FINANCIAL STATEMENTS.*?GENERAL COMMENTS", "", comment, flags=re.I|re.S)
    comment = comment.strip()
    if not comment:
        return pd.DataFrame()
    return pd.DataFrame([{
        "company_name": company,
        "general_comments": comment,
        "source_page": section_page(page_texts, r"GENERAL COMMENTS")
    }])

def parse_financial_summary(text, company, page_texts):
    profits = parse_profit_cells(text, company, page_texts, assets=False)
    asset_profits = parse_profit_cells(text, company, page_texts, assets=True)
    out = {"company_name": company}
    if not profits.empty:
        for c in profits.columns:
            if c not in {"company_name", "source_page"}:
                out["liability_" + c] = profits.iloc[0][c]
        out["profits_source_page"] = profits.iloc[0]["source_page"]
    if not asset_profits.empty:
        for c in asset_profits.columns:
            if c not in {"company_name", "source_page"}:
                out["asset_" + c] = asset_profits.iloc[0][c]
        out["asset_profits_source_page"] = asset_profits.iloc[0]["source_page"]
    return pd.DataFrame([out])

def run_qa(company_record, equity_df, profit_df, extraction):
    company = company_record.get("company_name", "")
    checks = []

    company_name_source = company_record.get("company_name_source", "PDF_TEXT")
    checks.append({
        "company_name": company,
        "check": "Company name extracted",
        "status": "REVIEW" if (
            not company or company_name_source == "FILENAME_FALLBACK"
        ) else "PASS",
        "value": company,
        "message": (
            "Company name could not be confidently extracted."
            if not company else
            "Company name was derived from the PDF filename because the PDF text layer was unavailable."
            if company_name_source == "FILENAME_FALLBACK"
            else ""
        )
    })

    checks.append({
        "company_name": company,
        "check": "OCR used",
        "status": "REVIEW" if extraction["ocr_pages"] else "PASS",
        "value": len(extraction["ocr_pages"]),
        "message": (
            "OCR was used; verify all OCR-derived values."
            if extraction["ocr_pages"] else "Native PDF text extraction used."
        )
    })

    if extraction.get("ocr_confidences"):
        vals = [x.get("confidence") for x in extraction["ocr_confidences"] if not pd.isna(x.get("confidence"))]
        if vals:
            mean_conf = float(np.mean(vals))
            checks.append({
                "company_name": company,
                "check": "OCR confidence",
                "status": "PASS" if mean_conf >= 70 else "REVIEW",
                "value": mean_conf,
                "message": (
                    "Average OCR confidence is acceptable."
                    if mean_conf >= 70 else
                    "Average OCR confidence is below the review threshold; verify OCR-derived values against the PDF."
                )
            })

    if not extraction["tesseract_available"] and extraction["ocr_pages"]:
        checks.append({
            "company_name": company,
            "check": "OCR engine availability",
            "status": "FAIL",
            "value": False,
            "message": "Scanned pages detected but Tesseract OCR is unavailable."
        })
    else:
        checks.append({
            "company_name": company,
            "check": "OCR engine availability",
            "status": "PASS",
            "value": bool(extraction["tesseract_available"]),
            "message": ""
        })

    if not profit_df.empty:
        d = profit_df.iloc[0]
        diff = d.get("retained_earnings_difference", np.nan)
        if pd.isna(diff):
            status, msg = "INFO", "Insufficient explicit values for reconciliation."
        elif abs(float(diff)) < 0.01:
            status, msg = "PASS", "Reported retained earnings reconcile to A-B."
        else:
            status, msg = "REVIEW", f"Reported vs calculated retained earnings differ by {diff:,.2f}."
        checks.append({
            "company_name": company,
            "check": "Retained earnings reconciliation",
            "status": status,
            "value": diff,
            "message": msg
        })

    if not equity_df.empty:
        for _, r in equity_df.iterrows():
            implied = r.get("implied_other_changes", np.nan)
            basis = r.get("implied_other_changes_basis", "")
            if pd.isna(implied):
                continue

            dvals = [
                r.get("other_price"),
                r.get("other_exchange_rate"),
                r.get("other_volume"),
            ]
            explicit_d = [x for x in dvals if not pd.isna(x)]

            if explicit_d:
                dsum = float(sum(explicit_d))
                residual = dsum - float(implied)
                status = "PASS" if abs(residual) < 0.01 else "REVIEW"
                msg = (
                    "Explicit D1+D2+D3 reconcile to the implied balance movement."
                    if status == "PASS"
                    else
                    f"Explicit D1+D2+D3 differ from the implied balance movement by {residual:,.2f}."
                )
                value = residual
            else:
                if abs(float(implied)) < 0.01:
                    status = "PASS"
                    msg = "Balances reconcile with no implied other change."
                else:
                    status = "INFO"
                    msg = (
                        f"Implied other changes = {float(implied):,.2f}. "
                        f"{basis}. D1-D3 were not explicitly reported."
                    )
                value = implied

            checks.append({
                "company_name": company,
                "check": f"Equity reconciliation - {r.get('equity_type','')}",
                "status": status,
                "value": value,
                "message": msg,
            })

    return pd.DataFrame(checks)



def build_data_quality_summary(result, company):
    """Create a compact, machine-readable extraction quality summary."""
    checks = []
    for key, label in [
        ("industrial", "Industrial classification"),
        ("shareholding", "Shareholding"),
        ("equity", "Equity"),
        ("profits", "Profits"),
        ("asset_profits", "Foreign asset profits"),
        ("non_equity_liabilities", "Non-equity liabilities"),
        ("non_equity_assets", "Non-equity assets"),
        ("services", "Services"),
        ("comments", "Comments"),
    ]:
        df = result.get(key)
        rows = len(df) if isinstance(df, pd.DataFrame) else 0
        checks.append({
            "company_name": company,
            "section": label,
            "rows_extracted": rows,
            "status": "POPULATED" if rows else "EMPTY",
        })

    eq = result.get("equity", pd.DataFrame())
    if isinstance(eq, pd.DataFrame) and not eq.empty:
        review_count = int(
            (eq.filter(regex="_status$").astype("string") == "TEXT").sum().sum()
        )
        if review_count:
            checks.append({
                "company_name": company,
                "section": "Equity",
                "rows_extracted": len(eq),
                "status": "REVIEW",
                "issue": f"{review_count} equity cell(s) contain non-numeric text.",
            })

    ne = result.get("non_equity_liabilities", pd.DataFrame())
    if isinstance(ne, pd.DataFrame) and not ne.empty and "status" in ne.columns:
        review_rows = int((ne["status"].astype("string") == "REVIEW").sum())
        if review_rows:
            checks.append({
                "company_name": company,
                "section": "Non-equity liabilities",
                "rows_extracted": len(ne),
                "status": "REVIEW",
                "issue": f"{review_rows} row(s) require source-PDF review.",
            })

    sv = result.get("services", pd.DataFrame())
    if isinstance(sv, pd.DataFrame) and not sv.empty:
        unresolved = int(
            sv.get("unallocated_value_status", pd.Series(dtype="string"))
              .astype("string").eq("VALUE").sum()
        )
        if unresolved:
            checks.append({
                "company_name": company,
                "section": "Services",
                "rows_extracted": len(sv),
                "status": "REVIEW",
                "issue": f"{unresolved} service row(s) contain value(s) whose export/import side was unresolved.",
            })

    return pd.DataFrame(checks)


def process_extraction_result(extraction):
    """Convert one PDF extraction into the canonical FPC record bundle."""
    text = extraction.get("text", "")
    filename = extraction.get("filename", "")

    company_record = parse_company_info(text)
    company = (company_record.get("company_name") or "").strip()

    # A scanned PDF can have no text at all. The filename is still useful and
    # is explicitly marked as a fallback source rather than presented as a
    # successfully extracted company name.
    company_from_filename = False
    if not company:
        company = canonical_company_name_from_filename(filename)
        company_from_filename = True
    company_record["company_name"] = company
    company_record["company_name_source"] = (
        "FILENAME_FALLBACK" if company_from_filename else "PDF_TEXT"
    )

    page_tables = extraction.get("page_tables", [])
    industrial = parse_industrial(text, company, extraction["page_texts"])
    shareholding = parse_shareholding(text, company, extraction["page_texts"])
    equity = parse_equity_table(
        text, company, extraction["page_texts"], page_tables=page_tables
    )
    profits = parse_profit_cells(
        text, company, extraction["page_texts"], assets=False,
        page_tables=page_tables
    )
    asset_profits = parse_profit_cells(
        text, company, extraction["page_texts"], assets=True,
        page_tables=page_tables
    )
    non_equity_liabilities = parse_non_equity_table(
        text, company, extraction["page_texts"],
        assets=False,
        page_words=extraction.get("page_words"),
        page_table_xbounds=extraction.get("page_table_xbounds"),
        page_sizes=extraction.get("page_sizes"),
        page_tables=page_tables,
    )
    non_equity_assets = parse_non_equity_table(
        text, company, extraction["page_texts"],
        assets=True,
        page_words=extraction.get("page_words"),
        page_table_xbounds=extraction.get("page_table_xbounds"),
        page_sizes=extraction.get("page_sizes"),
        page_tables=page_tables,
    )
    services = parse_services(
        text, company, extraction["page_texts"], page_tables=page_tables
    )
    comments = parse_comments(text, company, extraction["page_texts"])
    qa = run_qa(company_record, equity, profits, extraction)

    company_record.update({
        "source_file": filename,
        "page_count": extraction["page_count"],
        "ocr_pages": ", ".join(map(str, extraction["ocr_pages"])),
        "ocr_page_count": len(extraction["ocr_pages"]),
        "extraction_timestamp": datetime.now().isoformat(timespec="seconds"),
        "company_name_source": (
            "FILENAME_FALLBACK" if company_from_filename else "PDF_TEXT"
        ),
        "table_extraction_pages": ", ".join(
            str(i + 1) for i, tabs in enumerate(page_tables) if tabs
        ),
    })

    # Add explicit extraction-quality QA checks.
    if company_from_filename:
        qa = pd.concat([
            qa,
            pd.DataFrame([{
                "company_name": company,
                "check": "Company name source",
                "status": "REVIEW",
                "value": company,
                "message": "Company name derived from PDF filename because the PDF text layer was unavailable.",
            }])
        ], ignore_index=True)

    if not extraction["tesseract_available"] and extraction["page_count"] > 0:
        native_blank_pages = [
            m["page"] for m in extraction["page_meta"]
            if m["native_characters"] < 80
        ]
        if native_blank_pages:
            qa = pd.concat([
                qa,
                pd.DataFrame([{
                    "company_name": company,
                    "check": "OCR availability",
                    "status": "REVIEW",
                    "value": len(native_blank_pages),
                    "message": (
                        f"{len(native_blank_pages)} page(s) have little/no native text "
                        "and Tesseract is unavailable; scanned content could not be OCR-extracted."
                    ),
                }])
            ], ignore_index=True)

    raw = pd.DataFrame([{
        "source_file": filename,
        "company_name": company,
        "page_count": extraction["page_count"],
        "ocr_pages": ", ".join(map(str, extraction["ocr_pages"])),
        "ocr_confidence_mean": (
            float(np.nanmean([
                x.get("confidence", np.nan)
                for x in extraction.get("ocr_confidences", [])
            ]))
            if extraction.get("ocr_confidences") else np.nan
        ),
        "table_extraction_pages": ", ".join(
            str(i + 1) for i, tabs in enumerate(page_tables) if tabs
        ),
        "raw_text": text,
    }])

    bundle = {
        "company": company_record,
        "industrial": industrial,
        "shareholding": shareholding,
        "equity": equity,
        "profits": profits,
        "asset_profits": asset_profits,
        "non_equity_liabilities": non_equity_liabilities,
        "non_equity_assets": non_equity_assets,
        "services": services,
        "comments": comments,
        "qa": qa,
        "raw": raw,
        "extraction": extraction,
    }
    bundle["data_quality"] = build_data_quality_summary(bundle, company)
    return bundle



def process_uploaded_pdf(uploaded_file, ocr_if_needed=True):
    extraction = extract_pdf(uploaded_file, ocr_if_needed=ocr_if_needed)
    return process_extraction_result(extraction)


# -----------------------------
# Consolidation
# -----------------------------
def consolidate(results):
    def concat(key):
        frames = [
            r[key] for r in results
            if isinstance(r.get(key), pd.DataFrame) and not r[key].empty
        ]
        return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

    company = pd.DataFrame([r["company"] for r in results])

    industrial = concat("industrial")
    shareholding = concat("shareholding")
    equity = concat("equity")
    profits = concat("profits")
    asset_profits = concat("asset_profits")
    non_equity_liabilities = concat("non_equity_liabilities")
    non_equity_assets = concat("non_equity_assets")
    services = concat("services")
    comments = concat("comments")
    qa = concat("qa")
    data_quality = concat("data_quality")
    raw = concat("raw")

    # A clean company-level financial summary. This deliberately contains
    # only values parsed from the relevant questionnaire sections.
    financial_rows = []
    for r in results:
        company = r["company"].get("company_name", "")
        row = {"company_name": company}
        p = r["profits"]
        if isinstance(p, pd.DataFrame) and not p.empty:
            pr = p.iloc[0]
            for c in [
                "net_profit_after_tax", "dividends_declared",
                "dividends_paid_or_remitted", "retained_earnings_reported",
                "retained_earnings_calculated", "retained_earnings_difference"
            ]:
                if c in p.columns:
                    row[c] = pr[c]
                    if c + "_raw" in p.columns:
                        row[c + "_raw"] = pr[c + "_raw"]
                    if c + "_status" in p.columns:
                        row[c + "_status"] = pr[c + "_status"]

        ap = r["asset_profits"]
        if isinstance(ap, pd.DataFrame) and not ap.empty:
            ar = ap.iloc[0]
            for c in [
                "net_profit_after_tax", "dividends_declared",
                "dividends_paid_or_remitted", "retained_earnings_reported"
            ]:
                if c in ap.columns:
                    row["foreign_asset_" + c] = ar[c]
                    if c + "_raw" in ap.columns:
                        row["foreign_asset_" + c + "_raw"] = ar[c + "_raw"]
                    if c + "_status" in ap.columns:
                        row["foreign_asset_" + c + "_status"] = ar[c + "_status"]
        financial_rows.append(row)

    financials = pd.DataFrame(financial_rows)

    base_tables = {
        "company_master": company,
        "industrial": industrial,
        "shareholding": shareholding,
        "equity": equity,
        "profits": profits,
        "asset_profits": asset_profits,
        "non_equity_liabilities": non_equity_liabilities,
        "non_equity_assets": non_equity_assets,
        "services": services,
        "comments": comments,
        "financials": financials,
        "qa": qa,
        "data_quality": data_quality,
        "raw_text": raw,
    }
    bop_detail, bop_summary, iip_recon, bop_coverage = build_bop_iip_tables(base_tables)
    bpm6_catalog = build_bpm6_item_catalog()
    base_tables.update({
        "bop_iip_detail": bop_detail,
        "bop_iip_summary": bop_summary,
        "iip_reconciliation": iip_recon,
        "bop_coverage": bop_coverage,
        "bpm6_item_catalog": bpm6_catalog,
    })
    return base_tables




# ============================================================
# ACCOUNTING RULE VALIDATION ENGINE
# ============================================================

ACCOUNTING_TOLERANCE = 0.01


def _validation_numeric(row, aliases):
    """Return the first usable numeric value and the source column alias used."""
    if not isinstance(row, pd.Series):
        return np.nan, ""
    for col in aliases:
        if col in row.index:
            value = _num(row.get(col))
            if pd.notna(value):
                return value, col
    return np.nan, ""


def _validation_company_names(tables):
    names = set()
    for key in ["company_master", "profits", "equity", "non_equity_liabilities", "non_equity_assets"]:
        df = tables.get(key, pd.DataFrame())
        if isinstance(df, pd.DataFrame) and "company_name" in df.columns:
            vals = df["company_name"].astype("string").str.strip().replace({"": pd.NA})
            names.update(str(x) for x in vals.dropna().tolist())
    return sorted(names)


def _accounting_add_check(rows, company, category, rule_id, check, status,
                          reported=np.nan, expected=np.nan, difference=np.nan,
                          source="", objection="", recommendation="", detail=""):
    rows.append({
        "company_name": company,
        "category": category,
        "rule_id": rule_id,
        "check": check,
        "status": status,
        "reported_value": reported,
        "expected_value": expected,
        "difference": difference,
        "source": source,
        "objection": objection,
        "recommendation": recommendation,
        "detail": detail,
        "checked_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })


def _accounting_status_from_diff(diff, tolerance=ACCOUNTING_TOLERANCE):
    if pd.isna(diff):
        return "NOT TESTABLE"
    return "PASS" if abs(float(diff)) <= tolerance else "OBJECTION"


def validate_accounting_rules(tables, tolerance=ACCOUNTING_TOLERANCE):
    """
    Validate accounting identities using only fields actually available in the
    reviewed FPC questionnaire tables. Missing source fields are reported as
    NOT TESTABLE rather than being assumed to be zero.

    Core checks include:
      1. Profit after tax less dividends = retained earnings.
      2. Profit before tax less tax = profit after tax when those fields exist.
      3. Equity roll-forward for every reported equity/reserve category:
         closing = opening + increases - decreases + price + FX + volume/other.
      4. Accumulated retained earnings roll-forward against profit/dividends.
      5. Non-equity liability roll-forward using the same balance identity.
      6. Non-equity asset roll-forward using the same balance identity.
      7. Shareholding totals are screened for impossible totals (>100%).

    The function never fabricates missing accounting values. It also retains
    explicit source labels so an objection can be traced back to a questionnaire
    response and revised by the analyst.
    """
    rows = []
    tolerance = float(tolerance or ACCOUNTING_TOLERANCE)
    companies = _validation_company_names(tables)

    profits = tables.get("profits", pd.DataFrame())
    equity = tables.get("equity", pd.DataFrame())
    ne_liab = tables.get("non_equity_liabilities", pd.DataFrame())
    ne_asset = tables.get("non_equity_assets", pd.DataFrame())
    shareholding = tables.get("shareholding", pd.DataFrame())
    financials = tables.get("financials", pd.DataFrame())

    # ---------- Profit and retained earnings ----------
    for company in companies:
        p = profits[profits["company_name"].astype("string").str.strip() == company].copy() if (
            isinstance(profits, pd.DataFrame) and not profits.empty and "company_name" in profits.columns
        ) else pd.DataFrame()
        if not p.empty:
            pr = p.iloc[0]
            npat, npat_col = _validation_numeric(pr, ["net_profit_after_tax", "profit_after_tax", "net_income_after_tax"])
            div, div_col = _validation_numeric(pr, ["dividends_declared", "dividends_paid_or_remitted", "dividends_paid"])
            retained, retained_col = _validation_numeric(pr, ["retained_earnings_reported", "retained_earnings", "retained_profit"])

            # Primary retained-profit identity.
            if pd.notna(npat) and pd.notna(div) and pd.notna(retained):
                expected = npat - div
                diff = retained - expected
                status = _accounting_status_from_diff(diff, tolerance)
                _accounting_add_check(
                    rows, company, "Profit & Retained Earnings", "P001",
                    "Net profit after tax - dividends declared = retained earnings",
                    status, retained, expected, diff,
                    f"profits.{npat_col}; profits.{div_col}; profits.{retained_col}",
                    "Reported retained earnings do not reconcile with net profit after tax less dividends." if status == "OBJECTION" else "",
                    "Revise the affected questionnaire response and rerun Accounting Validation." if status == "OBJECTION" else "",
                    "Primary retained-profit reconciliation."
                )
            else:
                _accounting_add_check(
                    rows, company, "Profit & Retained Earnings", "P001",
                    "Net profit after tax - dividends declared = retained earnings",
                    "NOT TESTABLE", retained, npat - div if pd.notna(npat) and pd.notna(div) else np.nan, np.nan,
                    "profits", "", "", "Required source fields are not all available in the questionnaire extraction."
                )

            # Optional profit-before-tax/tax check. These aliases make the validator
            # future-proof if the questionnaire parser is expanded to capture tax.
            pbt, pbt_col = _validation_numeric(pr, ["profit_before_tax", "profit_before_tax_income", "profit_before_tax_loss", "profit_before_tax_reported", "profit_before_tax_or_loss"])
            tax, tax_col = _validation_numeric(pr, ["income_tax", "tax_on_income", "income_tax_expense", "tax_expense", "tax_payable"])
            if pd.notna(pbt) and pd.notna(tax) and pd.notna(npat):
                expected = pbt - tax
                diff = npat - expected
                status = _accounting_status_from_diff(diff, tolerance)
                _accounting_add_check(
                    rows, company, "Profit & Tax", "P002",
                    "Profit before tax - income tax = net profit after tax",
                    status, npat, expected, diff,
                    f"profits.{pbt_col}; profits.{tax_col}; profits.{npat_col}",
                    "Net profit after tax does not reconcile with profit before tax less tax." if status == "OBJECTION" else "",
                    "Revise the profit/tax response and rerun Accounting Validation." if status == "OBJECTION" else "",
                    "Applied only because all three source values were available."
                )
            else:
                _accounting_add_check(
                    rows, company, "Profit & Tax", "P002",
                    "Profit before tax - income tax = net profit after tax",
                    "NOT TESTABLE", npat, np.nan, np.nan,
                    "profits", "", "", "Tax/profit-before-tax fields are not currently captured by the questionnaire parser; no assumption was made."
                )

        else:
            _accounting_add_check(rows, company, "Profit & Retained Earnings", "P000", "Profit and retained earnings section available", "NOT TESTABLE", source="profits", detail="No profit record is available for this company.")

    # The foreign-asset profit schedule uses the same internal accounting
    # identity where the questionnaire supplies the four corresponding fields.
    asset_profits = tables.get("asset_profits", pd.DataFrame())
    if isinstance(asset_profits, pd.DataFrame) and not asset_profits.empty and "company_name" in asset_profits.columns:
        for company in _validation_company_names({"asset_profits": asset_profits}):
            ap = asset_profits[asset_profits["company_name"].astype("string").str.strip() == company].copy()
            if ap.empty:
                continue
            ar = ap.iloc[0]
            npat, npat_col = _validation_numeric(ar, ["net_profit_after_tax", "profit_after_tax", "net_income_after_tax"])
            div, div_col = _validation_numeric(ar, ["dividends_declared", "dividends_paid_or_remitted", "dividends_paid"])
            retained, retained_col = _validation_numeric(ar, ["retained_earnings_reported", "retained_earnings", "retained_profit"])
            if pd.notna(npat) and pd.notna(div) and pd.notna(retained):
                expected = npat - div
                diff = retained - expected
                status = _accounting_status_from_diff(diff, tolerance)
                _accounting_add_check(
                    rows, company, "Foreign Asset Profits", "FA001",
                    "Foreign asset net profit after tax - dividends declared = retained earnings",
                    status, retained, expected, diff,
                    f"asset_profits.{npat_col}; asset_profits.{div_col}; asset_profits.{retained_col}",
                    "Foreign asset profit/retained earnings do not reconcile." if status == "OBJECTION" else "",
                    "Revise the affected foreign-asset questionnaire response and rerun Accounting Validation." if status == "OBJECTION" else "",
                    "Applied only where all four source cells are available."
                )
            else:
                _accounting_add_check(
                    rows, company, "Foreign Asset Profits", "FA001",
                    "Foreign asset net profit after tax - dividends declared = retained earnings",
                    "NOT TESTABLE", retained, npat - div if pd.notna(npat) and pd.notna(div) else np.nan, np.nan,
                    "asset_profits", "", "", "The foreign-asset profit schedule does not contain all fields required for this accounting identity."
                )

    # ---------- Equity and reserves roll-forward ----------
    if isinstance(equity, pd.DataFrame) and not equity.empty and "company_name" in equity.columns:
        for _, r in equity.iterrows():
            company = _text(r.get("company_name"))
            label = _text(r.get("equity_type")) or "Equity / Reserve"
            a = _num(r.get("closing_2024"))
            b = _num(r.get("purchase_increase_2025"))
            c = _num(r.get("sales_decrease_2025"))
            d1 = _num(r.get("other_price"))
            d2 = _num(r.get("other_exchange_rate"))
            d3 = _num(r.get("other_volume"))
            e = _num(r.get("closing_2025"))
            required = [a, b, c, d1, d2, d3, e]
            # D1-D3 can legitimately be blank in the source. If the closing/opening
            # and transaction fields are sufficient, blank other-change components
            # are treated as zero but explicitly disclosed in the detail.
            if pd.notna(a) and pd.notna(e):
                # Retained earnings is governed by the profit/dividend roll-forward,
                # not by share-capital purchases/sales. Other equity/reserve classes
                # use the questionnaire A/B/C/D1-D3/E balance identity.
                if "retained earnings" in label.lower():
                    p = profits[profits["company_name"].astype("string").str.strip() == company].copy() if (
                        isinstance(profits, pd.DataFrame) and not profits.empty and "company_name" in profits.columns
                    ) else pd.DataFrame()
                    npat = div = np.nan
                    if not p.empty:
                        pr = p.iloc[0]
                        npat, _ = _validation_numeric(pr, ["net_profit_after_tax", "profit_after_tax", "net_income_after_tax"])
                        div, _ = _validation_numeric(pr, ["dividends_declared", "dividends_declared_current_year"])
                    if pd.notna(npat) and pd.notna(div):
                        d1v = 0.0 if pd.isna(d1) else d1
                        d2v = 0.0 if pd.isna(d2) else d2
                        d3v = 0.0 if pd.isna(d3) else d3
                        expected = a + npat - div + d1v + d2v + d3v
                        diff = e - expected
                        status = _accounting_status_from_diff(diff, tolerance)
                        note = "Retained earnings roll-forward uses opening retained earnings + net profit after tax - dividends; D1-D3 are included only when explicitly reported."
                        _accounting_add_check(
                            rows, company, "Equity & Reserves", "E001",
                            f"{label}: closing balance roll-forward",
                            status, e, expected, diff,
                            "equity.closing_2024 + profits.net_profit_after_tax - profits.dividends_declared + equity.D1-D3",
                            f"{label} closing balance does not reconcile with opening retained earnings plus retained profit less dividends." if status == "OBJECTION" else "",
                            "Revise the retained earnings, profit or dividend response and rerun Accounting Validation." if status == "OBJECTION" else "",
                            note
                        )
                    else:
                        _accounting_add_check(
                            rows, company, "Equity & Reserves", "E001",
                            f"{label}: closing balance roll-forward",
                            "NOT TESTABLE", e, np.nan, np.nan,
                            "equity + profits", "", "",
                            "Retained earnings requires net profit after tax and dividends to validate the closing balance. No missing value was assumed."
                        )
                elif pd.notna(b) and pd.notna(c):
                    d1v = 0.0 if pd.isna(d1) else d1
                    d2v = 0.0 if pd.isna(d2) else d2
                    d3v = 0.0 if pd.isna(d3) else d3
                    expected = a + b - c + d1v + d2v + d3v
                    diff = e - expected
                    status = _accounting_status_from_diff(diff, tolerance)
                    blanks = [x for x, v in [("price", d1), ("exchange-rate", d2), ("volume", d3)] if pd.isna(v)]
                    note = "" if not blanks else f"Blank other-change component(s) treated as zero for reconciliation: {', '.join(blanks)}."
                    _accounting_add_check(
                        rows, company, "Equity & Reserves", "E001",
                        f"{label}: closing balance roll-forward",
                        status, e, expected, diff,
                        "equity.A/B/C/D1-D3/E",
                        f"{label} closing balance does not reconcile with opening balance plus increases less decreases and other changes." if status == "OBJECTION" else "",
                        "Revise the affected equity/reserve row and rerun Accounting Validation." if status == "OBJECTION" else "",
                        note or "Explicit A/B/C/D1-D3/E roll-forward."
                    )
                else:
                    _accounting_add_check(
                        rows, company, "Equity & Reserves", "E001",
                        f"{label}: closing balance roll-forward",
                        "NOT TESTABLE", e, np.nan, np.nan,
                        "equity.A/B/C/D1-D3/E", "", "",
                        "Opening balance, transaction/change fields and closing balance are not all available; no value was assumed."
                    )
            else:
                _accounting_add_check(
                    rows, company, "Equity & Reserves", "E001",
                    f"{label}: closing balance roll-forward",
                    "NOT TESTABLE", e, np.nan, np.nan,
                    "equity.A/B/C/D1-D3/E", "", "", "Opening balance, transaction/change fields and closing balance are not all available; no value was assumed."
                )

        # Retained earnings equity row must also agree with the profit/dividend bridge.
        for _, er in equity[equity["equity_type"].astype("string").str.contains("retained earnings", case=False, na=False)].iterrows():
            company = _text(er.get("company_name"))
            opening = _num(er.get("closing_2024"))
            closing = _num(er.get("closing_2025"))
            p = profits[profits["company_name"].astype("string").str.strip() == company].copy() if (
                isinstance(profits, pd.DataFrame) and not profits.empty and "company_name" in profits.columns
            ) else pd.DataFrame()
            npat = div = np.nan
            if not p.empty:
                pr = p.iloc[0]
                npat, _ = _validation_numeric(pr, ["net_profit_after_tax", "profit_after_tax", "net_income_after_tax"])
                div, _ = _validation_numeric(pr, ["dividends_declared", "dividends_paid_or_remitted", "dividends_paid"])
            if pd.notna(opening) and pd.notna(closing) and pd.notna(npat) and pd.notna(div):
                expected_change = npat - div
                reported_change = closing - opening
                diff = reported_change - expected_change
                status = _accounting_status_from_diff(diff, tolerance)
                _accounting_add_check(
                    rows, company, "Equity & Reserves", "E002",
                    "Accumulated retained earnings change = net profit after tax - dividends",
                    status, reported_change, expected_change, diff,
                    "equity.retained_earnings + profits.net_profit_after_tax + profits.dividends_declared",
                    "The retained-earnings movement in the equity schedule does not agree with profit less dividends." if status == "OBJECTION" else "",
                    "Revise the retained earnings, profit or dividend response and rerun Accounting Validation." if status == "OBJECTION" else "",
                    "Cross-statement retained earnings reconciliation."
                )
            else:
                _accounting_add_check(rows, company, "Equity & Reserves", "E002", "Accumulated retained earnings change = net profit after tax - dividends", "NOT TESTABLE", source="equity + profits", detail="The retained-earnings roll-forward cannot be fully tested because one or more required values are unavailable.")

    # ---------- Non-equity liabilities/assets ----------
    def validate_balance_table(df, category, prefix):
        if not isinstance(df, pd.DataFrame) or df.empty:
            for company in companies:
                _accounting_add_check(rows, company, category, prefix + "000", f"{category} balance schedule available", "NOT TESTABLE", source=category, detail="No records are available for this section.")
            return
        for _, r in df.iterrows():
            company = _text(r.get("company_name"))
            instrument = _text(r.get("instrument")) or "Financial instrument"
            a = _num(r.get("closing_balance_2024"))
            b = _num(r.get("amount_received_2025"))
            c = _num(r.get("principal_repayment_2025"))
            d1 = _num(r.get("other_price"))
            d2 = _num(r.get("other_exchange_rate"))
            d3 = _num(r.get("other_volume"))
            e = _num(r.get("closing_balance_2025"))
            if pd.notna(a) and pd.notna(b) and pd.notna(c) and pd.notna(e):
                d1v = 0.0 if pd.isna(d1) else d1
                d2v = 0.0 if pd.isna(d2) else d2
                d3v = 0.0 if pd.isna(d3) else d3
                expected = a + b - c + d1v + d2v + d3v
                diff = e - expected
                status = _accounting_status_from_diff(diff, tolerance)
                blanks = [x for x, v in [("price", d1), ("exchange-rate", d2), ("volume", d3)] if pd.isna(v)]
                note = "" if not blanks else f"Blank other-change component(s) treated as zero: {', '.join(blanks)}."
                _accounting_add_check(
                    rows, company, category, prefix + "001",
                    f"{instrument}: closing balance roll-forward",
                    status, e, expected, diff,
                    f"{prefix}.A/B/C/D1-D3/E",
                    f"{instrument} closing balance does not reconcile with opening balance plus receipts less principal repayments and other changes." if status == "OBJECTION" else "",
                    "Revise the affected financial-instrument row and rerun Accounting Validation." if status == "OBJECTION" else "",
                    note or "Explicit A/B/C/D1-D3/E roll-forward."
                )
            else:
                _accounting_add_check(rows, company, category, prefix + "001", f"{instrument}: closing balance roll-forward", "NOT TESTABLE", source=prefix, detail="Opening balance, transaction/change fields and closing balance are not all available; no value was assumed.")

    validate_balance_table(ne_liab, "Non-equity Liabilities", "NEL")
    validate_balance_table(ne_asset, "Non-equity Assets", "NEA")

    # ---------- Shareholding impossibility screen ----------
    if isinstance(shareholding, pd.DataFrame) and not shareholding.empty and "company_name" in shareholding.columns:
        for company, g in shareholding.groupby("company_name", dropna=False):
            company = _text(company)
            cur = pd.to_numeric(g.get("percentage_shareholding_current"), errors="coerce") if "percentage_shareholding_current" in g.columns else pd.Series(dtype=float)
            prev = pd.to_numeric(g.get("percentage_shareholding_previous"), errors="coerce") if "percentage_shareholding_previous" in g.columns else pd.Series(dtype=float)
            cur_total = cur.sum(min_count=1)
            prev_total = prev.sum(min_count=1)
            if pd.notna(cur_total) and cur_total > 100 + tolerance:
                _accounting_add_check(rows, company, "Shareholding", "S001", "Current shareholding total does not exceed 100%", "OBJECTION", cur_total, 100.0, cur_total - 100.0, "shareholding.percentage_shareholding_current", "Current shareholding percentages exceed 100%.", "Revise shareholder percentages in the questionnaire and rerun Accounting Validation.", "Impossible total detected.")
            else:
                _accounting_add_check(rows, company, "Shareholding", "S001", "Current shareholding total does not exceed 100%", "PASS", cur_total, 100.0, 0.0 if pd.notna(cur_total) else np.nan, "shareholding.percentage_shareholding_current", detail="A total below 100% is not treated as an error because the questionnaire may not capture every holder in the available response.")
            if pd.notna(prev_total) and prev_total > 100 + tolerance:
                _accounting_add_check(rows, company, "Shareholding", "S002", "Previous-period shareholding total does not exceed 100%", "OBJECTION", prev_total, 100.0, prev_total - 100.0, "shareholding.percentage_shareholding_previous", "Previous-period shareholding percentages exceed 100%.", "Revise shareholder percentages in the questionnaire and rerun Accounting Validation.", "Impossible total detected.")
            else:
                _accounting_add_check(rows, company, "Shareholding", "S002", "Previous-period shareholding total does not exceed 100%", "PASS", prev_total, 100.0, 0.0 if pd.notna(prev_total) else np.nan, "shareholding.percentage_shareholding_previous", detail="A total below 100% is not treated as an error.")

    # ---------- Optional total-equity / total-assets / total-liabilities checks ----------
    # The current parser does not populate balance-sheet totals. If future versions
    # add them, the aliases below automatically activate the checks without changing
    # the core validation architecture.
    total_aliases = {
        "total_assets": ["total_assets", "assets_total", "total_asset"],
        "total_liabilities": ["total_liabilities", "liabilities_total", "total_liability"],
        "total_equity": ["total_equity", "equity_total", "shareholders_equity"],
    }
    if isinstance(financials, pd.DataFrame) and not financials.empty and "company_name" in financials.columns:
        for _, r in financials.iterrows():
            company = _text(r.get("company_name"))
            ta, ta_col = _validation_numeric(r, total_aliases["total_assets"])
            tl, tl_col = _validation_numeric(r, total_aliases["total_liabilities"])
            te, te_col = _validation_numeric(r, total_aliases["total_equity"])
            if pd.notna(ta) and pd.notna(tl) and pd.notna(te):
                diff = ta - (tl + te)
                status = _accounting_status_from_diff(diff, tolerance)
                _accounting_add_check(rows, company, "Balance Sheet", "BS001", "Total assets = total liabilities + total equity", status, ta, tl + te, diff, f"financials.{ta_col}; financials.{tl_col}; financials.{te_col}", "The balance-sheet accounting equation does not reconcile." if status == "OBJECTION" else "", "Revise the balance-sheet response and rerun Accounting Validation." if status == "OBJECTION" else "", "Activated because all three balance-sheet totals are present.")
            else:
                _accounting_add_check(rows, company, "Balance Sheet", "BS001", "Total assets = total liabilities + total equity", "NOT TESTABLE", source="financials", detail="The current questionnaire extraction does not contain all three balance-sheet totals.")

    result = pd.DataFrame(rows)
    if result.empty:
        result = pd.DataFrame(columns=[
            "company_name", "category", "rule_id", "check", "status", "reported_value",
            "expected_value", "difference", "source", "objection", "recommendation",
            "detail", "checked_at"
        ])

    objection_count = int((result["status"] == "OBJECTION").sum()) if "status" in result.columns else 0
    pass_count = int((result["status"] == "PASS").sum()) if "status" in result.columns else 0
    not_testable_count = int((result["status"] == "NOT TESTABLE").sum()) if "status" in result.columns else 0
    status = "ACCOUNTING RULES VALIDATED" if objection_count == 0 else "OBJECTIONS RAISED"
    summary = {
        "status": status,
        "objections": objection_count,
        "passes": pass_count,
        "not_testable": not_testable_count,
        "total_checks": len(result),
        "validated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    return result, summary


def accounting_validation_excel_bytes(validation_df):
    return table_excel_bytes(validation_df, "Accounting Validation")


# ============================================================
# DATA CLEANING ENGINE
# ============================================================

MISSING_STRINGS = {
    "", "na", "n/a", "n.a", "n.a.", "nil", "nill", "none", "null",
    "not available", "not applicable", "n//a", "n?a", "missing"
}


def safe_numeric_series(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    text = s.astype("string").str.strip()
    text = text.str.replace(",", "", regex=False)
    text = text.str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    text = text.str.replace(r"^(US\$|USS\$|USD|US|\$)", "", regex=True)
    return pd.to_numeric(text, errors="coerce")


def clean_dataframe(df, key, trim=True, standardise_missing=True,
                    numeric_conversion=True, remove_duplicates=True):
    if not isinstance(df, pd.DataFrame):
        return pd.DataFrame(), {
            "dataset": key, "rows_before": 0, "rows_after": 0,
            "duplicates_removed": 0, "missing_standardised": 0,
            "numeric_columns_converted": 0
        }

    out = df.copy()
    before = len(out)
    missing_changed = 0
    numeric_changed = 0

    # Standardise column names without losing meaning.
    out.columns = [
        re.sub(r"\s+", "_", str(c).strip().lower())
        for c in out.columns
    ]

    for col in out.columns:
        if out[col].dtype == "object" or pd.api.types.is_string_dtype(out[col]):
            if trim:
                original = out[col].astype("string")
                out[col] = original.str.replace(r"\s+", " ", regex=True).str.strip()

            if standardise_missing:
                s = out[col].astype("string")
                mask = s.str.lower().isin(MISSING_STRINGS)
                missing_changed += int(mask.sum())
                out.loc[mask, col] = pd.NA

    if numeric_conversion:
        for col in out.columns:
            name = str(col).lower()
            # Do not convert raw/status/source text.
            if any(x in name for x in [
                "_raw", "status", "message", "comment", "text", "file",
                "address", "street", "district", "region", "country",
                "instrument", "relationship", "activity", "sector",
                "respondent", "position", "date"
            ]):
                continue

            converted = safe_numeric_series(out[col])
            non_missing_original = out[col].notna().sum()
            if non_missing_original > 0:
                numeric_share = converted.notna().sum() / non_missing_original
                if numeric_share >= 0.85:
                    out[col] = converted
                    numeric_changed += 1

    duplicates_removed = 0
    if remove_duplicates and not out.empty:
        before_dedup = len(out)
        out = out.drop_duplicates().reset_index(drop=True)
        duplicates_removed = before_dedup - len(out)

    log = {
        "dataset": key,
        "rows_before": before,
        "rows_after": len(out),
        "duplicates_removed": duplicates_removed,
        "missing_standardised": missing_changed,
        "numeric_columns_converted": numeric_changed,
    }
    return out, log


def clean_all_tables(options=None):
    options = options or {}
    cleaned = {}
    logs = []
    derived_keys = {"bop_iip_detail", "bop_iip_summary", "iip_reconciliation", "bop_coverage", "bpm6_item_catalog", "accounting_validation"}

    # Clean source tables first. Derived BOP/IIP tables are then rebuilt from the
    # cleaned source tables so that the bridge always reflects the current data.
    for key in TABLE_KEYS:
        if key in derived_keys:
            continue
        df = st.session_state.get(key, pd.DataFrame())
        cdf, log = clean_dataframe(
            df, key,
            trim=options.get("trim", True),
            standardise_missing=options.get("missing", True),
            numeric_conversion=options.get("numeric", True),
            remove_duplicates=options.get("duplicates", True),
        )
        cleaned[key] = cdf
        logs.append(log)

    bd, bs, ir, bc = build_bop_iip_tables(cleaned)
    bpm6_catalog = build_bpm6_item_catalog()
    cleaned.update({
        "bop_iip_detail": bd,
        "bop_iip_summary": bs,
        "iip_reconciliation": ir,
        "bop_coverage": bc,
        "bpm6_item_catalog": bpm6_catalog,
    })
    for key, df in [("bop_iip_detail", bd), ("bop_iip_summary", bs), ("iip_reconciliation", ir), ("bop_coverage", bc), ("bpm6_item_catalog", bpm6_catalog)]:
        logs.append({"dataset": key, "rows_before": len(df), "rows_after": len(df), "duplicates_removed": 0, "missing_standardised": 0, "numeric_columns_converted": 0})

    st.session_state["cleaned_tables"] = cleaned
    st.session_state["cleaned"] = True
    st.session_state["last_cleaned"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state["cleaning_log"] = pd.DataFrame(logs)
    save_persisted_state()
    return cleaned


def active_table(key):
    if st.session_state.get("cleaned") and key in st.session_state.get("cleaned_tables", {}):
        df = st.session_state["cleaned_tables"].get(key)
        if isinstance(df, pd.DataFrame):
            return df
    return st.session_state.get(key, pd.DataFrame())


def active_tables():
    return {k: active_table(k) for k in TABLE_KEYS}


def style_worksheet(ws, header_row=5):
    """Apply a simple, robust style to an exported worksheet."""
    for cell in ws[1]:
        cell.font = Font(bold=True, size=14)
    for cell in ws[header_row]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            try:
                max_len = max(max_len, len(str(cell.value)) if cell.value is not None else 0)
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 55)
    ws.sheet_view.showGridLines = False


def _excel_safe_value(value):
    """Convert pandas/numpy values into values accepted by openpyxl."""
    if value is None:
        return None
    try:
        if value is pd.NA or value is pd.NaT:
            return None
    except Exception:
        pass
    if isinstance(value, (np.generic,)):
        value = value.item()
    if isinstance(value, (float, np.floating)):
        try:
            if not np.isfinite(float(value)):
                return None
        except Exception:
            return None
        return float(value)
    if isinstance(value, (int, np.integer, bool, str)):
        return value
    if isinstance(value, (pd.Timestamp, datetime)):
        return value.to_pydatetime() if isinstance(value, pd.Timestamp) else value
    if isinstance(value, (list, tuple, set, dict)):
        return str(value)
    try:
        missing = pd.isna(value)
        if isinstance(missing, (bool, np.bool_)) and bool(missing):
            return None
    except Exception:
        pass
    return value


def table_excel_bytes(df, sheet_name="Data"):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = safe_filename(str(sheet_name))[:31] or "Data"
    if df is None or df.empty:
        ws.append(["No records"])
    else:
        ws.append([str(c) for c in df.columns])
        for row in df.itertuples(index=False, name=None):
            ws.append([_excel_safe_value(x) for x in row])
    style_worksheet(ws, header_row=1)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions if df is not None and not df.empty else None
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def build_excel(tables, workbook_title="FPC Survey Intelligence & Population Estimation"):
    """Build a multi-sheet Excel workbook with strict openpyxl-safe values."""
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)

    preferred = [
        "company_master", "financials", "industrial", "shareholding",
        "equity", "profits", "asset_profits", "non_equity_liabilities",
        "non_equity_assets", "services", "comments", "qa", "data_quality", "accounting_validation",
        "bop_iip_summary", "bop_iip_detail", "iip_reconciliation", "bop_coverage", "bpm6_item_catalog", "raw_text"
    ]
    keys = [k for k in preferred if k in tables]
    keys.extend([k for k in tables if k not in keys])
    used_titles = set()

    for key in keys:
        df = tables.get(key)
        if not isinstance(df, pd.DataFrame):
            df = pd.DataFrame()
        base = safe_filename(str(key).replace("_", " "))[:31] or "Data"
        title = base
        n = 2
        while title in used_titles:
            suffix = f"_{n}"
            title = (base[:31-len(suffix)] + suffix)[:31]
            n += 1
        used_titles.add(title)
        ws = wb.create_sheet(title=title)
        ws.append([workbook_title])
        ws.append([f"Stage: {workbook_title}"])
        ws.append([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
        ws.append([])
        if df.empty:
            ws.append(["No records available"])
        else:
            ws.append([str(c) for c in df.columns])
            for row in df.itertuples(index=False, name=None):
                ws.append([_excel_safe_value(v) for v in row])
        style_worksheet(ws)
        ws.freeze_panes = "A6"
        if not df.empty:
            ws.auto_filter.ref = ws.dimensions

    readme = wb.create_sheet(title="README", index=0)
    readme.append([workbook_title])
    readme.append(["FPC Survey Intelligence stage export"])
    readme.append([f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"])
    readme.append([])
    readme.append(["Sheet", "Rows", "Columns"])
    for key in keys:
        df = tables.get(key)
        if not isinstance(df, pd.DataFrame):
            df = pd.DataFrame()
        readme.append([key, int(len(df)), int(len(df.columns))])
    style_worksheet(readme)
    readme.freeze_panes = "A6"
    wb.save(output)
    output.seek(0)
    return output.getvalue()

def full_stage_excel_bytes(tables, stage_name):
    return build_excel(tables, workbook_title=f"FPC Survey Intelligence & Population Estimation - {stage_name}")



# ============================================================
# BPM6 BOP / IIP BRIDGE ENGINE
# ============================================================

BPM6_SOURCE_NOTE = (
    "BPM6 functional categories: direct investment, portfolio investment, "
    "financial derivatives/employee stock options, other investment, and reserve assets. "
    "Financial-account flows are presented as net acquisition of financial assets and "
    "net incurrence of liabilities; increases in assets/liabilities are positive and "
    "decreases are negative."
)


def _num(v):
    try:
        x = pd.to_numeric(pd.Series([v]), errors="coerce").iloc[0]
        return float(x) if pd.notna(x) else np.nan
    except Exception:
        return np.nan


def _text(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    return str(v).strip()


def _is_foreign_counterparty(country, relationship=""):
    c = _text(country).upper()
    r = _text(relationship).upper()
    if not c and not r:
        return None
    if "RESIDENT" in r and "NON-RESIDENT" not in r:
        return False
    if c in {"ZIMBABWE", "ZIMBABWEAN", "RESIDENT"}:
        return False
    # Explicit international organisations / foreign locations are treated as foreign.
    return True


def _relationship_is_di(relationship):
    r = _text(relationship).upper()
    return any(x in r for x in ["DI", "PARENT", "FELLOW", "AFFILIATED"])


def _ownership_classification(shareholding):
    """Build BPM6-oriented ownership indicators from the survey shareholding table.

    The survey sometimes explicitly labels relationships as DI even where an individual
    percentage is below 10 percent. We retain that information but also calculate the
    strict BPM6 10-percent test so the compiler can review differences rather than
    silently overriding the source questionnaire.
    """
    if not isinstance(shareholding, pd.DataFrame) or shareholding.empty:
        return pd.DataFrame(columns=["company_name"])

    d = shareholding.copy()
    for c in ["percentage_shareholding_current", "percentage_shareholding_previous"]:
        if c in d.columns:
            d[c] = pd.to_numeric(d[c], errors="coerce")
    if "relationship_current" not in d.columns:
        d["relationship_current"] = ""
    if "source_country_or_organisation" not in d.columns:
        d["source_country_or_organisation"] = ""

    rows = []
    for company, g in d.groupby("company_name", dropna=False):
        foreign_di_reported = 0.0
        foreign_portfolio = 0.0
        foreign_total = 0.0
        strict_di = 0.0
        ambiguous = 0.0
        di_rows = 0
        portfolio_rows = 0

        for _, r in g.iterrows():
            pct = _num(r.get("percentage_shareholding_current"))
            if pd.isna(pct):
                continue
            foreign = _is_foreign_counterparty(
                r.get("source_country_or_organisation"),
                r.get("relationship_current")
            )
            rel = _text(r.get("relationship_current"))
            if foreign is not True:
                continue
            foreign_total += pct
            rel_di = _relationship_is_di(rel)
            # Strict BPM6 criterion is investor-level 10%+ voting power.
            if pct >= 10:
                strict_di += pct
            if rel_di:
                foreign_di_reported += pct
                di_rows += 1
                # Core BPM6 bridge: an individual foreign holding below 10%
                # does not qualify as DI merely because the survey labels it DI.
                # Keep it as provisional portfolio equity and flag the mismatch.
                if pct < 10:
                    foreign_portfolio += pct
                    portfolio_rows += 1
            elif pct < 10:
                foreign_portfolio += pct
                portfolio_rows += 1
            else:
                ambiguous += pct

        status = "PASS"
        note = ""
        if foreign_di_reported and abs(foreign_di_reported - strict_di) > 1e-8:
            status = "REVIEW"
            note = (
                "Survey relationship labels and the strict BPM6 10% investor-level test "
                "do not produce the same DI share. Review investor/group relationships."
            )
        if foreign_total > 100.0001:
            status = "FAIL"
            note = "Foreign ownership exceeds 100%; review shareholding extraction."

        rows.append({
            "company_name": company,
            "foreign_ownership_current_pct": foreign_total if foreign_total else np.nan,
            "foreign_di_pct_survey_reported": foreign_di_reported if foreign_di_reported else 0.0,
            "foreign_di_pct_bpm6_10pct": strict_di if strict_di else 0.0,
            "foreign_portfolio_equity_pct_estimate": foreign_portfolio if foreign_portfolio else 0.0,
            "foreign_ownership_ambiguous_pct": ambiguous if ambiguous else 0.0,
            "di_shareholder_rows": di_rows,
            "portfolio_shareholder_rows": portfolio_rows,
            "ownership_classification_status": status,
            "ownership_classification_note": note,
        })
    return pd.DataFrame(rows)


def _equity_totals(equity):
    if not isinstance(equity, pd.DataFrame) or equity.empty:
        return pd.DataFrame(columns=["company_name"])
    d = equity.copy()
    num_cols = [
        "closing_2024", "purchase_increase_2025", "sales_decrease_2025",
        "other_price", "other_exchange_rate", "other_volume", "closing_2025"
    ]
    for c in num_cols:
        if c in d.columns:
            d[c] = pd.to_numeric(d[c], errors="coerce")
    rows = []
    for company, g in d.groupby("company_name", dropna=False):
        row = {"company_name": company}
        for c in num_cols:
            row[c] = g[c].sum(min_count=1) if c in g.columns else np.nan
        # Retained earnings are not treated as an ordinary equity purchase flow.
        # They are linked separately to reinvestment of earnings in the financial account.
        re = g[g["equity_type"].astype("string").str.contains(
            "retained earnings", case=False, na=False
        )]
        row["retained_earnings_change_reported"] = (
            (_num(re.iloc[0].get("closing_2025")) - _num(re.iloc[0].get("closing_2024")))
            if not re.empty and pd.notna(_num(re.iloc[0].get("closing_2025"))) and pd.notna(_num(re.iloc[0].get("closing_2024")))
            else np.nan
        )
        row["equity_transaction_gross_proxy"] = (
            row.get("purchase_increase_2025", np.nan) - row.get("sales_decrease_2025", np.nan)
            if pd.notna(row.get("purchase_increase_2025", np.nan)) or pd.notna(row.get("sales_decrease_2025", np.nan))
            else np.nan
        )
        rows.append(row)
    return pd.DataFrame(rows)


def _profit_reinvestment(profits, company):
    if not isinstance(profits, pd.DataFrame) or profits.empty or "company_name" not in profits.columns:
        return np.nan, "No profit data"
    p = profits[profits["company_name"] == company]
    if p.empty:
        return np.nan, "No profit row"
    r = p.iloc[0]
    profit = _num(r.get("net_profit_after_tax"))
    div = _num(r.get("dividends_declared"))
    if pd.notna(profit) and pd.notna(div):
        return profit - div, "Net profit after tax minus dividends declared"
    if pd.notna(profit) and _text(r.get("dividends_declared_status")).upper() in {"BLANK", "N/A", "DASH"}:
        return np.nan, "Dividends declared not reported; reinvestment cannot be derived safely"
    return np.nan, "Insufficient profit/dividend information"


def build_bpm6_item_catalog():
    rows = [
        ("Financial account", "Direct investment", "Equity and investment fund shares", "FPC", "Yes", "Foreign shareholding + enterprise equity; book-value proxy"),
        ("Financial account", "Direct investment", "Reinvestment of earnings", "FPC", "Yes", "Profit less dividends, multiplied by qualifying foreign DI share"),
        ("Financial account", "Direct investment", "Debt instruments / intercompany debt", "FPC", "Yes", "Selected foreign loans/debt/trade credit with DI/fellow/parent relationships"),
        ("Financial account", "Portfolio investment", "Equity and investment fund shares", "FPC", "Partial", "Potential foreign equity below 10%; investor-level security classification required"),
        ("Financial account", "Portfolio investment", "Debt securities", "FPC", "Partial", "Only where reported in the survey; securities-level data required for completeness"),
        ("Financial account", "Financial derivatives and ESOs", "Financial derivatives", "Other source", "No", "Not captured by current FPC survey"),
        ("Financial account", "Financial derivatives and ESOs", "Employee stock options", "Other source", "No", "Not captured by current FPC survey"),
        ("Financial account", "Other investment", "Other equity", "FPC", "Partial", "Potentially relevant; current survey coverage needs compiler review"),
        ("Financial account", "Other investment", "Currency and deposits", "Other source", "No", "Use banking/external-sector source"),
        ("Financial account", "Other investment", "Loans", "FPC", "Yes", "Foreign non-DI loans; DI loans are mapped under direct investment"),
        ("Financial account", "Other investment", "Insurance, pension and standardized guarantee schemes", "FPC", "Partial", "Selected survey categories; completeness requires other sources"),
        ("Financial account", "Other investment", "Trade credit and advances", "FPC", "Yes", "Foreign suppliers/trade credits"),
        ("Financial account", "Other investment", "Other accounts receivable/payable", "FPC", "Partial", "Foreign liabilities captured; foreign assets need improvement"),
        ("Financial account", "Other investment", "Special drawing rights", "Other source", "No", "Use IMF/institution data"),
        ("Financial account", "Reserve assets", "Monetary gold", "Other source", "No", "institution reserve-asset source"),
        ("Financial account", "Reserve assets", "Special drawing rights", "Other source", "No", "institution/IMF reserve data"),
        ("Financial account", "Reserve assets", "Reserve position in the IMF", "Other source", "No", "institution/IMF reserve data"),
        ("Financial account", "Reserve assets", "Foreign exchange assets: currency and deposits", "Other source", "No", "institution reserve assets"),
        ("Financial account", "Reserve assets", "Foreign exchange assets: debt securities", "Other source", "No", "institution reserve assets"),
        ("Financial account", "Reserve assets", "Other reserve claims", "Other source", "No", "institution reserve assets"),
        ("IIP", "Direct investment", "Direct investment abroad — assets", "Other source + FPC", "Partial", "Resident enterprises' foreign investments require external-assets coverage"),
        ("IIP", "Direct investment", "Direct investment in reporting economy — liabilities", "FPC", "Yes", "Main FPC use: foreign DI equity + selected intercompany debt"),
        ("IIP", "Portfolio investment", "Portfolio investment assets", "Other source", "No", "Resident holdings of nonresident securities required"),
        ("IIP", "Portfolio investment", "Portfolio investment liabilities", "FPC", "Partial", "Potential foreign equity/debt securities reported by surveyed enterprises"),
        ("IIP", "Other investment", "Other investment assets", "Other source", "Partial", "Foreign assets section needs stronger extraction/coverage"),
        ("IIP", "Other investment", "Other investment liabilities", "FPC", "Yes", "Loans, trade credit, other accounts payable and selected schemes"),
        ("IIP", "Reserve assets", "Reserve assets", "Other source", "No", "institution reserve-assets data"),
    ]
    return pd.DataFrame(rows, columns=["account", "functional_category", "bpm6_item", "preferred_source", "fpc_coverage", "compiler_use"])


def build_bop_iip_tables(tables):
    """Create a transparent BPM6 bridge from the FPC survey to BOP/IIP.

    Important: this is a compiler-support layer, not a replacement for the full BOP/IIP.
    The FPC survey is strongest for direct-investment liabilities and related debt of
    reporting-economy enterprises. It does not by itself provide reserve assets,
    financial derivatives, or a complete external-assets universe.
    """
    sh = tables.get("shareholding", pd.DataFrame())
    eq = tables.get("equity", pd.DataFrame())
    prof = tables.get("profits", pd.DataFrame())
    ne = tables.get("non_equity_liabilities", pd.DataFrame())
    ownership = _ownership_classification(sh)
    eqtot = _equity_totals(eq)

    detail = []
    summary_rows = []

    companies = set()
    for df in [ownership, eqtot, ne, prof]:
        if isinstance(df, pd.DataFrame) and "company_name" in df.columns:
            companies.update(df["company_name"].dropna().astype(str).str.strip())

    # --- Equity / ownership bridge ---
    for company in sorted(x for x in companies if x):
        own = ownership[ownership["company_name"] == company]
        o = own.iloc[0].to_dict() if not own.empty else {}
        di_pct = _num(o.get("foreign_di_pct_bpm6_10pct"))
        survey_di_pct = _num(o.get("foreign_di_pct_survey_reported"))
        strict_di_pct = _num(o.get("foreign_di_pct_bpm6_10pct"))
        port_pct = _num(o.get("foreign_portfolio_equity_pct_estimate"))
        # Under the core BPM6 bridge, the strict 10% investor-level criterion is used.
        # The survey-reported DI share is retained as a memo/QA field.

        et = eqtot[eqtot["company_name"] == company]
        e = et.iloc[0].to_dict() if not et.empty else {}

        # Company equity is reported for the enterprise as a whole. Allocating it by
        # foreign ownership is therefore explicitly labelled a BOOK-VALUE PROXY.
        for category, pct, note in [
            ("Direct investment", di_pct, "Strict BPM6 10%+ foreign investor share applied to enterprise equity; survey/book-value proxy."),
            ("Portfolio investment", port_pct, "Foreign equity holdings below the strict 10% DI threshold are provisionally treated as portfolio-equity; review where the questionnaire labels the investor DI."),
        ]:
            if pd.isna(pct) or pct <= 0:
                continue
            for position_side, opening_col, closing_col in [
                ("liability", "closing_2024", "closing_2025")
            ]:
                opening = _num(e.get(opening_col))
                closing = _num(e.get(closing_col))
                position_open = opening * pct / 100 if pd.notna(opening) else np.nan
                position_close = closing * pct / 100 if pd.notna(closing) else np.nan
                total_flow = _num(e.get("equity_transaction_gross_proxy"))
                flow_est = total_flow * pct / 100 if pd.notna(total_flow) else np.nan
                detail.append({
                    "company_name": company,
                    "functional_category": category,
                    "side": "Net incurrence of liabilities",
                    "instrument": "Equity and investment fund shares",
                    "subcomponent": "Equity other than reinvestment of earnings",
                    "counterparty": "Foreign shareholders",
                    "opening_position_2024": position_open,
                    "financial_account_transaction_2025": flow_est,
                    "price_change": _num(e.get("other_price")) * pct / 100 if pd.notna(_num(e.get("other_price"))) else np.nan,
                    "exchange_rate_change": _num(e.get("other_exchange_rate")) * pct / 100 if pd.notna(_num(e.get("other_exchange_rate"))) else np.nan,
                    "other_volume_change": _num(e.get("other_volume")) * pct / 100 if pd.notna(_num(e.get("other_volume"))) else np.nan,
                    "closing_position_2025": position_close,
                    "derivation_status": "ESTIMATE_BOOK_VALUE",
                    "derivation_note": note,
                    "source": "FPC equity table + shareholding structure",
                })

        # Direct-investment reinvestment of earnings: use profit/dividends, not the
        # sign of the equity table's retained-earnings row, because the latter can be
        # entered inconsistently in survey forms.
        reinvestment, reinvestment_note = _profit_reinvestment(prof, company)
        if pd.notna(reinvestment) and pd.notna(di_pct) and di_pct > 0:
            di_re = reinvestment * di_pct / 100
            detail.append({
                "company_name": company,
                "functional_category": "Direct investment",
                "side": "Net incurrence of liabilities",
                "instrument": "Equity and investment fund shares",
                "subcomponent": "Reinvestment of earnings",
                "counterparty": "Foreign direct investors",
                "opening_position_2024": np.nan,
                "financial_account_transaction_2025": di_re,
                "price_change": np.nan,
                "exchange_rate_change": np.nan,
                "other_volume_change": np.nan,
                "closing_position_2025": np.nan,
                "derivation_status": "DERIVED_FROM_PROFIT_AND_DIVIDENDS",
                "derivation_note": reinvestment_note + "; multiplied by survey-reported foreign DI share.",
                "source": "FPC profits/dividends + shareholding structure",
            })

    # --- Non-equity liabilities bridge ---
    if isinstance(ne, pd.DataFrame) and not ne.empty:
        d = ne.copy()
        for c in ["closing_balance_2024", "amount_received_2025", "principal_repayment_2025", "other_price", "other_exchange_rate", "other_volume", "closing_balance_2025"]:
            if c in d.columns:
                d[c] = pd.to_numeric(d[c], errors="coerce")
        for _, r in d.iterrows():
            company = _text(r.get("company_name"))
            if not company:
                continue
            instrument = _text(r.get("instrument"))
            country = _text(r.get("country_or_organisation"))
            rel = _text(r.get("relationship"))
            foreign = _is_foreign_counterparty(country, rel)
            if foreign is not True:
                continue

            # BPM6 functional classification.
            if _relationship_is_di(rel):
                category = "Direct investment"
                sub = "Debt instruments / intercompany debt"
            elif "DEBT SECURITIES" in instrument.upper():
                category = "Portfolio investment"
                sub = "Debt securities"
            else:
                category = "Other investment"
                if "TRADE" in instrument.upper() or "SUPPLIERS" in instrument.upper():
                    sub = "Trade credit and advances"
                elif "LOAN" in instrument.upper():
                    sub = "Loans"
                elif "INSURANCE" in instrument.upper():
                    sub = "Insurance, pension and standardized guarantee schemes"
                elif "PENSION" in instrument.upper():
                    sub = "Insurance, pension and standardized guarantee schemes"
                else:
                    sub = "Other accounts payable"

            opening = _num(r.get("closing_balance_2024"))
            received = _num(r.get("amount_received_2025"))
            repayment = _num(r.get("principal_repayment_2025"))
            closing = _num(r.get("closing_balance_2025"))
            # BOP6 liability flow: increases positive; repayments/decreases negative.
            # We use the magnitude of the repayment field because some respondents
            # enter repayments in parentheses/with a negative sign.
            flow = np.nan
            if pd.notna(received) or pd.notna(repayment):
                flow = (received if pd.notna(received) else 0.0) - (abs(repayment) if pd.notna(repayment) else 0.0)
            residual = np.nan
            if pd.notna(opening) and pd.notna(closing) and pd.notna(flow):
                residual = closing - opening - flow

            detail.append({
                "company_name": company,
                "functional_category": category,
                "side": "Net incurrence of liabilities",
                "instrument": instrument,
                "subcomponent": sub,
                "counterparty": country,
                "relationship": rel,
                "maturity": _text(r.get("maturity")),
                "opening_position_2024": opening,
                "financial_account_transaction_2025": flow,
                "gross_increase_2025": received,
                "gross_decrease_2025": -abs(repayment) if pd.notna(repayment) else np.nan,
                "price_change": _num(r.get("other_price")),
                "exchange_rate_change": _num(r.get("other_exchange_rate")),
                "other_volume_change": _num(r.get("other_volume")),
                "closing_position_2025": closing,
                "iip_reconciliation_residual": residual,
                "derivation_status": "OBSERVED_SURVEY_FLOW",
                "derivation_note": "Amount received minus absolute principal repayment; residual highlights other changes/measurement differences.",
                "source": "FPC non-equity liabilities table",
            })

    detail_df = pd.DataFrame(detail)
    if detail_df.empty:
        detail_df = pd.DataFrame(columns=[
            "company_name", "functional_category", "side", "instrument", "subcomponent",
            "counterparty", "opening_position_2024", "financial_account_transaction_2025",
            "closing_position_2025", "derivation_status", "derivation_note", "source"
        ])

    # Company-level summary in BPM6 presentation.
    if not detail_df.empty:
        numeric = [
            "opening_position_2024", "financial_account_transaction_2025",
            "gross_increase_2025", "gross_decrease_2025", "closing_position_2025",
            "price_change", "exchange_rate_change", "other_volume_change"
        ]
        for c in numeric:
            if c in detail_df.columns:
                detail_df[c] = pd.to_numeric(detail_df[c], errors="coerce")

    summary_rows = []
    for company in sorted(set(detail_df.get("company_name", pd.Series(dtype=str)).dropna().astype(str))):
        g = detail_df[detail_df["company_name"] == company]
        row = {"company_name": company}
        for cat in ["Direct investment", "Portfolio investment", "Other investment"]:
            gc = g[g["functional_category"] == cat]
            row[f"{cat.lower().replace(' ', '_')}_liability_transaction"] = gc["financial_account_transaction_2025"].sum(min_count=1)
            row[f"{cat.lower().replace(' ', '_')}_liability_position_2025"] = gc["closing_position_2025"].sum(min_count=1)
        row["net_incurrence_of_liabilities_fpc_scope"] = g["financial_account_transaction_2025"].sum(min_count=1)
        row["iip_liabilities_fpc_scope_2025"] = g["closing_position_2025"].sum(min_count=1)
        own = ownership[ownership["company_name"] == company]
        if not own.empty:
            for c in ["foreign_ownership_current_pct", "foreign_di_pct_survey_reported", "foreign_di_pct_bpm6_10pct", "foreign_portfolio_equity_pct_estimate", "ownership_classification_status"]:
                row[c] = own.iloc[0].get(c)
        summary_rows.append(row)
    summary_df = pd.DataFrame(summary_rows)

    # Reconciliation / coverage table for compiler use.
    rec_rows = []
    if not detail_df.empty:
        for company, g in detail_df.groupby("company_name"):
            for cat, gc in g.groupby("functional_category"):
                for sub, gs in gc.groupby("subcomponent"):
                    rec_rows.append({
                        "company_name": company,
                        "functional_category": cat,
                        "subcomponent": sub,
                        "opening_position_2024": gs["opening_position_2024"].sum(min_count=1),
                        "financial_account_transaction_2025": gs["financial_account_transaction_2025"].sum(min_count=1),
                        "other_changes_reported": gs[["price_change", "exchange_rate_change", "other_volume_change"]].sum(axis=1, min_count=1).sum(min_count=1) if any(c in gs.columns for c in ["price_change", "exchange_rate_change", "other_volume_change"]) else np.nan,
                        "closing_position_2025": gs["closing_position_2025"].sum(min_count=1),
                        "source_scope": "FPC survey",
                    })
    reconciliation_df = pd.DataFrame(rec_rows)

    coverage = pd.DataFrame([
        {"bpm6_functional_category": "Direct investment", "fpc_coverage": "SUBSTANTIAL", "available_in_fpc": "Equity, profits/dividends, selected intercompany debt", "important_gap": "Foreign ownership allocation and investor-level transactions require compiler review; full assets side not covered by this survey."},
        {"bpm6_functional_category": "Portfolio investment", "fpc_coverage": "PARTIAL", "available_in_fpc": "Potential foreign equity below 10% and debt securities if reported", "important_gap": "Investor-level security classification/positions are not comprehensively captured."},
        {"bpm6_functional_category": "Financial derivatives and employee stock options", "fpc_coverage": "NOT COVERED", "available_in_fpc": "None identified", "important_gap": "Must be sourced from other external-sector datasets."},
        {"bpm6_functional_category": "Other investment", "fpc_coverage": "PARTIAL", "available_in_fpc": "Loans, trade credits/advances, other accounts payable and selected insurance/pension items", "important_gap": "Foreign assets side is not reliably populated in current FPC dataset."},
        {"bpm6_functional_category": "Reserve assets", "fpc_coverage": "NOT COVERED", "available_in_fpc": "None", "important_gap": "Use institution reserve-assets data: monetary gold, SDRs, reserve position in IMF, foreign exchange and other reserve claims."},
    ])

    return detail_df, summary_df, reconciliation_df, coverage


# ============================================================
# FPC-SPECIFIC ANALYTICS
# ============================================================

def calculate_fpc_indicators(tables):
    """Create transparent derived indicators without altering source tables."""
    rows = []
    prof = tables.get("profits", pd.DataFrame())
    eq = tables.get("equity", pd.DataFrame())
    sh = tables.get("shareholding", pd.DataFrame())
    svc = tables.get("services", pd.DataFrame())

    companies = set()
    for df in [prof, eq, sh, svc]:
        if isinstance(df, pd.DataFrame) and "company_name" in df.columns:
            companies.update(
                [x for x in df["company_name"].dropna().astype(str).str.strip() if x]
            )

    for company in sorted(companies):
        row = {"company_name": company}

        if isinstance(prof, pd.DataFrame) and not prof.empty:
            p = prof[prof["company_name"] == company]
            if not p.empty:
                r = p.iloc[0]
                profit = pd.to_numeric(r.get("net_profit_after_tax"), errors="coerce")
                div = pd.to_numeric(r.get("dividends_declared"), errors="coerce")
                paid = pd.to_numeric(r.get("dividends_paid_or_remitted"), errors="coerce")
                retained = pd.to_numeric(r.get("retained_earnings_reported"), errors="coerce")
                row["net_profit_after_tax"] = profit
                row["dividends_declared"] = div
                row["dividends_paid_or_remitted"] = paid
                row["retained_earnings_reported"] = retained
                row["dividend_payout_ratio"] = (
                    div / profit if pd.notna(profit) and profit > 0 and pd.notna(div) else np.nan
                )
                row["dividend_remittance_ratio"] = (
                    paid / div if pd.notna(div) and div != 0 and pd.notna(paid) else np.nan
                )

        if isinstance(eq, pd.DataFrame) and not eq.empty:
            e = eq[eq["company_name"] == company]
            if not e.empty:
                for col in ["closing_2024", "purchase_increase_2025",
                            "sales_decrease_2025", "closing_2025"]:
                    if col in e.columns:
                        row[f"equity_{col}_total"] = pd.to_numeric(
                            e[col], errors="coerce"
                        ).sum(min_count=1)
                if "closing_2025" in e.columns:
                    row["equity_net_change"] = (
                        row.get("equity_closing_2025_total", np.nan)
                        - row.get("equity_closing_2024_total", np.nan)
                    )

        if isinstance(sh, pd.DataFrame) and not sh.empty:
            s = sh[sh["company_name"] == company]
            share_col = next((c for c in [
                "percentage_shareholding_current",
                "percentage_shareholding_previous",
                "share_percent"
            ] if c in s.columns), None)
            if not s.empty and share_col:
                shares = pd.to_numeric(s[share_col], errors="coerce").dropna()
                if not shares.empty:
                    row["shareholding_total_percent"] = shares.sum()
                    row["shareholding_hhi"] = ((shares / 100.0) ** 2).sum() * 10000

        if isinstance(svc, pd.DataFrame) and not svc.empty:
            s = svc[svc["company_name"] == company]
            for col in ["exports", "imports"]:
                if col in s.columns:
                    row[f"services_{col}_total"] = pd.to_numeric(
                        s[col], errors="coerce"
                    ).sum(min_count=1)
            if "exports" in s.columns and "imports" in s.columns:
                ex = row.get("services_exports_total", np.nan)
                im = row.get("services_imports_total", np.nan)
                row["services_net_balance"] = (
                    ex - im if pd.notna(ex) and pd.notna(im) else np.nan
                )

        rows.append(row)

    return pd.DataFrame(rows)


def descriptive_statistics(df):
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    nums = df.select_dtypes(include=[np.number])
    if nums.empty:
        return pd.DataFrame()
    return nums.describe().T.reset_index().rename(columns={"index": "variable"})


def correlation_matrix(df):
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    nums = df.select_dtypes(include=[np.number])
    if nums.shape[1] < 2:
        return pd.DataFrame()
    return nums.corr(numeric_only=True)


def outlier_report(df):
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    nums = df.select_dtypes(include=[np.number])
    rows = []
    for col in nums.columns:
        s = pd.to_numeric(nums[col], errors="coerce").dropna()
        if len(s) < 4:
            continue
        q1, q3 = s.quantile([0.25, 0.75])
        iqr = q3 - q1
        lo, hi = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        mask = (s < lo) | (s > hi)
        rows.append({
            "variable": col,
            "q1": q1, "q3": q3, "iqr": iqr,
            "lower_bound": lo, "upper_bound": hi,
            "outlier_count": int(mask.sum())
        })
    return pd.DataFrame(rows)


def grouped_analysis(df, group_col, value_col, agg):
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    if group_col not in df.columns or value_col not in df.columns:
        return pd.DataFrame()
    temp = df[[group_col, value_col]].copy()
    temp[value_col] = pd.to_numeric(temp[value_col], errors="coerce")
    temp = temp.dropna(subset=[group_col])
    if temp.empty:
        return pd.DataFrame()
    if agg == "sum":
        out = temp.groupby(group_col, dropna=False)[value_col].sum(min_count=1)
    elif agg == "mean":
        out = temp.groupby(group_col, dropna=False)[value_col].mean()
    elif agg == "median":
        out = temp.groupby(group_col, dropna=False)[value_col].median()
    elif agg == "count":
        out = temp.groupby(group_col, dropna=False)[value_col].count()
    else:
        out = temp.groupby(group_col, dropna=False)[value_col].sum(min_count=1)
    return out.reset_index(name=agg)



# ============================================================
# SURVEY ESTIMATION & POPULATION EXPANSION ENGINE
# ============================================================

ESTIMATION_VARIABLES = [
    ("direct_investment_liability_transaction", "Direct investment liability transaction"),
    ("direct_investment_liability_position_2025", "Direct investment liability position"),
    ("portfolio_investment_liability_transaction", "Portfolio investment liability transaction"),
    ("portfolio_investment_liability_position_2025", "Portfolio investment liability position"),
    ("other_investment_liability_transaction", "Other investment liability transaction"),
    ("other_investment_liability_position_2025", "Other investment liability position"),
    ("net_incurrence_of_liabilities_fpc_scope", "Net incurrence of FPC liabilities"),
    ("iip_liabilities_fpc_scope_2025", "FPC-scope IIP liabilities"),
    ("reported_net_profit", "Reported net profit after tax"),
    ("dividends_declared", "Dividends declared"),
    ("reinvestment_of_earnings", "Reinvestment of earnings"),
    ("services_exports", "Services exports"),
    ("services_imports", "Services imports"),
]


def _normalise_key_series(s):
    return (s.astype("string").fillna("").str.upper().str.replace(r"\s+", " ", regex=True).str.strip())


def _read_sampling_frame(uploaded_file, sheet_name=None):
    """Read a sampling frame and intelligently select the frame sheet from Excel."""
    name = str(getattr(uploaded_file, "name", "")).lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    if name.endswith((".xlsx", ".xls", ".xlsm")):
        xl = pd.ExcelFile(uploaded_file)
        if sheet_name and sheet_name in xl.sheet_names:
            return pd.read_excel(uploaded_file, sheet_name=sheet_name)
        preferred = ["sampling_frame", "frame", "sample_frame", "sbr", "business_register"]
        match = next((x for x in preferred if x in xl.sheet_names), None)
        if match:
            return pd.read_excel(uploaded_file, sheet_name=match)
        # Choose the sheet with the strongest frame-field signature.
        best_sheet, best_score = xl.sheet_names[0], -1
        signatures = {"enterprise_id", "enterprise_name", "province", "isic_rev4_section", "size_band", "selection_probability"}
        for sh in xl.sheet_names:
            try:
                head = pd.read_excel(uploaded_file, sheet_name=sh, nrows=5)
                cols = {str(c).strip().lower() for c in head.columns}
                score = len(signatures.intersection(cols))
                if score > best_score:
                    best_sheet, best_score = sh, score
            except Exception:
                continue
        return pd.read_excel(uploaded_file, sheet_name=best_sheet)
    raise ValueError("Sampling frame must be CSV or Excel (.xlsx/.xls/.xlsm).")


def _clean_frame(df):
    if not isinstance(df, pd.DataFrame):
        return pd.DataFrame()
    d = df.copy()
    d.columns = [str(c).strip() for c in d.columns]
    for c in d.columns:
        if d[c].dtype == object:
            d[c] = d[c].astype("string").str.replace(r"\s+", " ", regex=True).str.strip()
    return d


def _flag_from_series(s, default=False):
    if s is None:
        return pd.Series(default, index=[])
    x = s.astype("string").str.upper().str.strip()
    true_vals = {"1", "TRUE", "YES", "Y", "SAMPLED", "RESPONDED", "RESPONDENT", "SELECTED"}
    false_vals = {"0", "FALSE", "NO", "N", "NOT SAMPLED", "NONRESPONDENT", "NON-RESPONDENT"}
    out = pd.Series(default, index=s.index, dtype=bool)
    out[x.isin(true_vals)] = True
    out[x.isin(false_vals)] = False
    numeric = pd.to_numeric(s, errors="coerce")
    out[numeric == 1] = True
    out[numeric == 0] = False
    return out


def build_estimation_microdata(tables, frame, config):
    """Merge respondent survey values to the sampling frame and create analysis variables."""
    if not isinstance(frame, pd.DataFrame) or frame.empty:
        return pd.DataFrame(), "Sampling frame is empty."
    cfg = config or {}
    d = _clean_frame(frame)
    frame_key = cfg.get("frame_key")
    survey_key = cfg.get("survey_key", "company_name")
    if not frame_key or frame_key not in d.columns:
        return pd.DataFrame(), "A valid frame identifier is required."

    summary = tables.get("bop_iip_summary", pd.DataFrame())
    if not isinstance(summary, pd.DataFrame) or summary.empty:
        return pd.DataFrame(), "BOP/IIP summary is not available. Process the FPC surveys first."

    survey = summary.copy()
    if survey_key not in survey.columns:
        if survey_key != "company_name" and "company_name" in survey.columns:
            survey_key = "company_name"
        else:
            return pd.DataFrame(), f"Survey identifier '{survey_key}' is not available."

    d["_frame_key_norm"] = _normalise_key_series(d[frame_key])
    survey["_survey_key_norm"] = _normalise_key_series(survey[survey_key])
    survey = survey.drop_duplicates("_survey_key_norm", keep="last")

    merged = d.merge(
        survey,
        how="left",
        left_on="_frame_key_norm",
        right_on="_survey_key_norm",
        suffixes=("", "_survey")
    )
    merged["responded_derived"] = merged["_survey_key_norm"].notna()

    sampled_col = cfg.get("sampled_col")
    responded_col = cfg.get("responded_col")
    if sampled_col and sampled_col in merged.columns:
        merged["sampled_flag"] = _flag_from_series(merged[sampled_col])
    else:
        # In a response-linked design, a matched respondent is treated as sampled.
        merged["sampled_flag"] = merged["responded_derived"]

    # Response status can be determined from actual survey linkage (recommended) or
    # from a frame field. This distinction matters when a synthetic/test frame has
    # simulated response flags that do not describe the actual uploaded survey.
    response_mode = cfg.get("response_mode", "Survey linkage")
    if response_mode == "Frame response field" and responded_col and responded_col in merged.columns:
        merged["responded_flag"] = _flag_from_series(merged[responded_col])
    else:
        merged["responded_flag"] = merged["responded_derived"]

    # A unit cannot be a respondent if it was not selected.
    merged["responded_flag"] = merged["responded_flag"] & merged["sampled_flag"]

    stratum_cols = [c for c in cfg.get("stratum_cols", []) if c in merged.columns]
    if stratum_cols:
        merged["estimation_stratum"] = merged[stratum_cols].fillna("MISSING").astype(str).agg(" | ".join, axis=1)
    else:
        merged["estimation_stratum"] = "ALL"

    # Optional explicit selection probability.
    pi_col = cfg.get("selection_probability_col")
    if pi_col and pi_col in merged.columns:
        merged["selection_probability"] = pd.to_numeric(merged[pi_col], errors="coerce")
    else:
        merged["selection_probability"] = np.nan

    # Frame totals within the selected design strata.
    frame_counts = merged.groupby("estimation_stratum", dropna=False).agg(
        N_h=(frame_key, "size"),
        n_sampled=("sampled_flag", "sum"),
        n_responded=("responded_flag", "sum")
    ).reset_index()
    merged = merged.merge(frame_counts, on="estimation_stratum", how="left")

    merged["base_weight"] = np.nan
    explicit_pi = merged["selection_probability"].notna() & (merged["selection_probability"] > 0)
    merged.loc[explicit_pi, "base_weight"] = 1.0 / merged.loc[explicit_pi, "selection_probability"]
    eligible = (~explicit_pi) & (merged["n_sampled"] > 0)
    merged.loc[eligible, "base_weight"] = merged.loc[eligible, "N_h"] / merged.loc[eligible, "n_sampled"]

    merged["nonresponse_adjustment"] = np.where(
        merged["n_responded"] > 0,
        merged["n_sampled"] / merged["n_responded"],
        np.nan
    )
    merged["final_weight_precalibration"] = np.where(
        merged["responded_flag"],
        merged["base_weight"] * merged["nonresponse_adjustment"],
        np.nan
    )
    merged["calibration_factor"] = 1.0
    merged["final_weight"] = merged["final_weight_precalibration"]

    # Attach survey variables under stable names.
    prof = tables.get("profits", pd.DataFrame())
    if isinstance(prof, pd.DataFrame) and not prof.empty and "company_name" in prof.columns:
        pp = prof.copy()
        pp["_survey_key_norm"] = _normalise_key_series(pp["company_name"])
        pcols = [c for c in ["net_profit_after_tax", "dividends_declared"] if c in pp.columns]
        if pcols:
            pp = pp.drop_duplicates("_survey_key_norm")[["_survey_key_norm"] + pcols]
            merged = merged.merge(pp, left_on="_survey_key_norm", right_on="_survey_key_norm", how="left", suffixes=("", "_profit"))

    # Reinvestment of earnings is derived only where the underlying profit/dividend inputs exist.
    if "net_profit_after_tax" in merged.columns and "dividends_declared" in merged.columns:
        merged["reinvestment_of_earnings"] = (
            pd.to_numeric(merged["net_profit_after_tax"], errors="coerce") -
            pd.to_numeric(merged["dividends_declared"], errors="coerce")
        )
    else:
        merged["reinvestment_of_earnings"] = np.nan

    services = tables.get("services", pd.DataFrame())
    if isinstance(services, pd.DataFrame) and not services.empty and "company_name" in services.columns:
        sv = services.copy()
        for c in ["exports", "imports"]:
            if c in sv.columns:
                sv[c] = pd.to_numeric(sv[c], errors="coerce")
        sg = sv.groupby("company_name", dropna=False)[[c for c in ["exports", "imports"] if c in sv.columns]].sum(min_count=1).reset_index()
        sg["_survey_key_norm"] = _normalise_key_series(sg["company_name"])
        ren = {"exports":"services_exports", "imports":"services_imports"}
        sg = sg.rename(columns=ren)
        merged = merged.merge(sg.drop(columns=["company_name"], errors="ignore"), left_on="_survey_key_norm", right_on="_survey_key_norm", how="left")

    # Clean up and keep a compact estimation microdata table.
    merged = merged.drop(columns=["_frame_key_norm", "_survey_key_norm"], errors="ignore")
    return merged, ""


def apply_ratio_calibration(micro, config):
    """Apply transparent one-variable ratio calibration to a known frame benchmark."""
    if not isinstance(micro, pd.DataFrame) or micro.empty:
        return micro, "No estimation microdata available."
    d = micro.copy()
    d["calibration_factor"] = 1.0
    d["final_weight"] = d["final_weight_precalibration"]
    benchmark_col = config.get("benchmark_col") if config else None
    if not benchmark_col or benchmark_col not in d.columns:
        return d, "No calibration benchmark selected; final weights equal nonresponse-adjusted design weights."
    target = pd.to_numeric(d[benchmark_col], errors="coerce")
    # Frame benchmark total: all frame units, not just respondents.
    frame_total = target.sum(min_count=1)
    weighted_total = (target * d["final_weight_precalibration"]).sum(min_count=1)
    if pd.isna(frame_total) or pd.isna(weighted_total) or weighted_total == 0:
        return d, "Calibration not applied because the benchmark total could not be calculated safely."
    factor = frame_total / weighted_total
    d["calibration_factor"] = np.where(d["responded_flag"], factor, np.nan)
    d["final_weight"] = d["final_weight_precalibration"] * d["calibration_factor"]
    return d, f"One-variable ratio calibration applied to {benchmark_col}; frame benchmark total = {frame_total:,.4f}."


def _weighted_total(series, weights):
    x = pd.to_numeric(series, errors="coerce")
    w = pd.to_numeric(weights, errors="coerce")
    mask = x.notna() & w.notna() & np.isfinite(x) & np.isfinite(w)
    if not mask.any():
        return np.nan
    return float((x[mask] * w[mask]).sum())


def build_population_estimates(micro, config):
    """Produce national totals, design-based approximate SE/CV/CI and domain estimates."""
    if not isinstance(micro, pd.DataFrame) or micro.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), "Estimation microdata are empty."
    d = micro.copy()
    vars_available = [(c,l) for c,l in ESTIMATION_VARIABLES if c in d.columns]
    rows=[]
    for col,label in vars_available:
        y=pd.to_numeric(d[col], errors="coerce")
        r=d[d["responded_flag"] & y.notna() & d["final_weight"].notna()].copy()
        total=_weighted_total(r[col], r["final_weight"])
        # Approximate stratified variance. This uses respondent-level within-stratum variance
        # and is explicitly labelled approximate when the actual design is more complex.
        var=0.0; variance_available=False
        if not r.empty and "estimation_stratum" in r.columns:
            for _,g in r.groupby("estimation_stratum", dropna=False):
                nh=float(g["n_responded"].iloc[0]) if "n_responded" in g.columns else len(g)
                Nh=float(g["N_h"].iloc[0]) if "N_h" in g.columns else np.nan
                if nh > 1 and pd.notna(Nh):
                    sy=float(pd.to_numeric(g[col],errors="coerce").var(ddof=1))
                    if pd.notna(sy):
                        var += (Nh**2) * max(0.0, 1.0 - nh/Nh) * sy / nh
                        variance_available=True
        se=float(np.sqrt(max(var,0.0))) if variance_available else np.nan
        cv=abs(se/total) if pd.notna(se) and pd.notna(total) and total != 0 else np.nan
        rows.append({
            "estimate_variable":col,"label":label,"estimate":total,
            "standard_error_approx":se,"cv_percent":cv*100 if pd.notna(cv) else np.nan,
            "ci95_lower":total-1.96*se if pd.notna(total) and pd.notna(se) else np.nan,
            "ci95_upper":total+1.96*se if pd.notna(total) and pd.notna(se) else np.nan,
            "respondent_units":len(r),"population_units":len(d),
            "estimation_note":"Weighted population estimate; variance is approximate and assumes the supplied strata/design are adequate."
        })
    estimates=pd.DataFrame(rows)

    # Domain estimates by available geography/industry/size fields.
    domain_cols=[c for c in config.get("domain_cols",[]) if c in d.columns]
    b=[]
    for dc in domain_cols:
        for dom,g0 in d.groupby(dc,dropna=False):
            g=g0[g0["responded_flag"]].copy()
            for col,label in vars_available:
                total=_weighted_total(g[col],g["final_weight"])
                b.append({"domain_type":dc,"domain":dom,"estimate_variable":col,"label":label,"estimate":total,"respondent_units":int(g[col].notna().sum())})
    breakdowns=pd.DataFrame(b)

    # Stratum diagnostics / weighting QA.
    q=[]
    if "estimation_stratum" in d.columns:
        for h,g in d.groupby("estimation_stratum",dropna=False):
            q.append({
                "stratum":h,"population_N":int(g["N_h"].iloc[0]),
                "sample_n":int(g["n_sampled"].iloc[0]),"respondent_n":int(g["n_responded"].iloc[0]),
                "response_rate_percent":100*g["n_responded"].iloc[0]/g["n_sampled"].iloc[0] if g["n_sampled"].iloc[0] else np.nan,
                "mean_final_weight":pd.to_numeric(g["final_weight"],errors="coerce").mean(),
                "max_final_weight":pd.to_numeric(g["final_weight"],errors="coerce").max(),
                "zero_response_flag":bool(g["n_responded"].iloc[0]==0),
            })
    qa=pd.DataFrame(q)
    return estimates, breakdowns, qa, ""


def run_population_estimation(tables, frame, config):
    micro, err=build_estimation_microdata(tables,frame,config)
    if err:
        return False, err
    micro, cal_note=apply_ratio_calibration(micro,config)
    est,br,qa,err=build_population_estimates(micro,config)
    if err:
        return False, err
    st.session_state["estimation_microdata"]=micro
    st.session_state["estimation_linkage"] = micro[[c for c in [config.get("frame_key"), "company_name", "responded_derived", "sampled_flag", "responded_flag"] if c and c in micro.columns]].copy()
    st.session_state["population_estimates"]=est
    st.session_state["population_breakdowns"]=br
    st.session_state["estimation_qa"]=qa
    st.session_state["sampling_frame"]=frame.copy()
    st.session_state["estimation_config"]=config.copy()
    st.session_state["estimation_metadata"]={
        "run_time":datetime.now().isoformat(timespec="seconds"),
        "frame_units":int(len(frame)),
        "sampled_units":int(micro["sampled_flag"].sum()),
        "respondent_units":int(micro["responded_flag"].sum()),
        "strata":int(micro["estimation_stratum"].nunique()),
        "calibration_note":cal_note,
        "weighting_method":"Base design weights + within-stratum nonresponse adjustment + optional one-variable ratio calibration",
        "variance_method":"Approximate stratified variance using within-stratum respondent variance; review for complex/multistage designs",
        "target_population":"All units represented in the supplied sampling frame",
        "matched_survey_units":int(micro["responded_derived"].sum()) if "responded_derived" in micro.columns else 0,
        "unmatched_frame_units":int((~micro["responded_derived"]).sum()) if "responded_derived" in micro.columns else int(len(micro)),
        "response_detection":config.get("response_mode", "Survey linkage"),
    }
    st.session_state["estimation_ready"]=True
    st.session_state["estimation_run"]=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    save_persisted_state()
    return True, ""


def estimation_status():
    return {
        "ready":bool(st.session_state.get("estimation_ready")),
        "frame_units":len(st.session_state.get("sampling_frame",pd.DataFrame())),
        "respondents":int(st.session_state.get("estimation_metadata",{}).get("respondent_units",0)),
        "strata":int(st.session_state.get("estimation_metadata",{}).get("strata",0)),
    }


def population_report_tables():
    return {
        "Population Estimates":st.session_state.get("population_estimates",pd.DataFrame()),
        "Domain Estimates":st.session_state.get("population_breakdowns",pd.DataFrame()),
        "Weighting QA":st.session_state.get("estimation_qa",pd.DataFrame()),
        "Frame-Survey Linkage":st.session_state.get("estimation_linkage",pd.DataFrame()),
        "Estimation Microdata":st.session_state.get("estimation_microdata",pd.DataFrame()),
    }


def build_population_word_report(title):
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    doc=Document()
    sec=doc.sections[0]
    sec.top_margin=Pt(42); sec.bottom_margin=Pt(42); sec.left_margin=Pt(50); sec.right_margin=Pt(50)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run(title); r.bold=True; r.font.size=Pt(18)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"FPC Survey Intelligence & Population Estimation • Generated {datetime.now():%Y-%m-%d %H:%M:%S}").italic=True
    meta=st.session_state.get("estimation_metadata",{})
    cfg=st.session_state.get("estimation_config",{})
    doc.add_heading("1. Report Scope and Statistical Basis",level=1)
    doc.add_paragraph(
        "This population report presents estimates for the target population represented by the supplied sampling frame. "
        "Survey responses are expanded using the documented sample design and weighting assumptions. The report therefore "
        "distinguishes observed respondent values from weighted population estimates and retains the assumptions used for estimation."
    )
    doc.add_paragraph(f"Target population: {meta.get('target_population','Not specified')}")
    doc.add_paragraph(f"Sampling frame units: {meta.get('frame_units','N/A'):,}")
    doc.add_paragraph(f"Sampled units: {meta.get('sampled_units','N/A'):,}")
    doc.add_paragraph(f"Responding units: {meta.get('respondent_units','N/A'):,}")
    doc.add_paragraph(f"Estimation strata: {meta.get('strata','N/A'):,}")
    doc.add_paragraph(f"Weighting method: {meta.get('weighting_method','N/A')}")
    doc.add_paragraph(f"Calibration: {meta.get('calibration_note','None')}")
    doc.add_paragraph(f"Variance method: {meta.get('variance_method','N/A')}")
    if cfg.get("stratum_cols"):
        doc.add_paragraph("Stratification variables: " + ", ".join(cfg["stratum_cols"]))
    doc.add_heading("2. National Population Estimates",level=1)
    est=st.session_state.get("population_estimates",pd.DataFrame())
    if not est.empty:
        show=est.copy(); t=doc.add_table(rows=1,cols=6); t.style="Table Grid"
        headers=["Indicator","Estimate","Approx. SE","CV (%)","95% CI lower","95% CI upper"]
        for i,h in enumerate(headers): t.rows[0].cells[i].text=h
        for _,r0 in show.iterrows():
            cells=t.add_row().cells
            vals=[r0.get("label",""),r0.get("estimate",""),r0.get("standard_error_approx",""),r0.get("cv_percent",""),r0.get("ci95_lower",""),r0.get("ci95_upper","")]
            for i,v in enumerate(vals): cells[i].text="" if pd.isna(v) else (f"{v:,.2f}" if isinstance(v,(int,float,np.integer,np.floating)) else str(v))
    doc.add_heading("3. Domain Estimates",level=1)
    br=st.session_state.get("population_breakdowns",pd.DataFrame())
    if not br.empty:
        show=br.head(250); t=doc.add_table(rows=1,cols=4); t.style="Table Grid"
        for i,h in enumerate(["Domain type","Domain","Indicator","Estimate"]): t.rows[0].cells[i].text=h
        for _,r0 in show.iterrows():
            cells=t.add_row().cells
            vals=[r0.get("domain_type"),r0.get("domain"),r0.get("label"),r0.get("estimate")]
            for i,v in enumerate(vals): cells[i].text="" if pd.isna(v) else (f"{v:,.2f}" if isinstance(v,(int,float,np.integer,np.floating)) else str(v))
    doc.add_heading("4. Sampling, Weighting and Nonresponse Assumptions",level=1)
    doc.add_paragraph(
        "Base weights are calculated as the inverse of the supplied selection probability where available; otherwise the application uses the population-to-sample ratio within the selected estimation strata. "
        "A within-stratum nonresponse adjustment is then applied using sampled units divided by responding units. Units with no responding cases in a stratum are flagged in the weighting QA table rather than silently imputed."
    )
    doc.add_paragraph(
        "Where a calibration benchmark is selected, a transparent one-variable ratio calibration factor is applied so that the weighted respondent total matches the corresponding frame total. "
        "This is a calibration assumption and should be reviewed when independent benchmark totals or a more complex calibration design are available."
    )
    doc.add_heading("5. Precision and Estimation QA",level=1)
    qa=st.session_state.get("estimation_qa",pd.DataFrame())
    if not qa.empty:
        show=qa.head(250); t=doc.add_table(rows=1,cols=7); t.style="Table Grid"
        for i,h in enumerate(["Stratum","Population N","Sample n","Respondent n","Response rate %","Mean weight","Zero response"]): t.rows[0].cells[i].text=h
        for _,r0 in show.iterrows():
            cells=t.add_row().cells
            vals=[r0.get("stratum"),r0.get("population_N"),r0.get("sample_n"),r0.get("respondent_n"),r0.get("response_rate_percent"),r0.get("mean_final_weight"),r0.get("zero_response_flag")]
            for i,v in enumerate(vals): cells[i].text="" if pd.isna(v) else (f"{v:,.2f}" if isinstance(v,(int,float,np.integer,np.floating)) else str(v))
    doc.add_heading("6. Interpretation and Use",level=1)
    doc.add_paragraph(
        "Population estimates are model-free survey expansions under the supplied design assumptions. They should be interpreted with the reported sampling error, nonresponse pattern, frame coverage and calibration assumptions in view. "
        "The FPC/BOP/IIP classification layer remains subject to the coverage and classification limitations documented in the sample report."
    )
    out=io.BytesIO(); doc.save(out); out.seek(0); return out.getvalue()


def _report_numbered_canvas_class():
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    class NumberedCanvas(canvas.Canvas):
        def __init__(self, *args, **kwargs):
            canvas.Canvas.__init__(self, *args, **kwargs)
            self._saved_page_states=[]
        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()
        def save(self):
            total=len(self._saved_page_states)
            for state in self._saved_page_states:
                self.__dict__.update(state)
                self.draw_page_number(total)
                canvas.Canvas.showPage(self)
            canvas.Canvas.save(self)
        def draw_page_number(self, page_count):
            self.saveState()
            w, h = self._pagesize
            # ---------------------------------------------------------
            # Neutral report header accent
            # ---------------------------------------------------------
            if self._pageNumber == 1:
                try:
                    self.setStrokeColor(colors.HexColor("#2A9D8F"))
                    self.setLineWidth(1.2)
                    self.line(30, h - 36, w - 30, h - 36)
                except Exception:
                    pass

            # Footer and page numbering
            # ---------------------------------------------------------
            self.setFont("Helvetica", 7.5)
            self.setFillColor(colors.HexColor("#666666"))

            self.drawString(
                30,
                18,
                "FPC Survey Intelligence • Public Prototype"
            )

            self.drawRightString(
                w - 30,
                18,
                f"Page {self._pageNumber} of {page_count}"
            )

            self.restoreState()
    return NumberedCanvas


def _report_styles():
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER
    styles=getSampleStyleSheet()
    styles.add(ParagraphStyle(name="FPCJustify", parent=styles["BodyText"], alignment=TA_JUSTIFY, leading=12, fontSize=8.7, spaceAfter=6))
    styles.add(ParagraphStyle(name="FPCSmall", parent=styles["BodyText"], fontSize=7.2, leading=9))
    styles.add(ParagraphStyle(name="FPCCenter", parent=styles["BodyText"], alignment=TA_CENTER, fontSize=8, leading=10))
    styles.add(ParagraphStyle(name="FPCSection", parent=styles["Heading1"], fontSize=15, leading=18, textColor=colors.HexColor("#17365D"), spaceBefore=8, spaceAfter=8))
    styles.add(ParagraphStyle(name="FPCSub", parent=styles["Heading2"], fontSize=11, leading=13, textColor=colors.HexColor("#1F4E78"), spaceBefore=6, spaceAfter=5))
    styles.add(ParagraphStyle(name="FPCAnnex", parent=styles["Heading1"], fontSize=14, leading=17, textColor=colors.HexColor("#17365D"), spaceBefore=6, spaceAfter=8))
    return styles


def _p(text, style):
    # Paragraph is imported here deliberately because this helper is called
    # by both the sample and population PDF report builders.  A local import
    # inside only one report builder would leave the other builder with a
    # NameError.
    from reportlab.platypus import Paragraph
    return Paragraph(str(text), style)


def _fmt(v):
    if v is None or (isinstance(v,float) and pd.isna(v)):
        return ""
    if isinstance(v,(int,float,np.integer,np.floating)):
        return f"{v:,.2f}"
    return str(v)


def _report_table(data, col_widths=None, font=7, repeatRows=1):
    from reportlab.lib import colors
    from reportlab.platypus import Table, TableStyle
    from reportlab.lib.enums import TA_LEFT
    wrapped=[]
    for ridx,row in enumerate(data):
        wrapped.append([_p(x, _report_styles()["FPCSmall"]) if not hasattr(x,'wrap') else x for x in row])
    t=Table(wrapped,colWidths=col_widths,repeatRows=repeatRows,hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#1F4E78")),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-1),0.25,colors.HexColor("#B7B7B7")),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F4F7FA")]),
        ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
        ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
    ]))
    return t


def _methodology_story(styles, meta, cfg, population=True):
    story=[]
    story += [_p("3. Analytics Methodology", styles["FPCSection"])]
    story += [_p(
        "The survey methodology follows a probability-sampling and expansion framework in which sampled enterprise responses are used to estimate characteristics of the population represented by the sampling frame. The methodology separates the stages of frame construction, sample selection, response, weighting, nonresponse adjustment, calibration, estimation and precision assessment. The BPM6 Compilation Guide identifies surveys as an important source for balance of payments and international investment position statistics and provides guidance on survey design and estimation (IMF, 2014, Chapters 2–3 and Chapter 8).",
        styles["FPCJustify"])]

    story += [_p("3.1 Target Population and Sampling Frame", styles["FPCSub"]), _p(
        "The target population is defined as the set of statistical units represented by the supplied sampling frame. In production, the frame should be derived from the Statistical Business Register or another authoritative business register and should identify the enterprise, province, ISIC Rev.4 activity, size class and any variables required for stratification and calibration. The frame must be assessed for coverage, duplicates, out-of-scope units and units with uncertain status before sample selection.", styles["FPCJustify"])]
    story += [_p("A frame count is written as N = Σ_h N_h, where N_h is the number of population units in stratum h. For example, if a stratum contains 140 enterprises, then N_h = 140. The population total is the sum across all strata.", styles["FPCJustify"])]

    story += [_p("3.2 Sampling Design and Stratification", styles["FPCSub"]), _p(
        "For the synthetic demonstration frame, the principal strata are constructed as province × ISIC Rev.4 section × size band. Large enterprises can be designated certainty (take-all) units, while small and medium enterprises can be sampled probabilistically within strata. A certainty unit has inclusion probability π_i = 1. For other sampled units, π_i is the probability of selection under the specified design. The actual production design must use the documented probabilities generated by the approved sample-selection procedure.", styles["FPCJustify"])]
    story += [_p("Example: suppose a stratum contains N_h = 200 enterprises and n_h = 40 are selected by simple random sampling. Then π_i = n_h/N_h = 40/200 = 0.20 for each unit and the base design weight is d_i = 1/π_i = 5. Each responding sampled enterprise initially represents five population enterprises before nonresponse adjustment.", styles["FPCJustify"])]

    story += [_p("3.3 Data Collection and Response", styles["FPCSub"]), _p(
        "Survey collection should maintain a unique enterprise identifier so that sampled units, returned questionnaires, follow-up records and compiled observations can be linked without duplication. A response indicator should distinguish sampled units from actual responding units. Item nonresponse (a missing variable within a returned questionnaire) should be distinguished from unit nonresponse (a sampled enterprise that does not provide a usable questionnaire). Missing, not applicable and zero should not be treated as interchangeable states.", styles["FPCJustify"])]
    story += [_p("The application records a survey-linkage response mode where possible. This is important when a synthetic or external frame contains a response flag that does not necessarily mean that the current FPC questionnaire was actually received and processed.", styles["FPCJustify"])]

    story += [_p("3.4 Base Design Weights", styles["FPCSub"]), _p(
        "The fundamental design-weight formula is d_i = 1/π_i, where π_i is the inclusion probability of unit i. Under simple random sampling without replacement within a stratum, π_i = n_h/N_h and therefore d_i = N_h/n_h. Under unequal-probability sampling, the unit-specific inclusion probability must be retained and used directly. The design weight represents the number of population units represented by the sampled unit under the selection design.", styles["FPCJustify"])]
    story += [_p("Example: if π_i = 0.05, then d_i = 1/0.05 = 20. The sampled enterprise represents 20 population units before adjustments for nonresponse or calibration.", styles["FPCJustify"])]

    story += [_p("3.5 Nonresponse Adjustment", styles["FPCSub"]), _p(
        "Where response rates differ across strata, design weights alone do not address unit nonresponse. A common within-stratum adjustment is A_h = n_h/r_h, where n_h is the number selected in stratum h and r_h is the number of responding units. The adjusted weight is w_i = d_i × A_h for respondents in that stratum. This adjustment is an expansion assumption: the responding units are treated as representing the sampled nonrespondents within the adjustment cell.", styles["FPCJustify"])]
    story += [_p("Example: if 40 enterprises are selected and 32 respond, A_h = 40/32 = 1.25. If a respondent has a design weight of 5, its nonresponse-adjusted weight is 5 × 1.25 = 6.25.", styles["FPCJustify"])]
    story += [_p("If r_h = 0, the adjustment is undefined. Such a stratum should be flagged for review rather than assigned an arbitrary weight. The IMF BPM6 Compilation Guide notes that expansion factors can be used to adjust for nonresponse and discusses poststratified estimation as an alternative when population characteristics are known (IMF, 2014, para. 8.24).", styles["FPCJustify"])]

    story += [_p("3.6 Calibration and Benchmarking", styles["FPCSub"]), _p(
        "Calibration adjusts respondent weights so that weighted totals reproduce trusted auxiliary benchmarks. For a single benchmark variable X, a simple ratio calibration factor can be written g = X_pop / Σ_s w_i X_i, and the calibrated weight becomes w_i* = w_i × g. Multiple benchmarks can require raking, generalized regression calibration or another approved calibration method. Calibration should only use independent and reliable benchmarks and should be documented because it changes the final expansion weights.", styles["FPCJustify"])]
    story += [_p("Example: if the weighted respondent employment is 92,000 but the frame benchmark is 100,000, the ratio factor is g = 100,000/92,000 = 1.08696. A respondent with weight 6.25 would receive a calibrated weight of 6.25 × 1.08696 = 6.7935.", styles["FPCJustify"])]

    story += [_p("3.7 Estimation of Population Totals", styles["FPCSub"]), _p(
        "For a population total Y, the weighted estimator is Ŷ = Σ_{i∈s_r} w_i y_i, where s_r is the set of responding sampled units, y_i is the reported value and w_i is the final survey weight. For a mean, the weighted estimator is ȳ_w = Σ w_i y_i / Σ w_i. For a proportion, the numerator and denominator are weighted consistently. Domain estimates are obtained by applying the same estimator to units belonging to the relevant domain.", styles["FPCJustify"])]
    story += [_p("Worked example: three responding enterprises report turnover of US$10m, US$20m and US$30m and have final weights 5, 8 and 12. The estimated population turnover is (5×10) + (8×20) + (12×30) = US$570m. The corresponding weighted number of enterprises is 5 + 8 + 12 = 25. A weighted mean turnover would therefore be US$570m/25 = US$22.8m per represented enterprise.", styles["FPCJustify"])]

    story += [_p("3.8 Variance, Standard Errors and Confidence Intervals", styles["FPCSub"]), _p(
        "Sampling estimates are subject to sampling variability. For simple stratified sampling, an approximate variance for a total can be expressed as Var(Ŷ) ≈ Σ_h N_h²(1−f_h)s_h²/n_h, where f_h=n_h/N_h is the sampling fraction and s_h² is the within-stratum sample variance. The standard error is SE(Ŷ)=√Var(Ŷ). A conventional approximate 95 percent confidence interval is Ŷ ± 1.96×SE(Ŷ), subject to the assumptions of the variance estimator. For complex designs involving clustering, multiple stages or unequal probabilities, a design-specific variance estimator should replace the simplified expression.", styles["FPCJustify"])]
    story += [_p("The coefficient of variation is CV = 100×SE(Ŷ)/|Ŷ|. CV provides a scale-free indication of sampling precision but should not be interpreted as a measure of nonsampling error.", styles["FPCJustify"])]

    story += [_p("3.9 Data Quality, Editing and Imputation", styles["FPCSub"]), _p(
        "Data validation should occur before estimation. Logical checks, range checks, accounting relationships, duplicate checks, unit checks and consistency with prior-period information should be documented. Item nonresponse may require donor or model-based imputation, but an imputed value must be flagged and its contribution to the estimate retained in the audit trail. Unit nonresponse should not be converted to zero. The estimation engine should distinguish reported, imputed, derived and unavailable values.", styles["FPCJustify"])]

    story += [_p("3.10 BPM6 Classification and BOP/IIP Linkage", styles["FPCSub"]), _p(
        "For external-sector compilation, survey variables are mapped to BPM6 functional categories and instruments. BPM6 defines an immediate direct investment relationship when a nonresident direct investor owns equity giving at least 10 percent of voting power; control is above 50 percent and significant influence is 10–50 percent (IMF, 2009, paras. 6.8–6.12). Direct investment includes associated debt positions subject to the BPM6 exceptions. Portfolio investment covers securities not included in direct investment or reserve assets, while other investment is a residual category including loans, trade credit and advances, insurance and pension entitlements and other accounts (IMF, 2009, Chapter 6).",
        styles["FPCJustify"])]
    story += [_p("The Jamaica IIP report supplied as the design reference illustrates the publication logic: it defines the IIP as a stock statement at a point in time and explains that changes in positions reflect transactions, valuation changes and other changes. Its summary table and subsequent sections separate direct investment, portfolio investment, other investment and reserve assets, while its glossary gives concise BPM6-based definitions (Bank of Jamaica, 2014, pp. 2–8).", styles["FPCJustify"])]

    story += [_p("3.11 Estimation Limitations", styles["FPCSub"]), _p(
        "Population estimates are conditional on the quality and coverage of the sampling frame, the documented sample-selection probabilities, response patterns, the validity of calibration benchmarks and the accuracy of the reported survey data. A probability sample can support population inference only to the extent that the sample design and response process are appropriately represented in the weights. In particular, a synthetic demonstration frame must not be interpreted as an official population estimate.", styles["FPCJustify"])]
    return story


def _reference_story(styles):
    return [
        _p("References", styles["FPCSection"]),
        _p("[1] International Monetary Fund (IMF). 2009. Balance of Payments and International Investment Position Manual, Sixth Edition (BPM6). Washington, DC. See especially Chapter 6 on functional categories and Chapter 7 on the international investment position.", styles["FPCJustify"]),
        _p("[2] International Monetary Fund (IMF). 2014. Balance of Payments and International Investment Position Manual, Sixth Edition: Compilation Guide. Washington, DC. See Chapters 2–3 on surveys and Chapter 8 on crosscutting estimation issues, including expansion and poststratification.", styles["FPCJustify"]),
        _p("[3] Bank of Jamaica. 2014. Quarterly International Investment Position Report, December 2014 Quarter. International Accounts Unit, Economic Information Publications Department. The supplied report was used as the publication-design reference for the report structure, summary presentation, tables, analytical discussion, figures and glossary.", styles["FPCJustify"]),
        _p("[4] International Monetary Fund. Coordinated Direct Investment Survey Guide / BPM6 direct-investment guidance, used for survey and direct-investment relationship concepts.", styles["FPCJustify"]),
    ]


def build_population_pdf_report(title):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, KeepTogether
    styles=_report_styles()
    meta=st.session_state.get("estimation_metadata",{}); cfg=st.session_state.get("estimation_config",{})
    out=io.BytesIO()
    doc=SimpleDocTemplate(out,pagesize=landscape(A4),rightMargin=30,leftMargin=30,topMargin=32,bottomMargin=30)
    story=[]
    story += [Spacer(1,45),_p("FPC SURVEY INTELLIGENCE",styles["FPCCenter"]),Spacer(1,12),_p(title,styles["Title"]),Spacer(1,10),_p("Foreign Private Capital Survey • Population Estimates, BOP and IIP Statistical Report",styles["FPCCenter"]),Spacer(1,12),_p(f"Reference period: {st.session_state.get('survey_year', '2025')} • Generated {datetime.now():%Y-%m-%d %H:%M:%S}",styles["FPCCenter"]),Spacer(1,55),_p("Prepared using the FPC Survey Intelligence & Population Estimation application",styles["FPCCenter"]),PageBreak()]

    story += [_p("1. Executive Summary",styles["FPCSection"]),_p(
        "This report presents population estimates generated from the responding FPC survey units using the supplied probability-sampling frame and the documented weighting procedure. The report follows the publication principle illustrated by the supplied Bank of Jamaica IIP report, which presents a concise statistical overview followed by analytical discussion, tables and supporting definitions. The IIP is a stock statement at a point in time and changes in positions reflect transactions, valuation changes and other changes (Bank of Jamaica, 2014, p. 2).",
        styles["FPCJustify"])]
    info=[
        ["Item","Value"],
        ["Target population",meta.get("target_population","Not specified")],
        ["Sampling frame units",f"{meta.get('frame_units','N/A'):,}"],
        ["Sampled units",f"{meta.get('sampled_units','N/A'):,}"],
        ["Responding units",f"{meta.get('respondent_units','N/A'):,}"],
        ["Estimation strata",f"{meta.get('strata','N/A'):,}"],
        ["Stratification variables",", ".join(cfg.get("stratum_cols",[])) or "Not specified"],
        ["Weighting method",meta.get("weighting_method","N/A")],
        ["Calibration",meta.get("calibration_note","None")],
    ]
    story += [_report_table(info,[155,610]),Spacer(1,10)]

    est=st.session_state.get("population_estimates",pd.DataFrame())
    story += [_p("2. Population Estimates",styles["FPCSection"])]
    if isinstance(est,pd.DataFrame) and not est.empty:
        data=[["Indicator","Estimate","Approx. SE","CV %","95% CI lower","95% CI upper"]]
        for _,r in est.iterrows():
            data.append([r.get("label",""),_fmt(r.get("estimate")),_fmt(r.get("standard_error_approx")),_fmt(r.get("cv_percent")),_fmt(r.get("ci95_lower")),_fmt(r.get("ci95_upper"))])
        story += [_report_table(data,[210,105,100,70,105,105])]
    else:
        story += [_p("No population estimates are available in the current estimation run.",styles["FPCJustify"])]

    story += _methodology_story(styles,meta,cfg,population=True)

    story += [_p("4. Statistical Results by Domain",styles["FPCSection"])]
    br=st.session_state.get("population_breakdowns",pd.DataFrame())
    if isinstance(br,pd.DataFrame) and not br.empty:
        data=[["Domain type","Domain","Indicator","Estimate"]]
        for _,r in br.head(500).iterrows():
            data.append([r.get("domain_type",""),r.get("domain",""),r.get("label",""),_fmt(r.get("estimate"))])
        story += [_report_table(data,[100,170,330,110])]
    else:
        story += [_p("No domain estimates are available.",styles["FPCJustify"])]

    story += [_p("5. Estimation Quality Assurance",styles["FPCSection"])]
    qa=st.session_state.get("estimation_qa",pd.DataFrame())
    if isinstance(qa,pd.DataFrame) and not qa.empty:
        data=[["Stratum","Population N","Sample n","Respondent n","Response rate %","Mean weight","Zero response"]]
        for _,r in qa.head(500).iterrows():
            data.append([r.get("stratum",""),_fmt(r.get("population_N")),_fmt(r.get("sample_n")),_fmt(r.get("respondent_n")),_fmt(r.get("response_rate_percent")),_fmt(r.get("mean_final_weight")),r.get("zero_response_flag","")])
        story += [_report_table(data,[240,75,70,85,90,90,90])]
    story += [_p(str(meta.get("variance_method","Approximate design-based variance.")),styles["FPCJustify"])]

    story += [PageBreak(),_p("6. Statistical Annex",styles["FPCAnnex"]),_p("The annex is intentionally presented in landscape orientation so that wide statistical tables can be read without clipping. This follows the supplied publication's use of tabular presentation for the main IIP results.",styles["FPCJustify"])]
    # Annex A: population estimates
    story += [_p("Annex A. National Population Estimates",styles["FPCSub"])]
    if isinstance(est,pd.DataFrame) and not est.empty:
        data=[[str(c) for c in est.columns]]
        for _,r in est.iterrows(): data.append([_fmt(r.get(c)) for c in est.columns])
        story += [_report_table(data,[90]*len(data[0]))]
    # Annex B: linkage
    link=st.session_state.get("estimation_linkage",pd.DataFrame())
    story += [_p("Annex B. Frame–Survey Linkage",styles["FPCSub"])]
    if isinstance(link,pd.DataFrame) and not link.empty:
        cols=list(link.columns)
        data=[[str(c) for c in cols]]
        for _,r in link.head(500).iterrows(): data.append([_fmt(r.get(c)) for c in cols])
        cw=[max(55,min(150,700/max(1,len(cols))))]*len(cols)
        story += [_report_table(data,cw)]
    # Annex C: domain estimates
    story += [_p("Annex C. Domain Estimates",styles["FPCSub"])]
    if isinstance(br,pd.DataFrame) and not br.empty:
        cols=list(br.columns)
        data=[[str(c) for c in cols]]
        for _,r in br.head(500).iterrows(): data.append([_fmt(r.get(c)) for c in cols])
        cw=[max(55,min(180,700/max(1,len(cols))))]*len(cols)
        story += [_report_table(data,cw)]
    # Annex D formula summary
    story += [_p("Annex D. Core Estimation Formulae",styles["FPCSub"])]
    formula_data=[
        ["Concept","Formula","Illustrative example"],
        ["Inclusion probability","π_i = n_h/N_h (simple random sampling within stratum)","40 selected from 200 → π=0.20"],
        ["Design weight","d_i = 1/π_i","1/0.20 = 5"],
        ["Nonresponse adjustment","A_h = n_h/r_h","40 selected, 32 responding → 1.25"],
        ["Adjusted weight","w_i = d_i×A_h","5×1.25 = 6.25"],
        ["Ratio calibration","g = X_pop/Σw_iX_i; w_i*=w_i×g","100,000/92,000=1.08696"],
        ["Population total","Ŷ = Σw_i y_i","5×10 + 8×20 + 12×30 = 570"],
        ["Weighted mean","ȳ_w = Σw_i y_i/Σw_i","570/25 = 22.8"],
        ["Approx. 95% CI","Ŷ ± 1.96×SE(Ŷ)","Estimate ± 1.96×SE"],
        ["Coefficient of variation","CV = 100×SE/|Ŷ|","SE/estimate expressed as %"],
    ]
    story += [_report_table(formula_data,[165,300,250])]
    story += _reference_story(styles)
    story += [_p("Glossary",styles["FPCSection"]),_p("Direct investment: a cross-border investment relationship involving control or significant influence; BPM6 operationalizes an immediate relationship at 10 percent or more of voting power. Portfolio investment: cross-border transactions and positions in securities not included in direct investment or reserve assets. Other investment: the residual financial category including loans, trade credit and advances, insurance and pension entitlements and other accounts. International investment position: a statement of external financial assets and liabilities at a point in time.",styles["FPCJustify"])]

    NumberedCanvas=_report_numbered_canvas_class()
    doc.build(story,canvasmaker=NumberedCanvas)
    out.seek(0)
    return out.getvalue()

def population_excel_bytes():
    payload=population_report_tables()
    return full_stage_excel_bytes({**payload,"sampling_frame":st.session_state.get("sampling_frame",pd.DataFrame())},"Population_Estimation")


# ============================================================
# REPORT GENERATION
# ============================================================

def _report_summary(tables):
    cm = tables.get("company_master", pd.DataFrame())
    qa = tables.get("qa", pd.DataFrame())
    prof = tables.get("profits", pd.DataFrame())
    av = tables.get("accounting_validation", pd.DataFrame())

    n = company_count(cm)
    qa_counts = qa["status"].value_counts().to_dict() if (
        isinstance(qa, pd.DataFrame) and not qa.empty and "status" in qa.columns
    ) else {}

    profit = np.nan
    div = np.nan
    if isinstance(prof, pd.DataFrame) and not prof.empty:
        profit = pd.to_numeric(prof.get("net_profit_after_tax"), errors="coerce").sum(min_count=1)
        div = pd.to_numeric(prof.get("dividends_declared"), errors="coerce").sum(min_count=1)

    return {
        "companies": n,
        "qa_pass": qa_counts.get("PASS", 0),
        "qa_review": qa_counts.get("REVIEW", 0),
        "qa_fail": qa_counts.get("FAIL", 0),
        "qa_info": qa_counts.get("INFO", 0),
        "profit": profit,
        "dividends": div,
        "accounting_objections": int((av["status"] == "OBJECTION").sum()) if isinstance(av, pd.DataFrame) and "status" in av.columns else 0,
        "accounting_pass": int((av["status"] == "PASS").sum()) if isinstance(av, pd.DataFrame) and "status" in av.columns else 0,
    }


def build_word_report(title, tables, include_tables=None):
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    include_tables = include_tables or ["company_master", "financials", "accounting_validation", "qa"]
    summary = _report_summary(tables)

    doc = Document()
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title_p.add_run(title)
    run.bold = True
    run.font.size = Pt(18)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(
        f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} • "
        f"FPC Survey Intelligence & Population Estimation"
    ).italic = True

    doc.add_heading("Executive Summary", level=1)
    for label, value in [
        ("Companies processed", summary["companies"]),
        ("QA PASS records", summary["qa_pass"]),
        ("QA REVIEW records", summary["qa_review"]),
        ("QA FAIL records", summary["qa_fail"]),
        ("QA INFO records", summary["qa_info"]),
        ("Reported net profit", "N/A" if pd.isna(summary["profit"]) else f"US$ {summary['profit']:,.2f}"),
        ("Dividends declared", "N/A" if pd.isna(summary["dividends"]) else f"US$ {summary['dividends']:,.2f}"),
        ("Accounting objections", summary["accounting_objections"]),
        ("Accounting checks passed", summary["accounting_pass"]),
    ]:
        doc.add_paragraph(f"{label}: {value}")

    for key in include_tables:
        df = tables.get(key, pd.DataFrame())
        if not isinstance(df, pd.DataFrame) or df.empty:
            continue
        doc.add_heading(key.replace("_", " ").title(), level=1)
        show = df.copy().head(50)
        table = doc.add_table(rows=1, cols=len(show.columns))
        table.style = "Table Grid"
        for i, c in enumerate(show.columns):
            table.rows[0].cells[i].text = str(c)
        for _, row in show.iterrows():
            cells = table.add_row().cells
            for i, c in enumerate(show.columns):
                val = row[c]
                cells[i].text = "" if pd.isna(val) else str(val)

    doc.add_paragraph(
        "Note: extracted and OCR-derived values should be verified against the source "
        "questionnaire before official statistical compilation."
    )

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


def build_ppt_report(title, tables):
    from pptx import Presentation
    from pptx.util import Inches, Pt

    summary = _report_summary(tables)
    prs = Presentation()

    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = title
    slide.placeholders[1].text = (
        f"FPC Survey Intelligence & Population Estimation\n"
        f"Generated {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "Executive Summary"
    tf = slide.placeholders[1].text_frame
    tf.clear()
    for label, value in [
        ("Companies processed", summary["companies"]),
        ("QA PASS", summary["qa_pass"]),
        ("QA REVIEW", summary["qa_review"]),
        ("QA FAIL", summary["qa_fail"]),
        ("Reported net profit", "N/A" if pd.isna(summary["profit"]) else f"US$ {summary['profit']:,.2f}"),
        ("Dividends declared", "N/A" if pd.isna(summary["dividends"]) else f"US$ {summary['dividends']:,.2f}"),
        ("Accounting objections", summary["accounting_objections"]),
        ("Accounting checks passed", summary["accounting_pass"]),
    ]:
        p = tf.add_paragraph()
        p.text = f"{label}: {value}"
        p.font.size = Pt(20)

    # Add compact table slides.
    for key in ["company_master", "financials", "accounting_validation", "qa"]:
        df = tables.get(key, pd.DataFrame())
        if not isinstance(df, pd.DataFrame) or df.empty:
            continue
        show = df.head(12).copy()
        cols = list(show.columns[:8])
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = key.replace("_", " ").title()
        rows = len(show) + 1
        cols_n = len(cols)
        table_shape = slide.shapes.add_table(
            rows, cols_n, Inches(0.3), Inches(1.2), Inches(12.7), Inches(5.5)
        )
        table = table_shape.table
        for j, c in enumerate(cols):
            table.cell(0, j).text = str(c)
        for i, (_, row) in enumerate(show.iterrows(), start=1):
            for j, c in enumerate(cols):
                val = row[c]
                table.cell(i, j).text = "" if pd.isna(val) else str(val)[:40]

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()


def build_pdf_report(title, tables):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Spacer, PageBreak
    styles=_report_styles(); summary=_report_summary(tables)
    output=io.BytesIO()
    doc=SimpleDocTemplate(output,pagesize=landscape(A4),rightMargin=30,leftMargin=30,topMargin=32,bottomMargin=30)
    story=[Spacer(1,45),_p("FPC SURVEY INTELLIGENCE",styles["FPCCenter"]),Spacer(1,12),_p(title,styles["Title"]),Spacer(1,10),_p("Foreign Private Capital Survey • Sample Statistical Report",styles["FPCCenter"]),Spacer(1,55),_p(f"Generated {datetime.now():%Y-%m-%d %H:%M:%S}",styles["FPCCenter"]),PageBreak()]
    story += [_p("1. Executive Summary",styles["FPCSection"])]
    summary_rows=[["Indicator","Value"],["Companies processed",summary["companies"]],["QA PASS",summary["qa_pass"]],["QA REVIEW",summary["qa_review"]],["QA FAIL",summary["qa_fail"]],["QA INFO",summary["qa_info"]],["Reported net profit","N/A" if pd.isna(summary["profit"]) else f"US$ {summary['profit']:,.2f}"],["Dividends declared","N/A" if pd.isna(summary["dividends"]) else f"US$ {summary['dividends']:,.2f}"],["Accounting objections",summary["accounting_objections"]],["Accounting checks passed",summary["accounting_pass"]]]
    story += [_report_table(summary_rows,[220,180]),Spacer(1,10),_p("This sample report describes the enterprises and values actually processed. It does not, by itself, constitute a population estimate. Population inference requires a documented probability-sampling design, response information and appropriate weights.",styles["FPCJustify"])]
    sec=2
    for key in ["company_master","financials","accounting_validation","bop_iip_summary","qa"]:
        df=tables.get(key,pd.DataFrame())
        if not isinstance(df,pd.DataFrame) or df.empty: continue
        story += [_p(f"{sec}. {key.replace('_',' ').title()}",styles["FPCSection"])]
        show=df.head(100); cols=list(show.columns[:10]); data=[[str(c) for c in cols]]
        for _,r in show.iterrows(): data.append([_fmt(r.get(c))[:45] for c in cols])
        story += [_report_table(data,[700/max(1,len(cols))]*len(cols)),Spacer(1,8)]
        sec += 1
    story += [PageBreak(),_p(f"{sec}. Statistical Annex",styles["FPCAnnex"]),_p("The annex is presented in landscape orientation to accommodate wide statistical tables and preserve readability.",styles["FPCJustify"])]
    sec += 1
    for key,df in tables.items():
        if not isinstance(df,pd.DataFrame) or df.empty: continue
        story += [_p(f"{sec}.{len(story)} {key.replace('_',' ').title()}",styles["FPCSub"])]
        show=df.head(200); cols=list(show.columns[:14]); data=[[str(c) for c in cols]]
        for _,r in show.iterrows(): data.append([_fmt(r.get(c))[:35] for c in cols])
        story += [_report_table(data,[700/max(1,len(cols))]*len(cols)),Spacer(1,10)]
        sec += 1
    story += [_p("Source and compilation note",styles["FPCSection"]),_p("The report is generated from the processed FPC survey dataset. Extracted and OCR-derived values should be verified against the source questionnaire before official compilation. BPM6 provides the conceptual basis for the external-sector classification layer.",styles["FPCJustify"]),*_reference_story(styles)]
    NumberedCanvas=_report_numbered_canvas_class(); doc.build(story,canvasmaker=NumberedCanvas); output.seek(0); return output.getvalue()


# Initialize state only after all extraction/consolidation functions are defined.
initialise_state()

# ============================================================
# NAVIGATION
# ============================================================

PAGES = [
    "🏠 Home",
    "ℹ️ About the Program",
    "📚 Analytics Methodology",
    "✍️ The Author",
    "©️ Copyright",
    "📥 1. Upload & Extract",
    "🔎 2. Review & Edit",
    "🧹 3. Data Cleaning",
    "📐 4. Sampling Frame & Design",
    "⚖️ 5. Weighting & Estimation",
    "🌍 6. Population Estimates",
    "📊 7. Dashboard",
    "🏢 8. Company Explorer",
    "📋 9. Data Tables",
    "📈 10. Analytics Studio",
    "✅ 11. QA & Audit",
    "📝 12. Reports & Tables",
    "📤 13. Export Center",
]
_pending_page = st.session_state.pop("_fpc_pending_page", None)
if _pending_page in PAGES:
    st.session_state["fpc_page"] = _pending_page

if st.session_state.get("fpc_page") not in PAGES:
    st.session_state["fpc_page"] = PAGES[0]

# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.markdown(
        """
        <div class="app-lockup">
            <div>
                <div class="bank-name">FPC SURVEY INTELLIGENCE</div>
                <div class="system-name">POPULATION ESTIMATION & ANALYTICS</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption("Foreign Private Capital • BOP • IIP • BPM6")
    page = st.radio("Navigation", PAGES, key="fpc_page")

    st.divider()

    st.info("Public prototype • use synthetic or anonymised test data only")

    if TESSERACT_PATH:
        st.success("OCR: Tesseract detected")
    else:
        st.warning("OCR: Tesseract not detected")

    st.caption(f"Application build: {APP_VERSION}")

    n = dataset_company_count()
    st.metric("Companies", n)

    stage1 = "✅" if st.session_state.get("processed") else "⬜"
    stage2 = "✅" if st.session_state.get("cleaned") else "⬜"
    st.caption(f"{stage1} Extraction / Review")
    st.caption(f"{stage2} Data Cleaning")
    av_status_sidebar = st.session_state.get("accounting_validation_status", "NOT_RUN")
    if av_status_sidebar == "ACCOUNTING RULES VALIDATED":
        st.caption("🟢 Accounting rules validated")
    elif av_status_sidebar == "OBJECTIONS RAISED":
        st.caption("🔴 Accounting objections raised")
    elif av_status_sidebar == "PENDING_REVALIDATION":
        st.caption("🟠 Accounting validation required again")
    else:
        st.caption("⬜ Accounting validation pending")
    est_status = "🟢 Population estimation ready" if st.session_state.get("estimation_ready") else "⬜ Population estimation pending"
    st.caption(est_status)
    st.caption(
        "⬜ Analytics / Reporting"
        if not st.session_state.get("processed")
        else "🟢 Analytics / Reporting"
    )

    if st.session_state.get("last_run"):
        st.caption(f"Last extraction: {st.session_state['last_run']}")
    if st.session_state.get("last_cleaned"):
        st.caption(f"Last cleaning: {st.session_state['last_cleaned']}")

    if st.session_state.get("_fpc_restore_error"):
        st.warning(
            "A previous saved dataset could not be restored cleanly. "
            "The application has isolated the saved state; process the PDFs again "
            "to create a fresh canonical dataset.\n\n"
            + str(st.session_state.get("_fpc_restore_error"))
        )

# ============================================================
# HEADER
# ============================================================

st.markdown(
    """
    <div class="main-header" style="display:flex;align-items:center;gap:.8rem;">
        <div>
            <h1 style="margin:0;color:#1F4E79;">FPC Survey Intelligence & Population Estimation</h1>
            <div class="subtle">Foreign Private Capital Survey • BPM6 • BOP • IIP • Extraction • Review • QA • Analytics • Reporting</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

if st.session_state.get("_fpc_restored"):
    st.success(
        f"Restored the last valid FPC dataset from local storage "
        f"({company_count(st.session_state.get('company_master'))} companies)."
    )

# ============================================================
# HOME
# ============================================================

if page == "🏠 Home":
    st.markdown(
        f"""
        <div class="app-hero">
            <div style="display:flex;align-items:center;gap:1.1rem;position:relative;z-index:2;">
                <div>
                    <div class="eyebrow">PUBLIC PROTOTYPE</div>
                    <h1>FPC Survey Intelligence</h1>
                    <p>Foreign Private Capital Survey • BOP &amp; IIP • BPM6 Statistical Intelligence Platform</p>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="app-info">
        <strong>Welcome.</strong> This application provides an end-to-end production workflow
        for extracting, reviewing, cleaning, validating and analysing Foreign Private Capital
        Survey questionnaires, with compiler-support views for BPM6, Balance of Payments and
        International Investment Position work.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("### Platform capabilities")
    cards = st.columns(4)
    card_data = [
        ("📄 Survey Processing", "Extracts structured information from typed and scanned FPC questionnaires while retaining raw values and OCR review flags."),
        ("🔍 Data Quality", "Provides review, reconciliation, cleaning and QA tools before analytical use."),
        ("📊 BOP & IIP Intelligence", "Maps available FPC information into BPM6 functional categories and compiler-support transaction/position views."),
        ("📦 Reporting & Export", "Produces separate Sample and Population reports, with documented weighting assumptions and statistical precision."),
    ]
    for col, (title, desc) in zip(cards, card_data):
        with col:
            st.markdown(
                f'<div class="app-card"><h4>{title}</h4><p>{desc}</p></div>',
                unsafe_allow_html=True,
            )

    st.markdown("### Current platform status")
    cols = st.columns(6)
    for c, label, value in zip(
        cols,
        ["PDFs", "Companies", "Review", "Cleaning", "Analytics", "Reports"],
        [
            "Ready",
            dataset_company_count(),
            "Available" if st.session_state["processed"] else "Pending",
            "Completed" if st.session_state["cleaned"] else "Pending",
            "Available" if st.session_state["processed"] else "Pending",
            "Excel • Word • PPT • PDF",
        ],
    ):
        c.metric(label, value)

    st.markdown("### Production workflow")
    workflow_cols = st.columns(9)
    workflow = [
        ("01", "Upload", "Secure questionnaire upload"),
        ("02", "Extract", "PDF + OCR extraction"),
        ("03", "Review", "Analyst verification"),
        ("04", "Validate", "Accounting-rule validation"),
        ("05", "Clean", "Controlled data cleaning"),
        ("06", "Analyse", "Dashboard & analytics"),
        ("07", "QA", "Quality & audit checks"),
        ("08", "Expand", "Population weighting & estimation"),
        ("09", "Report", "Sample or population outputs"),
    ]
    for col, (num, title, desc) in zip(workflow_cols, workflow):
        with col:
            st.markdown(
                f'<div class="app-card gold"><div style="font-size:.72rem;color:#1F7A6E;font-weight:800;">{num}</div>'
                f'<h4>{title}</h4><p>{desc}</p></div>',
                unsafe_allow_html=True,
            )

    st.markdown("### Current workflow status")
    status_df = pd.DataFrame({
        "Stage": [
            "PDF Upload & Extraction",
            "Analyst Review & Editing",
            "Accounting Rules Validation",
            "Data Cleaning",
            "Sampling Frame & Design",
            "Weighting & Population Estimation",
            "Dashboard & Analytics",
            "QA & Audit",
            "Reports & Export",
        ],
        "Status": [
            "Completed" if st.session_state["processed"] else "Pending",
            "Available" if st.session_state["processed"] else "Pending extraction",
            st.session_state.get("accounting_validation_status", "NOT_RUN").replace("ACCOUNTING RULES VALIDATED", "Validated").replace("OBJECTIONS RAISED", "Objections").replace("PENDING_REVALIDATION", "Revalidate").replace("NOT_RUN", "Pending"),
            "Completed" if st.session_state["cleaned"] else "Pending analyst cleaning",
            "Configured" if not st.session_state.get("sampling_frame",pd.DataFrame()).empty else "Pending frame",
            "Completed" if st.session_state.get("estimation_ready") else "Pending estimation",
            "Available" if st.session_state["processed"] else "Pending extraction",
            "Available" if st.session_state["processed"] else "Pending extraction",
            "Available" if st.session_state["processed"] else "Pending extraction",
        ]
    })
    st.dataframe(status_df, use_container_width=True, hide_index=True)

    st.info(
        "Raw extraction is never overwritten by cleaning. This separation is deliberate "
        "for official statistics, reproducibility and auditability."
    )

# ============================================================
# 1. UPLOAD & EXTRACT
# ============================================================

elif page == "ℹ️ About the Program":
    st.markdown("## About the Program")
    st.markdown(
        """
        <div class="app-hero" style="padding:1.25rem 1.5rem;">
            <div class="eyebrow">PUBLIC PROTOTYPE</div>
            <h1 style="font-size:1.9rem;">FPC Survey Intelligence</h1>
            <p>A controlled environment for Foreign Private Capital survey processing,
            statistical review, BPM6 mapping and BOP/IIP compiler support.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    about_sections = [
        "Program Overview",
        "Purpose of the Survey",
        "What the Program Does",
        "End-to-End Workflow",
        "Data Processing & Extraction",
        "Data Quality & Quality Assurance",
        "BPM6 Classification",
        "BOP Integration",
        "IIP Integration",
        "Security & Confidentiality",
        "User Guide",
    ]
    selected_about = st.selectbox(
        "Select an About the Program section",
        about_sections,
        index=0,
    )

    about_content = {
        "Program Overview": """
        <div class="about-box">
        <h3>Program Overview</h3>
        <p>The FPC Survey Intelligence & Population Estimation application is designed
        to turn completed Foreign Private Capital survey questionnaires into structured,
        reviewable and analytical datasets.</p>
        <p>The application preserves source-level information while providing controlled
        cleaning, quality assurance, statistical tables, analytical views and compiler-support
        outputs.</p>
        </div>
        """,
        "Purpose of the Survey": """
        <div class="about-box">
        <h3>Purpose of the Survey</h3>
        <p>The FPC survey captures information on foreign liabilities, foreign assets,
        equity, profits and dividends, non-equity instruments and services in a given period,  	which can either be monthly, quarterly or annually. The resulting
        information supports external-sector statistical compilation, including Balance
        of Payments and International Investment Position work.</p>
        <p>The application therefore treats extraction, classification, reconciliation and
        analyst review as separate stages rather than treating an extracted value as
        automatically verified.</p>
        </div>
        """,
        "What the Program Does": """
        <div class="about-box">
        <h3>What the Program Does</h3>
        <ul>
            <li>Processes multiple FPC questionnaire PDFs in a single batch.</li>
            <li>Uses native PDF extraction for typed forms.</li>
            <li>Uses OCR where scanned pages require it.</li>
            <li>Retains raw extraction values and semantic source states.</li>
            <li>Structures company, financial, industrial, shareholding, equity, profits,
                liabilities, assets, services and comments information.</li>
            <li>Provides analyst review and editing before cleaning.</li>
            <li>Builds data-quality and QA records.</li>
            <li>Provides BPM6 BOP/IIP compiler-support tables.</li>
            <li>Exports complete datasets and publication-oriented outputs.</li>
        </ul>
        </div>
        """,
        "End-to-End Workflow": """
        <div class="about-box">
        <h3>End-to-End Workflow</h3>
        <p><strong>Upload → Extract → Review → Clean → Analyse → QA → Report → Export</strong></p>
        <ol>
            <li><strong>Upload:</strong> demonstration users can upload synthetic or anonymised questionnaire PDFs.</li>
            <li><strong>Extract:</strong> the application reads native PDF text and invokes
                OCR for scanned pages when enabled.</li>
            <li><strong>Review:</strong> analysts inspect and, where necessary, edit extracted records.</li>
            <li><strong>Clean:</strong> controlled cleaning creates the analytical dataset.</li>
            <li><strong>Analyse:</strong> dashboards, company views, tables and analytical tools are available.</li>
            <li><strong>QA:</strong> reconciliation and quality flags are reviewed.</li>
            <li><strong>Report/Export:</strong> structured outputs can be downloaded in several formats.</li>
        </ol>
        </div>
        """,
        "Data Processing & Extraction": """
        <div class="about-box">
        <h3>Data Processing & Extraction</h3>
        <p>Native/typed PDFs are processed using PyMuPDF. Scanned pages can be processed using
        Tesseract OCR where it is installed and available to the application.</p>
        <p>OCR-derived information is explicitly retained with review indicators because
        scanned forms and handwriting can introduce recognition errors. The application does
        not treat OCR output as automatically verified statistical data.</p>
        </div>
        """,
        "Data Quality & Quality Assurance": """
        <div class="about-box">
        <h3>📋 Data Quality & Quality Assurance</h3>
        <p><strong>Ensuring accurate, consistent and reliable FPC survey data for official compilation.</strong></p>

        <p>The application separates extraction from cleaning and preserves the original
        extraction for auditability. Data Quality and Quality Assurance therefore go beyond
        checking whether fields are populated or correctly formatted. The system assesses
        whether related financial responses are internally coherent and whether the reported
        figures follow the accounting relationships implied by the questionnaire.</p>

        <hr>

        <h4>🎯 Objectives</h4>
        <ol>
            <li>Verify the <strong>completeness, plausibility and internal consistency</strong> of reported information.</li>
            <li>Check adherence to relevant <strong>accounting rules and financial relationships</strong>.</li>
            <li>Identify and flag potential <strong>errors, omissions, classification issues and unexplained differences</strong>.</li>
            <li>Provide compilers with clear diagnostics, affected values and recommended follow-up actions.</li>
            <li>Ensure that data are sufficiently reliable for subsequent <strong>BPM6, Balance of Payments and International Investment Position</strong> compilation.</li>
        </ol>

        <h4>🔎 What Data Quality and QA Check</h4>
        <ul>
            <li>Missing, blank, N/A and zero-like responses, while preserving the original source status.</li>
            <li>OCR and manual-review flags requiring verification against the source questionnaire.</li>
            <li>Duplicate records, implausible values and inconsistent classifications.</li>
            <li>Ownership and shareholding classifications and percentage consistency.</li>
            <li>Opening and closing balance reconciliations.</li>
            <li>Transaction, valuation/price, exchange-rate and other-volume movements where these are available.</li>
            <li>Cross-table consistency between financials, profits, equity, assets, liabilities and related questionnaire sections.</li>
            <li>Accounting identities and balance-sheet relationships.</li>
        </ul>

        <h4>🧮 Accounting Consistency Checks</h4>
        <p>Financial items are not assessed in isolation. Related responses are reconciled using
        accounting identities, opening and closing balances, reported transactions and other
        applicable changes.</p>

        <table style="width:100%; border-collapse:collapse; margin:0.7rem 0 1rem 0;">
          <tr>
            <th style="text-align:left; padding:8px; border-bottom:2px solid #1F4E79;">Check</th>
            <th style="text-align:left; padding:8px; border-bottom:2px solid #1F4E79;">Accounting relationship / example</th>
          </tr>
          <tr>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;"><strong>Profit and retained earnings</strong></td>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;">Net Profit After Tax − Dividends = Retained Profit/Earnings</td>
          </tr>
          <tr>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;"><strong>Profit before and after tax</strong></td>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;">Profit Before Tax − Tax Expense = Net Profit After Tax</td>
          </tr>
          <tr>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;"><strong>Retained earnings roll-forward</strong></td>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;">Opening Retained Earnings + Net Profit After Tax − Dividends ± Other Applicable Changes = Closing Retained Earnings</td>
          </tr>
          <tr>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;"><strong>Equity reconciliation</strong></td>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;">Opening Equity + Additional Equity − Equity Sold ± Applicable Changes = Closing Equity</td>
          </tr>
          <tr>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;"><strong>Share capital</strong></td>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;">Opening Share Capital + Capital Contributions − Reductions/Sales ± Other Applicable Movements = Closing Share Capital</td>
          </tr>
          <tr>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;"><strong>Reserves</strong></td>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;">Opening Reserves ± Additions, Transfers, Distributions, Revaluations, Exchange-Rate and Other Applicable Changes = Closing Reserves</td>
          </tr>
          <tr>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;"><strong>Assets and liabilities</strong></td>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;">Opening Position + Transactions ± Valuation/Price Changes ± Exchange-Rate Changes ± Other Volume Changes = Closing Position</td>
          </tr>
          <tr>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;"><strong>Balance sheet</strong></td>
            <td style="padding:8px; border-bottom:1px solid #D8E1EC;">Total Assets = Total Liabilities + Total Equity, where the required source fields are available</td>
          </tr>
          <tr>
            <td style="padding:8px;"><strong>Shareholding</strong></td>
            <td style="padding:8px;">Ownership percentages are checked for logical validity and applicable percentage constraints.</td>
          </tr>
        </table>

        <h4>🔗 Cross-Table Consistency</h4>
        <p>The QA process also compares information reported in different sections of the same
        questionnaire. For example, a profit reported in the <strong>Profits</strong> section
        may be compared with the amount used in the <strong>retained earnings/equity</strong>
        reconciliation. Similarly, opening and closing positions reported in financial tables
        are compared with the movements reported for the same instrument.</p>

        <h4>📊 QA Classification</h4>
        <ul>
            <li><strong>PASS</strong> – The accounting or data-quality relationship reconciles within the defined tolerance.</li>
            <li><strong>REVIEW</strong> – A difference exists but may be explainable by rounding, timing, classification, missing information or another documented adjustment.</li>
            <li><strong>FAIL / OBJECTION</strong> – A material inconsistency has been identified and requires investigation or questionnaire revision.</li>
            <li><strong>INFO / NOT TESTABLE</strong> – The relationship cannot be tested because the questionnaire does not contain all required source fields.</li>
        </ul>

        <div class="app-info">
        <strong>Important:</strong> A reconciliation difference does not automatically mean that
        the questionnaire is incorrect. It may arise from an omitted adjustment, classification
        issue, rounding difference, timing difference, legitimate valuation or exchange-rate
        movement, or an extraction/OCR error. The application therefore flags the difference
        for review rather than silently changing the respondent's reported value.
        </div>

        <h4>🛡️ Auditability and Compiler Review</h4>
        <p>The original extracted values remain available for comparison with reviewed and cleaned
        data. Accounting validation results record the company, accounting category, rule,
        reported value, expected value, difference, source area, objection and recommended action.</p>

        <p>Where an objection remains unresolved, the analyst may return the questionnaire for
        revision or explicitly authorise <strong>Proceed with Errors</strong>. Such an override
        does not remove the objection; it remains part of the QA and audit trail.</p>

        <h4>🔄 Quality Assurance Workflow</h4>
        <p><strong>Extract → Review & Edit → Accounting Validation → Data Cleaning → QA & Audit → Analysis → Reporting</strong></p>

        <p>This approach ensures that accounting consistency is considered before the information
        is transformed into analytical or compiler-support outputs.</p>
        </div>
        """,
        "BPM6 Classification": """
        <div class="about-box">
        <h3>BPM6 Classification</h3>
        <p>The compiler-support layer uses the BPM6 functional categories:</p>
        <ul>
            <li><strong>Direct investment</strong></li>
            <li><strong>Portfolio investment</strong></li>
            <li><strong>Financial derivatives and employee stock options</strong></li>
            <li><strong>Other investment</strong></li>
            <li><strong>Reserve assets</strong></li>
        </ul>
        <p>The FPC survey does not by itself provide complete coverage of every BOP/IIP
        component. The application's coverage and gap tables identify areas requiring
        additional external-sector datasets or compiler review.</p>
        </div>
        """,
        "BOP Integration": """
        <div class="about-box">
        <h3>BOP Integration</h3>
        <p>The BOP bridge provides compiler-support views of financial-account transactions
        available from the survey. For liabilities, the application distinguishes gross
        increases, gross decreases and derived net financial-account transactions where the
        source information permits.</p>
        <p>Derived values are labelled as derived rather than being presented as directly
        reported source values.</p>
        </div>
        """,
        "IIP Integration": """
        <div class="about-box">
        <h3>IIP Integration</h3>
        <p>The IIP bridge organises available opening positions, financial-account transactions,
        other changes and closing positions. A reconciliation signal is retained where the
        available components do not fully explain the closing position.</p>
        <p>For this reason, the FPC bridge should be viewed as a compiler-support component
        rather than as a complete replacement for a full external-sector database.</p>
        </div>
        """,
        "Security & Confidentiality": """
        <div class="about-box">
        <h3>Security & Confidentiality</h3>
        <p>FPC questionnaires contain enterprise-level financial and ownership information.
        Upload access in this application is therefore protected by authentication.</p>
        <p>Users should follow all applicable institutional information-security, confidentiality,
        records-management and official-statistics procedures. Authentication in this application
        is an application-level control and should not be considered a substitute for institutional
        access control, secure hosting or appropriate cybersecurity measures.</p>
        </div>
        """,
        "User Guide": """
        <div class="about-box">
        <h3>User Guide</h3>
        <ol>
            <li>Open <strong>1. Upload & Extract</strong>.</li>
                        <li>Upload one or more FPC questionnaire PDFs.</li>
            <li>Enable OCR where scanned pages need processing.</li>
            <li>Run <strong>Process FPC Surveys</strong>.</li>
            <li>Move to <strong>Review & Edit</strong> and verify extracted information.</li>
            <li>Run <strong>Validate Accounting Rules</strong> and review every objection.</li>
            <li>Revise questionnaire responses where required, or explicitly select <strong>Proceed with Errors</strong> when the analyst authorises continuation.</li>
            <li>Run <strong>Data Cleaning</strong>.</li>
            <li>Use the Dashboard, Company Explorer, Data Tables and Analytics Studio.</li>
            <li>Review the QA and BPM6 BOP/IIP views.</li>
            <li>Generate reports or export the final dataset.</li>
        </ol>
        </div>
        """,
    }
    st.markdown(about_content[selected_about], unsafe_allow_html=True)

elif page == "📚 Analytics Methodology":
    st.markdown("## Analytics Methodology")
    st.markdown("""
    <div class="app-hero" style="padding:1.15rem 1.4rem;">
        <div class="eyebrow">Statistical methodology and estimation framework</div>
        <h1 style="font-size:1.8rem;">From Sampling to Population Estimates</h1>
        <p>This section documents how FPC survey observations move from the sampling frame through sample selection, data collection, weighting, nonresponse adjustment, calibration and population estimation.</p>
    </div>
    """, unsafe_allow_html=True)
    method_sections=[
        "1. Purpose and Statistical Framework",
        "2. Target Population and Sampling Frame",
        "3. Sampling Design and Stratification",
        "4. Data Collection and Response",
        "5. Design Weights",
        "6. Nonresponse Adjustment",
        "7. Calibration and Benchmarking",
        "8. Population Estimation",
        "9. Sampling Variance and Precision",
        "10. Data Quality, Editing and Imputation",
        "11. BPM6 / BOP / IIP Linkage",
        "12. Worked End-to-End Example",
        "13. Limitations and Interpretation",
        "14. References",
    ]
    selected_method=st.selectbox("Select a methodology subsection",method_sections,index=0)
    method_html={
    "1. Purpose and Statistical Framework": """
    <div class="about-box"><h3>1. Purpose and Statistical Framework</h3>
    <p>The FPC survey is a source-data collection instrument for external-sector statistics. The estimation framework separates four questions: <b>who belongs to the population</b>, <b>who was selected</b>, <b>who responded</b>, and <b>how each response represents the population</b>. The IMF BPM6 Compilation Guide provides practical guidance on conducting surveys for balance of payments and IIP statistics and on expansion and poststratified estimation (IMF, 2014, Chapters 2–3 and 8).</p>
    <p>The general population-total estimator is:</p><p><b>Ŷ = Σ wᵢyᵢ</b></p><p>where yᵢ is the reported value for respondent i and wᵢ is the final survey weight. The application keeps the frame, linkage, weighting and estimation stages separate so that each step can be audited.</p></div>
    """,
    "2. Target Population and Sampling Frame": """
    <div class="about-box"><h3>2. Target Population and Sampling Frame</h3>
    <p>The target population is the set of enterprises represented by the supplied sampling frame. For production use, the frame should normally be based on the Statistical Business Register or another authoritative source and should contain a unique enterprise identifier, province, ISIC Rev.4 activity, size class and relevant auxiliary variables.</p>
    <p>If N<sub>h</sub> is the number of enterprises in stratum h, then the national frame size is:</p><p><b>N = Σ<sub>h</sub>N<sub>h</sub></b></p>
    <p>Frame quality matters because a perfectly weighted sample cannot correct for enterprises that are systematically missing from the frame.</p></div>
    """,
    "3. Sampling Design and Stratification": """
    <div class="about-box"><h3>3. Sampling Design and Stratification</h3>
    <p>The demonstration design uses <b>province × ISIC Rev.4 section × size band</b> as the principal stratification structure. Large enterprises can be treated as certainty or take-all units, while SMEs are selected probabilistically within strata. The actual survey should retain the official sample-selection probabilities.</p>
    <p>Under simple random sampling within a stratum:</p><p><b>πᵢ = n<sub>h</sub>/N<sub>h</sub></b></p>
    <p>Example: N<sub>h</sub>=200 and n<sub>h</sub>=40 gives πᵢ=0.20. Thus the selected unit initially represents 1/0.20=5 population units.</p>
    <p>IMF BPM6 Compilation Guide guidance covers survey design and sources for BOP/IIP compilation (IMF, 2014, Chapters 2–3).</p></div>
    """,
    "4. Data Collection and Response": """
    <div class="about-box"><h3>4. Data Collection and Response</h3>
    <p>Each selected enterprise should have a unique identifier so that selection, questionnaire receipt, editing and estimation remain linked. <b>Unit nonresponse</b> occurs when a selected enterprise does not provide a usable response. <b>Item nonresponse</b> occurs when a returned questionnaire has missing information for particular variables.</p>
    <p>Zero, blank, not applicable and unavailable must remain distinguishable. Treating missing values as zero can materially distort external-sector flows and positions. The application therefore records response/linkage separately from the frame itself.</p></div>
    """,
    "5. Design Weights": """
    <div class="about-box"><h3>5. Design Weights</h3>
    <p>The design weight is the inverse of the inclusion probability:</p><p><b>dᵢ = 1/πᵢ</b></p>
    <p>For simple random sampling, this becomes:</p><p><b>dᵢ = N<sub>h</sub>/n<sub>h</sub></b></p>
    <p>Example: if πᵢ=0.05, then dᵢ=20. The respondent represents 20 population units before response and calibration adjustments.</p>
    <p>For take-all units, πᵢ=1 and dᵢ=1.</p></div>
    """,
    "6. Nonresponse Adjustment": """
    <div class="about-box"><h3>6. Nonresponse Adjustment</h3>
    <p>Within-stratum nonresponse adjustment can be expressed as:</p><p><b>A<sub>h</sub> = n<sub>h</sub>/r<sub>h</sub></b></p><p>and the adjusted respondent weight as:</p><p><b>wᵢ = dᵢ × A<sub>h</sub></b></p>
    <p>Example: 40 enterprises selected, 32 responding → A<sub>h</sub>=40/32=1.25. A design weight of 5 becomes 6.25.</p>
    <p>If r<sub>h</sub>=0, the adjustment is undefined. The stratum should be flagged for review rather than assigning an arbitrary weight. The BPM6 Compilation Guide discusses expansion factors for nonresponse and poststratified estimation (IMF, 2014, para. 8.24).</p></div>
    """,
    "7. Calibration and Benchmarking": """
    <div class="about-box"><h3>7. Calibration and Benchmarking</h3>
    <p>Calibration uses trusted auxiliary population totals to adjust weights. For one benchmark variable X:</p><p><b>g = X<sub>pop</sub> / ΣwᵢXᵢ</b></p><p><b>wᵢ* = wᵢ × g</b></p>
    <p>Example: weighted employment is 92,000 while the frame benchmark is 100,000. Then g=100,000/92,000=1.08696. A weight of 6.25 becomes 6.7935.</p>
    <p>For several benchmarks, raking or generalized regression calibration may be preferable. The benchmark source and calibration constraints must be documented.</p></div>
    """,
    "8. Population Estimation": """
    <div class="about-box"><h3>8. Population Estimation</h3>
    <p>For a total:</p><p><b>Ŷ = Σ<sub>i∈sᵣ</sub>wᵢyᵢ</b></p>
    <p>For a weighted mean:</p><p><b>ȳ<sub>w</sub> = Σwᵢyᵢ / Σwᵢ</b></p>
    <p>For a weighted proportion:</p><p><b>p̂ = ΣwᵢI(yᵢ=1) / Σwᵢ</b></p>
    <p>Domain estimates use the same formulas after restricting the summation to the relevant province, ISIC section, size class or other domain.</p>
    <p>Example: three enterprises report US$10m, US$20m and US$30m with weights 5, 8 and 12. Estimated turnover = 5×10 + 8×20 + 12×30 = <b>US$570m</b>. Weighted number of enterprises = 25, so weighted mean turnover = <b>US$22.8m</b>.</p></div>
    """,
    "9. Sampling Variance and Precision": """
    <div class="about-box"><h3>9. Sampling Variance and Precision</h3>
    <p>For a simple stratified design, an approximate variance for a total is:</p><p><b>Var(Ŷ) ≈ Σ<sub>h</sub>N<sub>h</sub>²(1−f<sub>h</sub>)s<sub>h</sub>²/n<sub>h</sub></b></p>
    <p>where f<sub>h</sub>=n<sub>h</sub>/N<sub>h</sub>. Then:</p><p><b>SE(Ŷ)=√Var(Ŷ)</b></p><p><b>95% CI ≈ Ŷ ± 1.96×SE(Ŷ)</b></p><p><b>CV = 100×SE(Ŷ)/|Ŷ|</b></p>
    <p>These are approximate design-based measures. If the real survey is clustered, multistage or uses unequal-probability selection, a design-specific variance estimator should be used.</p></div>
    """,
    "10. Data Quality, Editing and Imputation": """
    <div class="about-box"><h3>10. Data Quality, Editing and Imputation</h3>
    <p>Quality assurance includes range checks, logical checks, duplicate checks, unit checks, accounting identities, period-to-period comparisons and review of unusual observations. The source value, cleaned value, derivation status and reviewer action should remain auditable.</p>
    <p>Item nonresponse may require imputation. Any imputed value should be flagged and its method documented. Unit nonresponse should not simply be converted to zero.</p></div>
    """,
    "11. BPM6 / BOP / IIP Linkage": """
    <div class="about-box"><h3>11. BPM6 / BOP / IIP Linkage</h3>
    <p>BPM6 defines an immediate direct investment relationship when a nonresident investor owns equity giving <b>10 percent or more of voting power</b>; more than 50 percent indicates control and 10–50 percent indicates significant influence (IMF, 2009, paras. 6.8–6.12).</p>
    <p>Direct investment includes the equity associated with the relationship and relevant debt positions, subject to BPM6 exceptions. Portfolio investment covers securities not included in direct investment or reserve assets. Other investment includes loans, trade credit and advances, insurance and pension entitlements and other accounts (IMF, 2009, Chapter 6).</p>
    <p>The supplied Jamaica IIP report demonstrates the publication structure: it presents an IIP summary, then discusses portfolio investment, FDI, other investment and reserve assets, and includes a BPM6-based glossary (Bank of Jamaica, 2014, pp. 2–8).</p></div>
    """,
    "12. Worked End-to-End Example": """
    <div class="about-box"><h3>12. Worked End-to-End Example</h3>
    <p>Suppose a stratum has N=200 enterprises. Forty are sampled and 32 respond. A respondent reports turnover of US$20m.</p>
    <ol><li>Selection probability: π=40/200=0.20.</li><li>Design weight: d=1/0.20=5.</li><li>Nonresponse adjustment: A=40/32=1.25.</li><li>Final weight: w=5×1.25=6.25.</li><li>Contribution to estimated turnover: 6.25×US$20m=US$125m.</li></ol>
    <p>If all 32 respondents were processed, their weighted contributions would be summed to obtain the stratum estimate. Stratum estimates are then summed across strata to obtain the national estimate, subject to any calibration or benchmarking procedure.</p></div>
    """,
    "13. Limitations and Interpretation": """
    <div class="about-box"><h3>13. Limitations and Interpretation</h3>
    <p>Population estimates depend on frame coverage, correct inclusion probabilities, response patterns, data quality and valid calibration benchmarks. Weighting does not automatically remove all nonsampling errors. The synthetic 2,500-establishment frame supplied with this application is a demonstration frame and must not be interpreted as an official Zimbabwean FPC population.</p>
    <p>Zero-response strata, missing frame identifiers, uncertain linkage, excessive weights and weak benchmark coverage should be treated as estimation risks and reported through the QA layer.</p></div>
    """,
    "14. References": """
    <div class="about-box"><h3>14. References</h3>
    <p><b>[1]</b> International Monetary Fund (2009), <i>Balance of Payments and International Investment Position Manual, Sixth Edition (BPM6)</i>, Washington, DC. Chapters 6–7.</p>
    <p><b>[2]</b> International Monetary Fund (2014), <i>Balance of Payments and International Investment Position Manual, Sixth Edition: Compilation Guide</i>, Washington, DC. Chapters 2–3 and 8.</p>
    <p><b>[3]</b> Bank of Jamaica (2014), <i>Quarterly International Investment Position Report — December 2014 Quarter</i>, International Accounts Unit, Economic Information Publications Department.</p>
    <p><b>[4]</b> International Monetary Fund, <i>Coordinated Direct Investment Survey Guide</i> and related BPM6 direct-investment guidance.</p></div>
    """,
    }
    st.markdown(method_html[selected_method], unsafe_allow_html=True)

elif page == "✍️ The Author":
    st.markdown("## The Author")
    st.markdown(
        """
        <div class="app-hero">
            <div class="eyebrow">Application Development</div>
            <h1 style="font-size:2rem;">FPC Survey Intelligence</h1>
            <p>Designed and developed as a quantitative statistical intelligence and
            data-processing application for Foreign Private Capital survey workflows.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="about-box">
        <h3>Author & Developer</h3>
        <p><strong>Admire Tarisirayi Chirume</strong></p>
        <p>This application brings together statistical data extraction, data engineering,
        quantitative analysis, external-sector concepts and software development into a
        controlled workflow for FPC survey processing.</p>
        <p>The application is intended to support professional statistical compilation,
        analytical review and evidence-based decision-making. Source data, classifications,
        derived values and compiler judgements should always be reviewed in accordance with
        applicable institutional procedures.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

elif page == "©️ Copyright":
    st.markdown("## Copyright & Use")
    st.markdown(
        """
        <div class="app-hero">
            <div class="eyebrow">PUBLIC PROTOTYPE</div>
            <h1 style="font-size:2rem;">FPC Survey Intelligence</h1>
            <p>Foreign Private Capital • BOP • IIP • BPM6</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="about-box">
        <h3>Copyright</h3>
        <p><strong>© 2026 Admire Tarisirayi Chirume</strong></p>
        <p>FPC Survey Intelligence & Population Estimation, including its interface,
        analytical workflow, data-processing architecture and associated documentation,
        is protected intellectual work.</p>
        <p>Unauthorised reproduction, redistribution, modification or commercial exploitation
        should not be undertaken without the appropriate institutional or rights-holder
        authorisation.</p>
        <p style="color:#667085;font-size:.86rem;">
        No government or institutional logo, endorsement or official status is implied by this prototype.
        </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

elif page == "📥 1. Upload & Extract":
    st.subheader("1. FPC Survey Upload & Extraction")
    st.markdown(
        """
        <div class="app-security">
        <strong>🧪 Public Prototype Upload Area</strong><br>
        This demonstration application is publicly accessible. Do not upload confidential,
        personal, commercially sensitive or production survey data. Use synthetic or
        appropriately anonymised test questionnaires for demonstration purposes.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("### Upload FPC Survey PDFs")
    st.write(
        "Upload one or more completed questionnaires. The extraction engine retains "
        "section-specific values, raw text, OCR flags and QA records."
    )

    files = st.file_uploader(
        "Select completed FPC questionnaire PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        help="Multiple company questionnaires can be processed in one batch.",
    )

    c1, c2 = st.columns([1, 1])
    with c1:
        ocr_enabled = st.checkbox("Use OCR for scanned pages", value=True)
    with c2:
        if TESSERACT_PATH:
            st.success("Tesseract available")
        else:
            st.warning("Tesseract not installed — scanned pages may remain blank.")

    if files:
        preview = pd.DataFrame({
            "File": [f.name for f in files],
            "Size (MB)": [round(len(f.getvalue()) / 1024**2, 2) for f in files],
        })
        st.success(f"{len(files)} PDF file(s) selected.")
        st.dataframe(preview, use_container_width=True, hide_index=True)

    if st.button("🚀 Process FPC Surveys", type="primary", use_container_width=True):
        if not files:
            st.error("Please upload at least one PDF.")
        else:
            results = []
            progress = st.progress(0)
            status = st.empty()
            errors = []

            for i, f in enumerate(files, start=1):
                status.info(f"Processing {i}/{len(files)}: {f.name}")
                try:
                    result = process_uploaded_pdf(f, ocr_if_needed=ocr_enabled)
                    results.append(result)
                except Exception as exc:
                    errors.append({"file": f.name, "error": str(exc)})
                progress.progress(i / len(files))

            if errors:
                st.error("Some files could not be processed.")
                st.dataframe(pd.DataFrame(errors), use_container_width=True, hide_index=True)

            if results:
                data = consolidate(results)
                _apply_tables(data, records=results)
                st.session_state["records"] = results
                st.session_state["processed"] = True
                st.session_state["cleaned"] = False
                st.session_state["cleaned_tables"] = {}
                st.session_state["cleaning_log"] = pd.DataFrame()
                st.session_state["accounting_validation"] = pd.DataFrame()
                st.session_state["accounting_validation_status"] = "NOT_RUN"
                st.session_state["accounting_validation_run"] = None
                st.session_state["accounting_validation_override"] = False
                st.session_state["last_run"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                st.session_state["dataset_id"] = (
                    "FPC-" + datetime.now().strftime("%Y%m%d-%H%M%S")
                )

                saved, save_error = save_persisted_state()

                st.success(
                    f"Completed: {len(results)} questionnaire(s) processed. "
                    f"{dataset_company_count()} companies are available."
                )
                if saved:
                    st.success("The complete dataset has been saved as the canonical V3 dataset.")
                else:
                    st.warning(f"Dataset processed but local save failed: {save_error}")

                st.markdown("### Extraction-stage downloads")
                tables = active_tables()
                workbook = full_stage_excel_bytes(tables, "Extraction")
                st.download_button(
                    "⬇️ Download Extraction Workbook (Excel)",
                    data=workbook,
                    file_name=f"FPC_Extraction_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

                st.session_state["_fpc_pending_page"] = "🔎 2. Review & Edit"
                st.rerun()


# ============================================================
# 2. REVIEW & EDIT
# ============================================================

elif page == "🔎 2. Review & Edit":
    st.subheader("2. Review & Edit Extracted Data")

    if not st.session_state["processed"]:
        st.info("Process PDFs first using **1. Upload & Extract**.")
    else:
        st.success(
            f"{company_count(st.session_state['company_master'])} companies available for analyst review."
        )
        tabs = st.tabs([
            "Company Master", "Industrial", "Shareholding", "Equity", "Profits",
            "Foreign Asset Profits", "Non-Equity Liabilities", "Non-Equity Assets",
            "Services", "Comments"
        ])
        names = [
            "company_master", "industrial", "shareholding", "equity", "profits",
            "asset_profits", "non_equity_liabilities", "non_equity_assets",
            "services", "comments"
        ]

        for tab, key in zip(tabs, names):
            with tab:
                df = st.session_state.get(key, pd.DataFrame())
                if not isinstance(df, pd.DataFrame):
                    df = pd.DataFrame()
                st.caption(f"{len(df)} rows × {len(df.columns)} columns")
                if df.empty:
                    st.info("No records extracted for this section.")
                else:
                    edited = st.data_editor(
                        df,
                        use_container_width=True,
                        num_rows="dynamic",
                        key=f"review_editor_v3_{key}",
                    )
                    if isinstance(edited, pd.DataFrame):
                        # st.data_editor() returns its current dataframe on every
                        # Streamlit rerun.  Do NOT invalidate Accounting Validation
                        # merely because the page reran (for example after clicking
                        # "Validate Accounting Rules" or "Proceed with Errors").
                        # Invalidate only when the analyst has actually changed
                        # the data in this editor.
                        original_df = st.session_state.get(key, pd.DataFrame())
                        data_changed = (
                            isinstance(original_df, pd.DataFrame)
                            and not edited.equals(original_df)
                        )
                        st.session_state[key] = edited.copy()

                        if data_changed:
                            st.session_state["accounting_validation_status"] = "PENDING_REVALIDATION"
                            st.session_state["accounting_validation_override"] = False
                            st.session_state["accounting_validation"] = pd.DataFrame()

                    c1, c2 = st.columns(2)
                    with c1:
                        if st.button(
                            f"💾 Save {key.replace('_',' ').title()}",
                            key=f"save_review_{key}",
                            use_container_width=True
                        ):
                            st.session_state["processed"] = True
                            st.session_state["cleaned"] = False
                            st.session_state["accounting_validation_status"] = "PENDING_REVALIDATION"
                            st.session_state["accounting_validation_override"] = False
                            st.session_state["cleaned_tables"] = {}
                            save_persisted_state()
                            st.success("Saved.")
                    with c2:
                        st.download_button(
                            f"⬇️ Download {key.replace('_',' ').title()} (Excel)",
                            data=table_excel_bytes(
                                st.session_state[key], key
                            ),
                            file_name=f"{safe_filename(key)}_reviewed.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_review_{key}",
                            use_container_width=True,
                        )

        st.divider()
        st.markdown("### 🧾 Accounting Rules Validation")
        st.markdown(
            "Run this validation **after reviewing and editing every questionnaire response**. "
            "The engine checks available accounting identities without inventing missing values. "
            "Objections identify the company, rule, reported value, expected value and the questionnaire section that should be revised."
        )

        av_status = st.session_state.get("accounting_validation_status", "NOT_RUN")
        av_df = st.session_state.get("accounting_validation", pd.DataFrame())
        av_override = bool(st.session_state.get("accounting_validation_override", False))

        if av_status in {"NOT_RUN", "PENDING_REVALIDATION"}:
            if av_status == "PENDING_REVALIDATION":
                st.warning("The reviewed dataset has changed since the last Accounting Validation. Validation must be run again before cleaning.")
            if st.button("🧮 Validate Accounting Rules", type="primary", use_container_width=True):
                validation_df, validation_summary = validate_accounting_rules(active_tables(), tolerance=ACCOUNTING_TOLERANCE)
                st.session_state["accounting_validation"] = validation_df
                st.session_state["accounting_validation_status"] = validation_summary["status"]
                st.session_state["accounting_validation_run"] = validation_summary["validated_at"]
                st.session_state["accounting_validation_override"] = False
                save_persisted_state()
                st.rerun()

        elif av_status == "ACCOUNTING RULES VALIDATED":
            st.success(
                f"✅ Accounting rules validated — {int((av_df.get('status', pd.Series(dtype=object)) == 'PASS').sum())} checks passed; "
                f"{int((av_df.get('status', pd.Series(dtype=object)) == 'NOT TESTABLE').sum())} checks were not testable because the questionnaire does not contain the required source fields."
            )
            if st.session_state.get("accounting_validation_run"):
                st.caption(f"Validation run: {st.session_state['accounting_validation_run']}")
            if isinstance(av_df, pd.DataFrame) and not av_df.empty:
                c1, c2, c3 = st.columns(3)
                c1.metric("Passed", int((av_df["status"] == "PASS").sum()))
                c2.metric("Objections", int((av_df["status"] == "OBJECTION").sum()))
                c3.metric("Not testable", int((av_df["status"] == "NOT TESTABLE").sum()))
                st.dataframe(av_df, use_container_width=True, hide_index=True)
                st.download_button(
                    "⬇️ Download Accounting Validation Report (Excel)",
                    data=accounting_validation_excel_bytes(av_df),
                    file_name=f"FPC_Accounting_Validation_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            st.info("All available accounting rules passed. You may now proceed to Data Cleaning.")
            if st.button("➡️ Proceed to Data Cleaning", type="primary", use_container_width=True):
                st.session_state["_fpc_pending_page"] = "🧹 3. Data Cleaning"
                st.rerun()

        elif av_status == "OBJECTIONS RAISED":
            objections = av_df[av_df["status"] == "OBJECTION"].copy() if isinstance(av_df, pd.DataFrame) else pd.DataFrame()
            st.error(
                f"⚠️ Accounting validation raised {len(objections)} objection(s). "
                "The questionnaire responses should be revised where indicated, or the analyst may explicitly proceed with errors."
            )
            if isinstance(objections, pd.DataFrame) and not objections.empty:
                display_cols = [c for c in ["company_name", "category", "rule_id", "check", "reported_value", "expected_value", "difference", "source", "objection", "recommendation"] if c in objections.columns]
                st.dataframe(objections[display_cols], use_container_width=True, hide_index=True)

            c1, c2 = st.columns(2)
            with c1:
                if st.button("↩️ Return to Review & Revise", use_container_width=True):
                    st.session_state["accounting_validation_status"] = "PENDING_REVALIDATION"
                    st.session_state["accounting_validation_override"] = False
                    save_persisted_state()
                    st.rerun()
            with c2:
                if st.button("⚠️ Proceed with Errors", use_container_width=True):
                    st.session_state["accounting_validation_override"] = True
                    save_persisted_state()
                    st.warning("Proceed-with-errors authorisation recorded. The next stage can continue, but the accounting objections remain in the audit trail.")

            if st.session_state.get("accounting_validation_override"):
                st.warning("Proceed-with-errors has been selected. The objections will remain visible in the Accounting Validation Report and audit trail.")
                if st.button("➡️ Proceed to Data Cleaning with Errors", type="primary", use_container_width=True):
                    st.session_state["_fpc_pending_page"] = "🧹 3. Data Cleaning"
                    st.rerun()

        if isinstance(av_df, pd.DataFrame) and not av_df.empty and av_status not in {"ACCOUNTING RULES VALIDATED", "OBJECTIONS RAISED"}:
            st.dataframe(av_df, use_container_width=True, hide_index=True)

        st.divider()
        st.markdown("### Download complete reviewed dataset")
        reviewed_workbook = full_stage_excel_bytes(active_tables(), "Reviewed")
        st.download_button(
            "⬇️ Download Complete Reviewed Dataset (Excel)",
            data=reviewed_workbook,
            file_name=f"FPC_Reviewed_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )


# ============================================================
# 3. DATA CLEANING
# ============================================================

elif page == "🧹 3. Data Cleaning":
    st.subheader("3. Data Cleaning & Preparation")

    if not st.session_state["processed"]:
        st.info("Complete extraction and review first.")
    elif st.session_state.get("accounting_validation_status") not in {"ACCOUNTING RULES VALIDATED", "OBJECTIONS RAISED"}:
        st.warning("🧾 Accounting Rules Validation is required after Review & Edit and before Data Cleaning.")
        st.info("Return to **2. Review & Edit**, run **Validate Accounting Rules**, resolve any objections or explicitly select **Proceed with Errors**, then continue here.")
    elif st.session_state.get("accounting_validation_status") == "OBJECTIONS RAISED" and not st.session_state.get("accounting_validation_override", False):
        st.error("Accounting objections are still unresolved. Revise the questionnaire responses or explicitly choose **Proceed with Errors** in Review & Edit.")
    else:
        if st.session_state.get("accounting_validation_override"):
            st.warning("⚠️ Data Cleaning is proceeding under an explicit **Proceed with Errors** override. Accounting objections remain in the audit trail.")
        st.markdown("""
        **Raw extracted data remain unchanged.** Cleaning creates a separate analytical
        dataset. You can inspect missingness, duplicates, numeric conversion and then
        download the cleaned stage.
        """)

        c1, c2, c3, c4 = st.columns(4)
        trim = c1.checkbox("Trim text", True)
        missing = c2.checkbox("Standardise missing values", True)
        numeric = c3.checkbox("Convert numeric fields", True)
        duplicates = c4.checkbox("Remove exact duplicates", True)

        if st.button("🧹 Run Safe Cleaning", type="primary", use_container_width=True):
            clean_all_tables({
                "trim": trim,
                "missing": missing,
                "numeric": numeric,
                "duplicates": duplicates
            })
            st.success("Cleaning completed. Raw extraction tables remain preserved.")
            st.rerun()

        if st.session_state.get("cleaned"):
            st.success(
                f"Cleaned analytical dataset available — {st.session_state.get('last_cleaned','')}"
            )

            if not st.session_state["cleaning_log"].empty:
                st.markdown("### Cleaning audit log")
                st.dataframe(
                    st.session_state["cleaning_log"],
                    use_container_width=True,
                    hide_index=True
                )

            keys = list(TABLE_KEYS)
            selected = st.selectbox("Inspect cleaned table", keys)
            raw_df = st.session_state.get(selected, pd.DataFrame())
            clean_df = active_table(selected)

            c1, c2 = st.columns(2)
            c1.metric("Raw rows", len(raw_df))
            c2.metric("Clean rows", len(clean_df))

            st.markdown("### Cleaned table preview")
            st.dataframe(clean_df, use_container_width=True, hide_index=True)

            st.download_button(
                "⬇️ Download Cleaned Table (Excel)",
                data=table_excel_bytes(clean_df, selected),
                file_name=f"{safe_filename(selected)}_cleaned.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.download_button(
                "⬇️ Download Cleaned Table (CSV)",
                data=clean_df.to_csv(index=False).encode("utf-8"),
                file_name=f"{safe_filename(selected)}_cleaned.csv",
                mime="text/csv",
            )

            cleaned_workbook = full_stage_excel_bytes(active_tables(), "Cleaned")
            st.download_button(
                "⬇️ Download Complete Cleaned Dataset (Excel)",
                data=cleaned_workbook,
                file_name=f"FPC_Cleaned_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )


# ============================================================
# 4. DASHBOARD
# ============================================================


elif page == "📐 4. Sampling Frame & Design":
    st.subheader("4. Sampling Frame & Survey Design")
    st.markdown("""<div class="app-info"><strong>Population estimation begins with the sampling frame.</strong><br>Upload the Statistical Business Register or the final survey sampling frame used to select the enterprises. The frame defines the target population to which the responding sample can be expanded.</div>""",unsafe_allow_html=True)
    up=st.file_uploader("Upload sampling frame (CSV or Excel)",type=["csv","xlsx","xls","xlsm"],key="sampling_frame_uploader")
    if up is not None:
        try:
            _upload_name=str(getattr(up,"name","")).lower()
            _sheet=None
            if _upload_name.endswith((".xlsx",".xls",".xlsm")):
                _xl=pd.ExcelFile(up)
                _preferred=[x for x in ["sampling_frame","frame","sample_frame","sbr","business_register"] if x in _xl.sheet_names]
                _default_sheet=_preferred[0] if _preferred else _xl.sheet_names[0]
                _sheet=st.selectbox("Sampling-frame worksheet",_xl.sheet_names,index=_xl.sheet_names.index(_default_sheet),key="sampling_frame_sheet")
            frame=_read_sampling_frame(up,_sheet)
            st.session_state["sampling_frame"]=_clean_frame(frame)
            st.success(f"Sampling frame loaded: {len(frame):,} units and {len(frame.columns):,} fields.")
        except Exception as exc:
            st.error(f"Could not read sampling frame: {exc}")
    frame=st.session_state.get("sampling_frame",pd.DataFrame())
    if frame.empty:
        st.info("Minimum recommended frame fields: enterprise identifier, enterprise name, province, ISIC Rev.4, size class, sampled/responded status and, where available, selection probability or benchmark variables.")
        st.markdown("### Recommended sampling-frame structure")
        st.dataframe(pd.DataFrame({"Field":["enterprise_id","enterprise_name","province","isic_rev4","size_class","sampled","responded","selection_probability","employment","turnover"],"Purpose":["Unique unit identifier","Matching to survey respondent","Geographic domain","Industry domain","Size stratum/domain","Sample selection status","Response status","Design probability","Optional calibration benchmark","Optional calibration benchmark"]}),use_container_width=True,hide_index=True)
    else:
        cols=list(frame.columns)
        cfg=st.session_state.get("estimation_config",{})
        _frame_default = "enterprise_name" if "enterprise_name" in cols else ("enterprise_id" if "enterprise_id" in cols else cols[0])
        frame_key=st.selectbox("Frame enterprise identifier",cols,index=cols.index(cfg.get("frame_key",_frame_default)) if cfg.get("frame_key",_frame_default) in cols else 0)
        survey_options=["company_name"]+[c for c in st.session_state.get("bop_iip_summary",pd.DataFrame()).columns if c not in {"company_name"}]
        survey_key=st.selectbox("Survey identifier",survey_options,index=survey_options.index(cfg.get("survey_key","company_name")) if cfg.get("survey_key","company_name") in survey_options else 0)
        flag_options=["(derive from matched respondents)"]+cols
        sampled_col=st.selectbox("Sampled-status field (optional)",flag_options,index=(flag_options.index(cfg.get("sampled_col")) if cfg.get("sampled_col") in flag_options else 0))
        responded_col=st.selectbox("Responded-status field (optional)",flag_options,index=(flag_options.index(cfg.get("responded_col")) if cfg.get("responded_col") in flag_options else 0))
        response_mode=st.selectbox("Response determination",["Survey linkage","Frame response field"],index=1 if cfg.get("response_mode")=="Frame response field" else 0,help="Survey linkage is recommended when the uploaded frame contains simulated/planned response flags or when response status should be determined from the actual uploaded survey records.")
        pi_options=["(not supplied)"]+cols
        pi_col=st.selectbox("Selection probability field (optional)",pi_options,index=(pi_options.index(cfg.get("selection_probability_col")) if cfg.get("selection_probability_col") in pi_options else 0))
        default_strata=[c for c in cfg.get("stratum_cols",[]) if c in cols]
        stratum_cols=st.multiselect("Estimation strata",cols,default=default_strata,help="Typical design variables are province × ISIC Rev.4 × size class. Use the actual sample design, not an arbitrary cross-classification.")
        domain_candidates=[c for c in cols if c not in {frame_key}]
        domain_cols=st.multiselect("Population reporting domains",domain_candidates,default=[c for c in cfg.get("domain_cols",[]) if c in domain_candidates])
        numeric_cols=[c for c in cols if pd.api.types.is_numeric_dtype(frame[c])]
        benchmark_options=["(none)"]+numeric_cols
        benchmark_col=st.selectbox("Optional calibration benchmark",benchmark_options,index=(benchmark_options.index(cfg.get("benchmark_col")) if cfg.get("benchmark_col") in benchmark_options else 0),help="The selected frame variable is used as a known population total for transparent one-variable ratio calibration.")
        if st.button("💾 Save Sampling Design",type="primary",use_container_width=True):
            st.session_state["estimation_config"]={
                "frame_key":frame_key,"survey_key":survey_key,
                "sampled_col":None if sampled_col.startswith("(") else sampled_col,
                "responded_col":None if responded_col.startswith("(") else responded_col,
                "response_mode":response_mode,
                "selection_probability_col":None if pi_col.startswith("(") else pi_col,
                "stratum_cols":stratum_cols,"domain_cols":domain_cols,
                "benchmark_col":None if benchmark_col.startswith("(") else benchmark_col,
            }
            save_persisted_state(); st.success("Sampling design saved. Proceed to Weighting & Estimation.")
        st.markdown("### Frame profile")
        p1,p2,p3=st.columns(3); p1.metric("Frame units",f"{len(frame):,}"); p2.metric("Fields",f"{len(frame.columns):,}"); p3.metric("Duplicate frame IDs",f"{frame[frame_key].duplicated().sum():,}")
        if stratum_cols:
            st.dataframe(frame.groupby(stratum_cols,dropna=False).size().reset_index(name="population_N").head(100),use_container_width=True,hide_index=True)
        st.download_button("⬇️ Sampling Frame Profile — Excel",table_excel_bytes(frame,"Sampling_Frame"),"FPC_Sampling_Frame_Profile.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif page == "⚖️ 5. Weighting & Estimation":
    st.subheader("5. Weighting, Nonresponse Adjustment & Calibration")
    frame=st.session_state.get("sampling_frame",pd.DataFrame()); cfg=st.session_state.get("estimation_config",{})
    if frame.empty or not cfg.get("frame_key"):
        st.info("Upload and configure the sampling frame first.")
    elif not st.session_state.get("processed"):
        st.info("Process the FPC survey responses before running population estimation.")
    else:
        st.markdown("""<div class="app-info"><strong>Weighting sequence</strong><br>1. Base design weight = 1/selection probability where supplied, otherwise N<sub>h</sub>/n<sub>h</sub>.<br>2. Within-stratum nonresponse adjustment = n<sub>sampled,h</sub>/n<sub>responded,h</sub>.<br>3. Optional one-variable ratio calibration to a known frame benchmark.<br>4. Weighted totals and approximate precision are calculated from the resulting respondent weights.</div>""",unsafe_allow_html=True)
        if st.button("▶️ Run Population Estimation",type="primary",use_container_width=True):
            with st.spinner("Building estimation microdata, weights, population totals and precision measures..."):
                ok,err=run_population_estimation(active_tables(),frame,cfg)
            if ok: st.success("Population estimation completed and the results were saved.")
            else: st.error(err)
        if st.session_state.get("estimation_ready"):
            meta=st.session_state.get("estimation_metadata",{})
            c=st.columns(5); c[0].metric("Frame",f"{meta.get('frame_units',0):,}"); c[1].metric("Sampled",f"{meta.get('sampled_units',0):,}"); c[2].metric("Respondents",f"{meta.get('respondent_units',0):,}"); c[3].metric("Strata",f"{meta.get('strata',0):,}"); c[4].metric("Run",str(meta.get('run_time',''))[:10])
            qa=st.session_state.get("estimation_qa",pd.DataFrame())
            if not qa.empty:
                zero=int(qa["zero_response_flag"].sum()) if "zero_response_flag" in qa.columns else 0
                if zero: st.warning(f"{zero} estimation stratum(s) have zero respondents. Those strata require design review before national estimates are treated as complete.")
                st.dataframe(qa,use_container_width=True,hide_index=True)
            meta=st.session_state.get("estimation_metadata",{})
            st.info(f"Survey linkage: {meta.get('matched_survey_units',0):,} frame units matched to survey records; {meta.get('unmatched_frame_units',0):,} frame units unmatched. Response determination: {meta.get('response_detection','Survey linkage')}.")
            link=st.session_state.get("estimation_linkage",pd.DataFrame())
            if not link.empty:
                with st.expander("View frame-to-survey linkage diagnostics"):
                    st.dataframe(link,use_container_width=True,hide_index=True)
            st.download_button("⬇️ Download Estimation Weights — Excel",table_excel_bytes(st.session_state.get("estimation_microdata",pd.DataFrame()),"Estimation_Weights"),"FPC_Estimation_Weights.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif page == "🌍 6. Population Estimates":
    st.subheader("6. Population Estimates & National Expansion")
    if not st.session_state.get("estimation_ready"):
        st.info("Run the population estimation engine first.")
    else:
        meta=st.session_state.get("estimation_metadata",{})
        st.markdown("""<div class="app-info"><strong>Interpretation</strong><br>These are weighted estimates for the target population represented by the supplied sampling frame. They are not simple multiplications of the sample totals. The report retains the design, nonresponse, calibration and precision assumptions used to obtain them.</div>""",unsafe_allow_html=True)
        est=st.session_state.get("population_estimates",pd.DataFrame())
        if not est.empty:
            st.dataframe(est,use_container_width=True,hide_index=True)
            st.markdown("### National estimates — visualisation")
            chart=est.dropna(subset=["estimate"]).copy()
            if not chart.empty:
                fig=px.bar(chart,x="label",y="estimate",title="Estimated National FPC Indicators")
                fig.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig,use_container_width=True)
        br=st.session_state.get("population_breakdowns",pd.DataFrame())
        if not br.empty:
            st.markdown("### Domain estimates")
            domain=st.selectbox("Select reporting domain",sorted(br["domain_type"].dropna().unique().tolist()))
            view=br[br["domain_type"]==domain]
            st.dataframe(view,use_container_width=True,hide_index=True)
        st.markdown("### Estimation assumptions")
        st.json({k:v for k,v in meta.items()})
        st.download_button("⬇️ Population Estimates — CSV",est.to_csv(index=False).encode("utf-8"),"FPC_Population_Estimates.csv","text/csv",type="primary")
        st.download_button("⬇️ Population Domain Estimates — CSV",br.to_csv(index=False).encode("utf-8"),"FPC_Population_Domain_Estimates.csv","text/csv")

elif page == "📊 7. Dashboard":
    st.subheader("4. FPC Survey Dashboard")

    if not st.session_state["processed"]:
        st.info("Process PDFs first.")
    else:
        tables = active_tables()
        cm = tables["company_master"]
        prof = tables["profits"]
        eq = tables["equity"]
        qa = tables["qa"]

        total_profit = np.nan
        total_div = np.nan
        if isinstance(prof, pd.DataFrame) and not prof.empty:
            total_profit = pd.to_numeric(
                prof.get("net_profit_after_tax"), errors="coerce"
            ).sum(min_count=1)
            total_div = pd.to_numeric(
                prof.get("dividends_declared"), errors="coerce"
            ).sum(min_count=1)

        qa_counts = (
            qa["status"].value_counts().to_dict()
            if not qa.empty and "status" in qa.columns else {}
        )

        cards = st.columns(6)
        cards[0].metric("Companies", company_count(cm))
        cards[1].metric("QA PASS", qa_counts.get("PASS", 0))
        cards[2].metric("QA REVIEW", qa_counts.get("REVIEW", 0))
        cards[3].metric("QA FAIL", qa_counts.get("FAIL", 0))
        cards[4].metric(
            "Net profit",
            "N/A" if pd.isna(total_profit) else f"US$ {total_profit:,.0f}"
        )
        cards[5].metric(
            "Dividends",
            "N/A" if pd.isna(total_div) else f"US$ {total_div:,.0f}"
        )

        dq = tables.get("data_quality", pd.DataFrame())
        if isinstance(dq, pd.DataFrame) and not dq.empty:
            st.markdown("### Extraction quality overview")
            qcols = st.columns(4)
            qcols[0].metric("Sections populated", int((dq.get("status", pd.Series(dtype="string")) == "POPULATED").sum()))
            qcols[1].metric("Sections needing review", int((dq.get("status", pd.Series(dtype="string")) == "REVIEW").sum()))
            qcols[2].metric("Sections empty", int((dq.get("status", pd.Series(dtype="string")) == "EMPTY").sum()))
            qcols[3].metric("QA rows", len(qa))
            st.dataframe(dq, use_container_width=True, hide_index=True)

        indicators = calculate_fpc_indicators(tables)
        if not indicators.empty:
            st.markdown("### Derived FPC indicators")
            st.dataframe(indicators, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ Download Derived Indicators (Excel)",
                data=table_excel_bytes(indicators, "FPC_Indicators"),
                file_name="FPC_Derived_Indicators.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        if not prof.empty and "company_name" in prof.columns:
            long_cols = [
                c for c in [
                    "net_profit_after_tax", "dividends_declared",
                    "dividends_paid_or_remitted"
                ] if c in prof.columns
            ]
            if long_cols:
                p = prof.copy()
                for c in long_cols:
                    p[c] = pd.to_numeric(p[c], errors="coerce")
                long = p.melt(
                    id_vars="company_name", value_vars=long_cols,
                    var_name="Indicator", value_name="USD"
                ).dropna()
                if not long.empty:
                    fig = px.bar(
                        long, x="company_name", y="USD", color="Indicator",
                        barmode="group", title="Profit and Dividend Indicators"
                    )
                    fig.update_layout(xaxis_tickangle=-35)
                    st.plotly_chart(fig, use_container_width=True)

        if not eq.empty and "closing_2025" in eq.columns:
            e = eq.copy()
            e["closing_2025"] = pd.to_numeric(e["closing_2025"], errors="coerce")
            agg = e.groupby("equity_type", as_index=False)["closing_2025"].sum(min_count=1)
            if not agg.empty:
                st.plotly_chart(
                    px.bar(
                        agg, x="equity_type", y="closing_2025",
                        title="Aggregate Closing Equity Balances"
                    ),
                    use_container_width=True
                )


# ============================================================
# 5. COMPANY EXPLORER
# ============================================================

elif page == "🏢 8. Company Explorer":
    st.subheader("5. Company Explorer")

    if not st.session_state["processed"]:
        st.info("Process PDFs first.")
    else:
        cm = active_table("company_master")
        if cm.empty:
            st.warning("No company records are available.")
        else:
            companies = sorted([
                x for x in cm["company_name"].dropna().astype(str).str.strip().unique()
                if x
            ])
            selected = st.selectbox("Select company", companies)

            rec = cm[cm["company_name"] == selected]
            st.markdown("### Company Master")
            st.dataframe(rec.T.rename(columns={rec.index[0]: "Value"}),
                         use_container_width=True)

            for title, key in [
                ("Industrial classification", "industrial"),
                ("Shareholding", "shareholding"),
                ("Equity", "equity"),
                ("Profits and dividends", "profits"),
                ("Foreign asset profits", "asset_profits"),
                ("Non-equity liabilities", "non_equity_liabilities"),
                ("Non-equity assets", "non_equity_assets"),
                ("Services", "services"),
                ("Comments", "comments"),
                ("QA", "qa"),
            ]:
                df = active_table(key)
                if not df.empty and "company_name" in df.columns:
                    sub = df[df["company_name"] == selected]
                    if not sub.empty:
                        st.markdown(f"### {title}")
                        st.dataframe(sub, use_container_width=True, hide_index=True)
                        st.download_button(
                            f"⬇️ Download {title} (Excel)",
                            data=table_excel_bytes(sub, safe_filename(title)),
                            file_name=f"{safe_filename(selected)}_{safe_filename(title)}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"company_dl_{safe_filename(selected)}_{safe_filename(title)}"
                        )


# ============================================================
# 6. DATA TABLES
# ============================================================

elif page == "📋 9. Data Tables":
    st.subheader("6. Data Tables")

    if not st.session_state["processed"]:
        st.info("Process PDFs first.")
    else:
        tables = active_tables()
        selected = st.selectbox(
            "Select dataset",
            list(tables.keys())
        )
        df = tables[selected]

        st.caption(f"{len(df):,} rows × {len(df.columns):,} columns")
        if df.empty:
            st.info("No records available.")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)

            c1, c2 = st.columns(2)
            c1.download_button(
                "⬇️ CSV",
                data=df.to_csv(index=False).encode("utf-8"),
                file_name=f"{safe_filename(selected)}.csv",
                mime="text/csv",
                use_container_width=True
            )
            c2.download_button(
                "⬇️ Excel",
                data=table_excel_bytes(df, selected),
                file_name=f"{safe_filename(selected)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# ============================================================
# 7. ANALYTICS STUDIO
# ============================================================

elif page == "📈 10. Analytics Studio":
    st.subheader("7. Analytics Studio")

    if not st.session_state["processed"]:
        st.info("Process PDFs first.")
    else:
        if not st.session_state.get("cleaned"):
            st.warning(
                "The analytical dataset has not been cleaned yet. You can explore the "
                "extracted data, but for official analytical work it is recommended to "
                "run Data Cleaning first."
            )

        tables = active_tables()
        choices = [
            k for k, v in tables.items()
            if isinstance(v, pd.DataFrame) and not v.empty
        ]
        if not choices:
            st.info("No populated datasets are available for analysis.")
            st.stop()

        def _contains_numeric(k):
            d = tables[k]
            return any(pd.api.types.is_numeric_dtype(d[c]) for c in d.columns)

        numeric_choices = [k for k in choices if _contains_numeric(k)]
        default_choice = numeric_choices[0] if numeric_choices else choices[0]
        selected = st.selectbox(
            "Dataset for analysis", choices,
            index=choices.index(default_choice)
        )
        df = tables[selected].copy()

        st.markdown(f"### {selected.replace('_',' ').title()}")
        st.caption(f"{len(df)} rows × {len(df.columns)} columns")

        tabs = st.tabs([
            "Descriptive Statistics",
            "Group Comparison",
            "Correlation",
            "Outliers",
            "Pivot / Custom Table",
            "Charts",
            "FPC Indicators",
            "BOP / IIP Bridge"
        ])

        with tabs[0]:
            stats = descriptive_statistics(df)
            if stats.empty:
                st.info("No numeric variables are available.")
            else:
                st.dataframe(stats, use_container_width=True, hide_index=True)
                st.download_button(
                    "⬇️ Download Descriptive Statistics",
                    data=table_excel_bytes(stats, "Descriptive"),
                    file_name="FPC_Descriptive_Statistics.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="analytics_desc_download"
                )

        with tabs[1]:
            num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            obj_cols = df.select_dtypes(exclude=[np.number]).columns.tolist()
            if not num_cols or not obj_cols:
                st.info("Need at least one categorical and one numeric variable.")
            else:
                gc1, gc2, gc3 = st.columns(3)
                group_col = gc1.selectbox("Group by", obj_cols, key="group_col_v3")
                value_col = gc2.selectbox("Measure", num_cols, key="group_value_v3")
                agg = gc3.selectbox("Aggregation", ["sum", "mean", "median", "count"], key="group_agg_v3")
                result = grouped_analysis(df, group_col, value_col, agg)
                st.dataframe(result, use_container_width=True, hide_index=True)
                if not result.empty:
                    fig = px.bar(result, x=group_col, y=agg, title=f"{agg.title()} of {value_col} by {group_col}")
                    st.plotly_chart(fig, use_container_width=True)
                    st.download_button(
                        "⬇️ Download Group Analysis",
                        data=table_excel_bytes(result, "GroupAnalysis"),
                        file_name="FPC_Group_Analysis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="group_download_v3"
                    )

        with tabs[2]:
            corr = correlation_matrix(df)
            if corr.empty:
                st.info("At least two numeric variables are required.")
            else:
                st.dataframe(corr, use_container_width=True)
                fig = px.imshow(corr, text_auto=".2f", aspect="auto", title="Correlation Matrix")
                st.plotly_chart(fig, use_container_width=True)
                st.download_button(
                    "⬇️ Download Correlation Matrix",
                    data=table_excel_bytes(corr.reset_index(), "Correlation"),
                    file_name="FPC_Correlation_Matrix.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="corr_download_v3"
                )

        with tabs[3]:
            outliers = outlier_report(df)
            if outliers.empty:
                st.info("No outlier analysis available.")
            else:
                st.dataframe(outliers, use_container_width=True, hide_index=True)
                st.download_button(
                    "⬇️ Download Outlier Report",
                    data=table_excel_bytes(outliers, "Outliers"),
                    file_name="FPC_Outlier_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="outlier_download_v3"
                )

        with tabs[4]:
            st.markdown("Build a two-dimensional analytical table from the current dataset.")
            obj_cols = df.select_dtypes(exclude=[np.number]).columns.tolist()
            num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            if obj_cols and num_cols:
                pc1, pc2, pc3 = st.columns(3)
                p_rows = pc1.selectbox("Rows", obj_cols, key="pivot_rows_v3")
                p_vals = pc2.selectbox("Values", num_cols, key="pivot_values_v3")
                p_agg = pc3.selectbox("Aggregation", ["sum", "mean", "count", "median"], key="pivot_agg_v3")
                pv = df.copy()
                pv[p_vals] = pd.to_numeric(pv[p_vals], errors="coerce")
                pivot = pd.pivot_table(
                    pv, index=p_rows, values=p_vals,
                    aggfunc=p_agg, dropna=False
                ).reset_index()
                st.dataframe(pivot, use_container_width=True, hide_index=True)
                st.download_button(
                    "⬇️ Download Analytical Table",
                    data=table_excel_bytes(pivot, "AnalyticalTable"),
                    file_name="FPC_Analytical_Table.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="pivot_download_v3"
                )
            else:
                st.info("Need categorical and numeric variables.")

        with tabs[5]:
            num_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            all_cols = df.columns.tolist()
            if num_cols and all_cols:
                ch1, ch2, ch3 = st.columns(3)
                xcol = ch1.selectbox("X variable", all_cols, key="chart_x_v3")
                ycol = ch2.selectbox("Y variable", num_cols, key="chart_y_v3")
                ctype = ch3.selectbox("Chart", ["Bar", "Scatter", "Box", "Histogram"], key="chart_type_v3")

                plot_df = df.copy()
                plot_df[ycol] = pd.to_numeric(plot_df[ycol], errors="coerce")
                plot_df = plot_df.dropna(subset=[ycol])

                if ctype == "Bar":
                    fig = px.bar(plot_df, x=xcol, y=ycol, title=f"{ycol} by {xcol}")
                elif ctype == "Scatter":
                    numeric_x = xcol if pd.api.types.is_numeric_dtype(plot_df[xcol]) else None
                    if numeric_x:
                        fig = px.scatter(plot_df, x=xcol, y=ycol, title=f"{ycol} vs {xcol}")
                    else:
                        st.info("Scatter requires a numeric X variable.")
                        fig = None
                elif ctype == "Box":
                    fig = px.box(plot_df, x=xcol, y=ycol, title=f"{ycol} by {xcol}")
                else:
                    fig = px.histogram(plot_df, x=ycol, title=f"Distribution of {ycol}")

                if fig is not None:
                    st.plotly_chart(fig, use_container_width=True)
                    st.download_button(
                        "⬇️ Download Chart Data",
                        data=plot_df.to_csv(index=False).encode("utf-8"),
                        file_name="FPC_Chart_Data.csv",
                        mime="text/csv",
                        key="chart_data_download_v3"
                    )

        with tabs[6]:
            ind = calculate_fpc_indicators(tables)
            if ind.empty:
                st.info("FPC-specific indicators could not be calculated from the available tables.")
            else:
                st.dataframe(ind, use_container_width=True, hide_index=True)
                st.download_button(
                    "⬇️ Download FPC Indicators",
                    data=table_excel_bytes(ind, "FPC_Indicators"),
                    file_name="FPC_Indicators.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="fpc_indicators_download_v3"
                )


        with tabs[7]:
            st.markdown("## BPM6 Financial Account & IIP Bridge")
            st.caption(
                "Compiler-support view: maps FPC survey information into BPM6 functional categories. "
                "It does not replace the complete BOP/IIP dataset."
            )
            st.info(
                "BPM6 uses five functional categories: direct investment, portfolio investment, "
                "financial derivatives/employee stock options, other investment, and reserve assets. "
                "This FPC survey has strongest coverage of direct-investment liabilities and selected "
                "other-investment liabilities."
            )

            bsum = active_table("bop_iip_summary")
            bdet = active_table("bop_iip_detail")
            irecon = active_table("iip_reconciliation")
            coverage = active_table("bop_coverage")

            m1, m2, m3, m4 = st.columns(4)
            if not bsum.empty:
                m1.metric("Companies in bridge", len(bsum))
                m2.metric(
                    "DI liability flows",
                    f"{pd.to_numeric(bsum.get('direct_investment_liability_transaction'), errors='coerce').sum(min_count=1):,.2f}"
                    if "direct_investment_liability_transaction" in bsum.columns else "—"
                )
                m3.metric(
                    "Other investment liability flows",
                    f"{pd.to_numeric(bsum.get('other_investment_liability_transaction'), errors='coerce').sum(min_count=1):,.2f}"
                    if "other_investment_liability_transaction" in bsum.columns else "—"
                )
                m4.metric(
                    "FPC liability-flow scope",
                    f"{pd.to_numeric(bsum.get('net_incurrence_of_liabilities_fpc_scope'), errors='coerce').sum(min_count=1):,.2f}"
                    if "net_incurrence_of_liabilities_fpc_scope" in bsum.columns else "—"
                )

            btabs = st.tabs(["BPM6 Summary", "Detailed Mapping", "IIP Reconciliation", "BPM6 Item Catalogue", "Coverage / Gaps"])
            with btabs[0]:
                if bsum.empty:
                    st.info("No BOP/IIP bridge could be derived from the current dataset.")
                else:
                    st.dataframe(bsum, use_container_width=True, hide_index=True)
                    st.download_button(
                        "⬇️ Download BPM6 BOP/IIP Summary",
                        data=table_excel_bytes(bsum, "BOP_IIP_Summary"),
                        file_name="FPC_BPM6_BOP_IIP_Summary.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="bop_iip_summary_dl_v34"
                    )
            with btabs[1]:
                if bdet.empty:
                    st.info("No detailed mappings are available.")
                else:
                    show_cols = [c for c in [
                        "company_name", "functional_category", "side", "instrument", "subcomponent",
                        "counterparty", "relationship", "maturity", "opening_position_2024",
                        "financial_account_transaction_2025", "gross_increase_2025", "gross_decrease_2025",
                        "price_change", "exchange_rate_change", "other_volume_change",
                        "closing_position_2025", "iip_reconciliation_residual", "derivation_status",
                        "derivation_note", "source"
                    ] if c in bdet.columns]
                    st.dataframe(bdet[show_cols], use_container_width=True, hide_index=True)
                    st.download_button(
                        "⬇️ Download Detailed BPM6 Mapping",
                        data=table_excel_bytes(bdet, "BOP_IIP_Detail"),
                        file_name="FPC_BPM6_BOP_IIP_Detail.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="bop_iip_detail_dl_v34"
                    )
            with btabs[2]:
                if irecon.empty:
                    st.info("No IIP reconciliation rows are available.")
                else:
                    st.dataframe(irecon, use_container_width=True, hide_index=True)
                    st.caption(
                        "For liabilities, the compiler relationship is: opening position + financial-account transaction + other changes = closing position. "
                        "A non-zero residual is a review signal, not automatically an error."
                    )
            with btabs[3]:
                catalog = active_table("bpm6_item_catalog")
                st.dataframe(catalog, use_container_width=True, hide_index=True)
                st.download_button(
                    "⬇️ Download BPM6 Item Catalogue",
                    data=table_excel_bytes(catalog, "BPM6_Item_Catalog"),
                    file_name="FPC_BPM6_Item_Catalogue.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="bpm6_catalog_dl_v34"
                )
            with btabs[4]:
                st.dataframe(coverage, use_container_width=True, hide_index=True)
                st.markdown("### Recommended additional BOP/IIP datasets")
                st.markdown(
                    "- **Direct investment assets:** resident enterprises' investments abroad, including equity and selected intercompany debt.\n"
                    "- **Portfolio investment:** resident holdings of foreign equity/debt securities and nonresident holdings of Zimbabwe securities.\n"
                    "- **Financial derivatives:** assets and liabilities other than reserve-related derivatives.\n"
                    "- **Other investment assets:** foreign loans, deposits, trade credit and other claims of residents on nonresidents.\n"
                    "- **Reserve assets:** monetary gold, SDR holdings, reserve position in the IMF, foreign exchange assets and other reserve claims."
                )
                st.caption(BPM6_SOURCE_NOTE)


# ============================================================
# 8. QA & AUDIT
# ============================================================

elif page == "✅ 11. QA & Audit":
    st.subheader("8. Quality Assurance & Audit")

    if not st.session_state["processed"]:
        st.info("Process PDFs first.")
    else:
        qa = active_table("qa")
        if qa.empty:
            st.info("No QA records available.")
        else:
            counts = qa["status"].value_counts().to_dict() if "status" in qa.columns else {}
            c = st.columns(4)
            c[0].metric("PASS", counts.get("PASS", 0))
            c[1].metric("REVIEW", counts.get("REVIEW", 0))
            c[2].metric("FAIL", counts.get("FAIL", 0))
            c[3].metric("INFO", counts.get("INFO", 0))

            statuses = st.multiselect(
                "Filter status",
                ["PASS", "REVIEW", "FAIL", "INFO"],
                default=["REVIEW", "FAIL"]
            )
            view = qa[qa["status"].isin(statuses)] if statuses else qa
            st.dataframe(view, use_container_width=True, hide_index=True)

            if not TESSERACT_PATH:
                st.warning(
                    "Tesseract was not detected. Scanned questionnaires cannot be reliably "
                    "OCR-extracted until Tesseract is installed."
                )

            st.markdown("### QA downloads")
            st.download_button(
                "⬇️ Download Full QA Log (Excel)",
                data=table_excel_bytes(qa, "QA_Log"),
                file_name="FPC_QA_Log.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            st.download_button(
                "⬇️ Download QA Log (CSV)",
                data=qa.to_csv(index=False).encode("utf-8"),
                file_name="FPC_QA_Log.csv",
                mime="text/csv"
            )


# ============================================================
# 9. REPORTS & TABLES
# ============================================================

elif page == "📝 12. Reports & Tables":
    st.subheader("12. Reports & Publication Outputs")

    if not st.session_state["processed"]:
        st.info("Process PDFs first.")
    else:
        tables = active_tables()
        report_type = st.radio(
            "Select report population",
            ["Sample Report", "Population Report"],
            horizontal=True,
            help="Sample Report describes the enterprises actually surveyed. Population Report presents weighted estimates for the target population represented by the supplied sampling frame."
        )

        if report_type == "Sample Report":
            title = st.text_input(
                "Sample report title",
                "Zimbabwe Foreign Private Capital Survey — Sample Statistical Report"
            )
            available = [k for k,v in tables.items() if isinstance(v,pd.DataFrame) and not v.empty]
            selected_tables = st.multiselect(
                "Sample tables to include",
                available,
                default=[x for x in ["company_master","financials","bop_iip_summary","qa"] if x in available]
            )
            selected_payload = {k:tables[k] for k in selected_tables}
            st.markdown("### Download Sample Report")
            c1,c2,c3=st.columns(3)
            with c1:
                st.download_button("⬇️ Sample Word Report",build_word_report(title,selected_payload,selected_tables),"FPC_Sample_Statistical_Report.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
            with c2:
                st.download_button("⬇️ Sample PowerPoint Report",build_ppt_report(title,selected_payload),"FPC_Sample_Statistical_Report.pptx","application/vnd.openxmlformats-officedocument.presentationml.presentation",use_container_width=True)
            with c3:
                st.download_button("⬇️ Sample PDF Report",build_pdf_report(title,selected_payload),"FPC_Sample_Statistical_Report.pdf","application/pdf",use_container_width=True)
            st.download_button("⬇️ Sample Statistical Tables — Excel",full_stage_excel_bytes(selected_payload,"Sample_Report"),"FPC_Sample_Statistical_Tables.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary",use_container_width=True)
            st.info("The Sample Report is descriptive of the responding sample. It does not expand the results to the national enterprise population.")

        else:
            if not st.session_state.get("estimation_ready"):
                st.warning("Population estimation has not been run yet. Complete Sampling Frame & Design and Weighting & Estimation first.")
                st.info("The Population Report becomes available only after a sampling frame has been supplied and the estimation engine has completed successfully.")
            else:
                title = st.text_input(
                    "Population report title",
                    "Zimbabwe Foreign Private Capital Survey — Population Estimates, BOP and IIP Report"
                )
                meta=st.session_state.get("estimation_metadata",{})
                st.markdown("### Population report package")
                st.caption(f"Frame: {meta.get('frame_units','N/A'):,} units • Sampled: {meta.get('sampled_units','N/A'):,} • Respondents: {meta.get('respondent_units','N/A'):,} • Strata: {meta.get('strata','N/A'):,}")
                st.info("The population report documents the sampling frame, weighting, nonresponse adjustment, calibration assumptions, estimated national totals, domain estimates and approximate sampling precision. The report itself is available through the download controls below and is not rendered as a report preview in the application.")
                c1,c2,c3=st.columns(3)
                with c1:
                    st.download_button("⬇️ Population Word Report",build_population_word_report(title),"FPC_Population_Estimation_Report.docx","application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
                with c2:
                    st.download_button("⬇️ Population PDF Report",build_population_pdf_report(title),"FPC_Population_Estimation_Report.pdf","application/pdf",use_container_width=True)
                with c3:
                    st.download_button("⬇️ Population Excel Tables",population_excel_bytes(),"FPC_Population_Estimation_Tables.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,type="primary")


# ============================================================
# 10. EXPORT CENTRE
# ============================================================

elif page == "📤 13. Export Center":
    st.subheader("13. Export Centre")

    if not st.session_state["processed"]:
        st.info("Process PDFs first.")
    else:
        tables = active_tables()

        if st.session_state.get("estimation_ready"):
            st.markdown("### Population estimation package")
            st.download_button(
                "⬇️ Population Estimation Package — Excel",
                data=population_excel_bytes(),
                file_name=f"FPC_Population_Estimation_Package_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        st.markdown("### Complete dataset exports")
        c1, c2 = st.columns(2)
        with c1:
            excel = full_stage_excel_bytes(tables, "Final")
            st.download_button(
                "⬇️ Final Dataset — Excel",
                data=excel,
                file_name=f"FPC_Final_Dataset_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        with c2:
            all_csv = io.BytesIO()
            with __import__("zipfile").ZipFile(all_csv, "w") as z:
                for key, df in tables.items():
                    if isinstance(df, pd.DataFrame):
                        z.writestr(
                            f"{safe_filename(key)}.csv",
                            df.to_csv(index=False)
                        )
            st.download_button(
                "⬇️ All Tables — CSV ZIP",
                data=all_csv.getvalue(),
                file_name="FPC_All_Tables_CSV.zip",
                mime="application/zip",
                use_container_width=True
            )

        st.markdown("### Stage-specific downloads")
        for stage, table_set in [
            ("Extraction", {k: st.session_state[k] for k in TABLE_KEYS}),
            ("Cleaned", active_tables()),
        ]:
            with st.expander(f"{stage} dataset"):
                x = full_stage_excel_bytes(table_set, stage)
                st.download_button(
                    f"⬇️ {stage} Stage — Excel",
                    data=x,
                    file_name=f"FPC_{stage}_Stage.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"export_stage_{stage}"
                )

        st.markdown("### Audit / provenance")
        provenance = pd.DataFrame({
            "Field": [
                "Dataset ID", "Extraction time", "Cleaning time",
                "Companies", "OCR engine", "Persistence file"
            ],
            "Value": [
                st.session_state.get("dataset_id"),
                st.session_state.get("last_run"),
                st.session_state.get("last_cleaned"),
                company_count(tables["company_master"]),
                TESSERACT_PATH or "Not detected",
                str(PERSISTENCE_FILE),
            ]
        })
        st.dataframe(provenance, use_container_width=True, hide_index=True)
        st.download_button(
            "⬇️ Download Provenance (CSV)",
            data=provenance.to_csv(index=False).encode("utf-8"),
            file_name="FPC_Provenance.csv",
            mime="text/csv"
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()
st.markdown(
    """
    <div class="app-footer">
        <strong>FPC Survey Intelligence & Population Estimation</strong><br>
        BPM6 • BOP • IIP • Foreign Private Capital Survey Processing<br>
        © 2026 Admire Tarisirayi Chirume
        <br><span style="color:#BFD3E3;">Verify OCR and manual-review values against the source questionnaire before official compilation.</span>
    </div>
    """,
    unsafe_allow_html=True,
)
