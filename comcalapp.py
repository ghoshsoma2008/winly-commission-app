# -*- coding: utf-8 -*-
"""
Commission Calculator - 2026
Complete ready-to-run Streamlit app with:
- Beautiful Admin Dashboard matching the provided screenshot
- User/CRD detail dashboard
- MSAL login with session token reuse
- Dataverse opportunity fetch
- SQLite persistent payments/clawbacks visible to all admins

Run:
    streamlit run commission_calculator_2026_complete_fixed.py
"""

import math
import os
import re
import sqlite3
from datetime import datetime
from urllib.parse import quote_plus, unquote_plus
from textwrap import dedent

import msal
import pandas as pd
import requests
import streamlit as st
import streamlit.components.v1 as components
from io import BytesIO

# =====================================================
# APP CONFIG
# =====================================================
st.set_page_config(page_title="Commission Calculator - 2026", layout="wide")

CURRENT_YEAR = 2026
TENANT_ID = "85f66ea0-8fe4-48b9-a1a7-8633937d534a"
CLIENT_ID = "c2859476-c750-4a48-8bdd-b66e5dbaa732"
DATAVERSE_URL = "https://netwoveninc.crm.dynamics.com"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
GRAPH_SCOPES = ["User.Read"]
DATAVERSE_SCOPES = [f"{DATAVERSE_URL}/user_impersonation"]

ROLE_CONFIG = {
    "admin": [
        "usarkar@netwoven.com",
        "ntenany@netwoven.com",
        "sghosh@netwoven.com",
        "shekhar.dhongade@netwoven.com",
    ],
    "users": {
        "adey@netwoven.com": "Angira Dey",
        "nsimas@netwoven.com": "Nicholas Simas",
        "chris.wilkinson@netwoven.com": "Chris Wilkinson",
        "mandeep.nagpal@netwoven.com": "Mandeep Nagpal",
    },
}

CRD_REPS = ["Nicholas Simas", "Angira Dey", "Mandeep Nagpal", "Chris Wilkinson"]

rep_config = {
    "Angira Dey": {"rate": 0.006250, "quota": 8_000_000},
    "Nicholas Simas": {"rate": 0.008000, "quota": 7_500_000},
    "Chris Wilkinson": {"rate": 0.007200, "quota": 5_000_000},
    "Mandeep Nagpal": {"rate": 0.0087600, "quota": 1_500_000},
}

OTC_CONFIG = {
    "Angira Dey": 50_000,
    "Nicholas Simas": 60_000,
    "Chris Wilkinson": 36_000,
    "Mandeep Nagpal": 13_140,
}

ELIGIBILITY_MAP = {
    "CSP": 0.25,
    "Consulting": 1.00,
    "Staffing 1 (>=20% markup)": 1.00,
    "Staffing 2 (10-19% markup)": 0.70,
    "Staffing 3 (<10% markup)": 0.00,
    "Govern 365": 1.00,
    "Product": 1.00,
}

CSP_TCV_COL = "nw_totalcontractvaluenw"
CSP_START_COL = "new_startdate"
CSP_END_COL = "new_enddate"
DB_PATH = os.path.join(os.getcwd(), "commission_ledger.db")
CACHE_PATH = os.path.join(os.getcwd(), ".msal_cache.json")

# =====================================================
# CSS
# =====================================================
def inject_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800;900&display=swap');

        html, body, [class*="css"] { font-family: 'Inter', system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; }
        .stApp { background:#F4F8FE; }
        .block-container { padding-top:1.75rem; padding-left:2.2rem; padding-right:2.2rem; max-width:1600px; }

        /* ================= SIDEBAR ================= */
        section[data-testid="stSidebar"] {
            background:linear-gradient(180deg,#07306B 0%,#061F49 100%);
            border-right:1px solid #0A3473;
        }
        section[data-testid="stSidebar"] > div { padding:1.25rem 1rem 1rem 1rem; }
        section[data-testid="stSidebar"] * { color:white; }
        section[data-testid="stSidebar"] .stMarkdown { margin-bottom:0 !important; }

        .sidebar-logo { display:flex; align-items:center; gap:14px; margin:0 0 24px 0; }
        .logo-icon { width:50px; height:50px; border-radius:11px; background:linear-gradient(135deg,#0B57D0,#0C6BDB); display:flex; align-items:center; justify-content:center; box-shadow:0 8px 20px rgba(0,0,0,.25); }
        .logo-bars { display:flex; align-items:end; gap:5px; height:27px; }
        .logo-bars span { display:block; width:7px; background:#ffffff; border-radius:2px; }
        .logo-bars span:nth-child(1){ height:16px; opacity:.88; }
        .logo-bars span:nth-child(2){ height:25px; }
        .logo-bars span:nth-child(3){ height:21px; opacity:.92; }
        .logo-text { font-size:21px; line-height:1.24; font-weight:900; letter-spacing:-.4px; }

        .sidebar-section-label { display:none !important; }

        section[data-testid="stSidebar"] .stButton > button {
            background:rgba(255,255,255,.08) !important;
            border:1px solid rgba(255,255,255,.16) !important;
            color:white !important;
            border-radius:10px !important;
            min-height:54px !important;
            font-weight:900 !important;
            font-size:18px !important;
            text-align:left !important;
            justify-content:flex-start !important;
            padding-left:22px !important;
            box-shadow:none !important;
        }
        section[data-testid="stSidebar"] .stButton > button:hover,
        section[data-testid="stSidebar"] .stButton > button:focus {
            background:linear-gradient(90deg,#0B57D0,#075AD8) !important;
            border-color:#2D7BFF !important;
            color:white !important;
        }
        section[data-testid="stSidebar"] div[data-baseweb="select"] > div {
            background:#FFFFFF !important;
            border-radius:10px !important;
            border:1px solid rgba(255,255,255,.2) !important;
            min-height:46px !important;
        }
        section[data-testid="stSidebar"] div[data-baseweb="select"] span,
        section[data-testid="stSidebar"] div[data-baseweb="select"] svg { color:#061A3A !important; fill:#061A3A !important; }
        section[data-testid="stSidebar"] label { display:none !important; }

        .signed-box {
            margin-top:110px;
            padding:15px 16px;
            border-radius:12px;
            border:1px solid rgba(255,255,255,.25);
            background:rgba(255,255,255,.08);
            box-shadow:inset 0 1px rgba(255,255,255,.12);
        }
        .signed-title { font-weight:900; font-size:14px; margin-bottom:8px; }
        .signed-email { font-size:13px; word-break:break-word; font-weight:700; }
        .signed-role { font-size:12px; margin-top:9px; opacity:.85; }

        /* ================= PAGE ================= */
        .page-title { font-size:34px; color:#061A3A; margin:0; font-weight:900; letter-spacing:-.7px; }
        .page-subtitle { color:#061A3A; font-size:16px; margin-top:10px; }
        .divider { height:1px; background:#CBD7E7; margin:18px 0 22px 0; }

        .top-actions { display:flex; justify-content:flex-end; align-items:center; gap:14px; padding-top:4px; }
        .pretty-btn {
            display:inline-flex; align-items:center; justify-content:center; gap:8px;
            min-height:44px; padding:0 18px; border-radius:8px;
            font-weight:900; font-size:14px; text-decoration:none !important;
            border:1px solid #C9D5E7; box-shadow:0 4px 12px rgba(19,38,68,.05);
        }
        .pretty-btn.secondary { color:#061A3A !important; background:white; }
        .pretty-btn.primary { color:white !important; background:#062B66; border-color:#062B66; }
        div[data-testid="stDownloadButton"] button, .main .stButton > button {
            border-radius:8px !important; min-height:44px !important; font-weight:900 !important;
        }

        /* Native top action buttons */
        div[data-testid="stDownloadButton"] { display:flex !important; justify-content:flex-end !important; }
        div[data-testid="stDownloadButton"] button {
            width:auto !important;
            min-width:150px !important;
            max-width:190px !important;
            background:#062B66 !important;
            color:white !important;
            border:1px solid #062B66 !important;
            border-radius:8px !important;
            box-shadow:0 6px 16px rgba(6,43,102,.18) !important;
            white-space:nowrap !important;
            padding-left:16px !important;
            padding-right:16px !important;
            font-weight:900 !important;
        }
        div[data-testid="stDownloadButton"] button:hover {
            background:#041F4C !important;
            color:white !important;
            border-color:#041F4C !important;
        }
        div[data-testid="stDownloadButton"] button p,
        div[data-testid="stDownloadButton"] button span { color:white !important; }
        .main .stButton > button {
            background:white !important;
            color:#061A3A !important;
            border:1px solid #C9D5E7 !important;
            box-shadow:0 4px 12px rgba(19,38,68,.05) !important;
            white-space:nowrap !important;
        }

        /* ================= KPI ================= */
        .kpi-card {
            background:#fff;
            border:1.5px solid #C9D8EA;
            border-radius:13px;
            padding:18px 18px;
            min-height:142px;
            box-shadow:0 9px 24px rgba(19,38,68,.10);
        }
        .kpi-head { display:flex; align-items:center; gap:13px; min-height:48px; }
        .kpi-icon { width:48px; height:48px; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:24px; flex-shrink:0; }
        .kpi-purple { background:#EEE7FF; color:#5B3FD6; }
        .kpi-green { background:#DFF8E9; color:#16A34A; }
        .kpi-blue { background:#DCEBFF; color:#0B65D8; }
        .kpi-gold { background:#FFF1C8; color:#D88900; }
        .kpi-red { background:#FFE1E4; color:#C8303D; }
        .kpi-title { color:#061A3A; font-size:13px; font-weight:800; line-height:1.35; }
        .kpi-value { color:#020B22; font-size:22px; font-weight:900; margin-top:14px; }
        .kpi-sub { color:#061A3A; font-size:12px; margin-top:16px; text-align:center; }

        /* ================= ADMIN TABLE ================= */
        .admin-table-wrap { background:white; border:1px solid #D5E0EF; border-radius:9px; overflow:hidden; box-shadow:0 8px 20px rgba(19,38,68,.06); margin-top:18px; }
        table.admin-table { width:100%; border-collapse:collapse; background:white; }
        table.admin-table th { background:#062B66; color:white; padding:15px 16px; font-weight:900; text-align:center; font-size:15px; border-right:1px solid rgba(255,255,255,.08); }
        table.admin-table td { padding:16px 18px; text-align:center; border:1px solid #DCE4EF; color:#061A3A; font-size:15px; font-weight:600; }
        table.admin-table td:first-child { text-align:left; }
        .rep-link { color:#0052CC !important; text-decoration:none !important; font-weight:900; }
        .status { display:inline-block; min-width:98px; padding:9px 14px; border-radius:7px; font-weight:900; font-size:13px; }
        .status-ramping { background:#E2F6E6; color:#14712E; border:1px solid #BDEBC8; }
        .status-locked { background:#FFE5E7; color:#C1121F; border:1px solid #FFC8CE; }
        .status-unlocked { background:#DBF3FF; color:#075A9C; border:1px solid #B8E4FF; }
        .view-details { display:inline-block; min-width:116px; padding:9px 12px; border-radius:7px; border:1px solid #2D7BFF; color:#0052CC !important; text-decoration:none !important; background:white; font-weight:900; }
        .info-note { margin-top:22px; border:1px solid #A9CDFB; color:#061A3A; background:#F0F7FF; border-radius:8px; padding:13px 18px; font-size:14px; }
        .info-dot { display:inline-flex; align-items:center; justify-content:center; width:22px; height:22px; border-radius:50%; background:#0B65D8; color:white; font-weight:900; margin-right:8px; }

        /* ================= SUMMARY CARDS ================= */
        .bottom-grid { display:grid; grid-template-columns:1fr 1fr; gap:20px; margin-top:18px; }
        .card { background:white; border:1px solid #D5E0EF; border-radius:13px; padding:18px 20px; box-shadow:0 8px 24px rgba(19,38,68,.06); }
        .box-title { color:#061A3A; font-size:20px; font-weight:900; }
        .box-line { height:1px; background:#CFD8E7; margin:14px 0; }
        .summary-head { display:grid; grid-template-columns:1fr 160px; padding:2px 10px 10px; color:#061A3A; font-size:14px; font-weight:900; }
        .summary-row { display:grid; grid-template-columns:1fr 160px; align-items:center; padding:12px 12px; border-radius:7px; margin-bottom:5px; color:#061A3A; font-size:14px; }
        .summary-row > div:last-child { text-align:right; font-weight:800; }
        .row-green { background:#EAF8EE; } .row-red { background:#FFF0F1; color:#D00016; } .row-soft { background:#F9FBFE; }
        .recent-row { display:grid; grid-template-columns:42px 1fr 120px 145px; align-items:center; gap:10px; padding:14px 10px; border:1px solid #D5E0EF; border-radius:10px; color:#061A3A; font-size:14px; margin-bottom:8px; }
        .circle-ok,.circle-bad,.circle-info { width:28px; height:28px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:white; font-weight:900; }
        .circle-ok,.circle-info { background:#12A66A; } .circle-bad { background:#E85462; }
        .money-green { color:#058A31; font-weight:900; text-align:right; } .money-red { color:#D00016; font-weight:900; text-align:right; } .small-muted { color:#1F2D4A; font-size:12px; text-align:right; } .view-all { color:#0052CC; font-weight:900; }

        .detail-card { background:white; border:1px solid #D5E0EF; border-radius:13px; padding:20px; box-shadow:0 8px 24px rgba(19,38,68,.06); margin:18px 0; }
        table.summary-table { width:100%; border-collapse:collapse; background:white; border-radius:10px; overflow:hidden; margin-top:18px; }
        table.summary-table th { background:#062B66; color:white; padding:12px; text-align:left; }
        table.summary-table td { padding:11px 12px; border:1px solid #E5ECF5; }
        table.summary-table td.value { text-align:right; font-weight:800; }

        section[data-testid="stSidebar"] div[data-testid="stSelectbox"] { margin-top:-4px !important; margin-bottom:14px !important; }
        section[data-testid="stSidebar"] div[data-testid="stSelectbox"] > label { display:none !important; height:0 !important; min-height:0 !important; margin:0 !important; padding:0 !important; }

        .summary-logic-grid { display:grid; grid-template-columns:minmax(430px, 1.05fr) minmax(330px, .95fr); gap:18px; align-items:start; margin-top:18px; }
        .logic-card { border:1px solid #D0D5DD; border-radius:12px; padding:16px 18px; background:#F9FAFB; box-shadow:0 6px 18px rgba(19,38,68,.05); color:#061A3A; font-size:14px; line-height:1.45; }
        .logic-card h4 { color:#1F4E79; margin:0 0 10px 0; font-size:18px; font-weight:900; }
        .logic-card ul { margin-top:6px; margin-bottom:10px; padding-left:20px; }
        table.summary-table.compact { margin-top:0; font-size:13px; }
        table.summary-table.compact th { padding:9px 11px; }
        table.summary-table.compact td { padding:8px 10px; }
        </style>
        """,
        unsafe_allow_html=True,
    )

# =====================================================
# FORMATTERS + QUERIES
# =====================================================
def fmt_money(x):
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"


def fmt_pct(x):
    try:
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return "0.00%"


def clear_query_params():
    try:
        st.query_params.clear()
    except Exception:
        pass


def get_requested_crd():
    try:
        val = st.query_params.get("crd")
        return unquote_plus(val) if val else None
    except Exception:
        return None


def requested_dashboard():
    try:
        return st.query_params.get("view") == "dashboard"
    except Exception:
        return False


def set_view(view, selected_rep=None):
    st.session_state["view"] = view
    st.session_state["selected_rep"] = selected_rep
    clear_query_params()
    st.rerun()

# =====================================================
# AUTH
# =====================================================
def _load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_PATH):
        try:
            with open(CACHE_PATH, "r", encoding="utf-8") as f:
                cache.deserialize(f.read())
        except Exception:
            pass
    return cache


def _save_cache(cache):
    if cache.has_state_changed:
        try:
            with open(CACHE_PATH, "w", encoding="utf-8") as f:
                f.write(cache.serialize())
        except Exception:
            pass


def get_msal_app():
    if "msal_cache" not in st.session_state:
        st.session_state["msal_cache"] = _load_cache()
    if "msal_app" not in st.session_state:
        st.session_state["msal_app"] = msal.PublicClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            token_cache=st.session_state["msal_cache"],
        )
    return st.session_state["msal_app"], st.session_state["msal_cache"]


def acquire_token_device_flow(app, cache, scopes):
    flow = app.initiate_device_flow(scopes=scopes)
    if not flow or "user_code" not in flow:
        st.error("Failed to start Microsoft login. Please check Azure app registration.")
        st.stop()
    st.info(flow.get("message", "Complete the Microsoft device login."))
    result = app.acquire_token_by_device_flow(flow)
    if not result or "access_token" not in result:
        st.error(result.get("error_description", "Microsoft authentication failed."))
        st.stop()
    _save_cache(cache)
    return result["access_token"]


def get_token_for_scopes(scopes, token_key):
    if token_key in st.session_state:
        return st.session_state[token_key]
    app, cache = get_msal_app()
    accounts = app.get_accounts()
    if accounts:
        silent = app.acquire_token_silent(scopes, account=accounts[0])
        if silent and "access_token" in silent:
            st.session_state[token_key] = silent["access_token"]
            _save_cache(cache)
            return st.session_state[token_key]
    st.session_state[token_key] = acquire_token_device_flow(app, cache, scopes)
    return st.session_state[token_key]


def get_logged_in_user_upn(graph_token):
    if "signed_in_upn" in st.session_state:
        return st.session_state["signed_in_upn"]
    headers = {"Authorization": f"Bearer {graph_token}"}
    r = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers, timeout=30)
    r.raise_for_status()
    st.session_state["signed_in_upn"] = r.json()["userPrincipalName"].lower()
    return st.session_state["signed_in_upn"]

# =====================================================
# DATABASE
# =====================================================
def db_connect():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn


def ensure_database():
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS payments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rep TEXT,
                amount REAL NOT NULL DEFAULT 0,
                paid_by TEXT,
                ts TEXT NOT NULL
            )
        """)
        cur.execute("""
            CREATE TABLE IF NOT EXISTS clawbacks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rep TEXT,
                project TEXT,
                amount REAL NOT NULL DEFAULT 0,
                created_by TEXT,
                ts TEXT NOT NULL
            )
        """)
        pay_cols = [r[1] for r in cur.execute("PRAGMA table_info(payments)").fetchall()]
        if "rep" not in pay_cols:
            cur.execute("ALTER TABLE payments ADD COLUMN rep TEXT")
        if "paid_by" not in pay_cols:
            cur.execute("ALTER TABLE payments ADD COLUMN paid_by TEXT")
        if "ts" not in pay_cols:
            cur.execute("ALTER TABLE payments ADD COLUMN ts TEXT")
        pay_cols = [r[1] for r in cur.execute("PRAGMA table_info(payments)").fetchall()]
        if "rep_name" in pay_cols:
            cur.execute("UPDATE payments SET rep = COALESCE(rep, rep_name)")
        if "created_at" in pay_cols:
            cur.execute("UPDATE payments SET ts = COALESCE(ts, created_at)")

        claw_cols = [r[1] for r in cur.execute("PRAGMA table_info(clawbacks)").fetchall()]
        if "rep" not in claw_cols:
            cur.execute("ALTER TABLE clawbacks ADD COLUMN rep TEXT")
        if "project" not in claw_cols:
            cur.execute("ALTER TABLE clawbacks ADD COLUMN project TEXT")
        if "created_by" not in claw_cols:
            cur.execute("ALTER TABLE clawbacks ADD COLUMN created_by TEXT")
        if "ts" not in claw_cols:
            cur.execute("ALTER TABLE clawbacks ADD COLUMN ts TEXT")
        claw_cols = [r[1] for r in cur.execute("PRAGMA table_info(clawbacks)").fetchall()]
        if "rep_name" in claw_cols:
            cur.execute("UPDATE clawbacks SET rep = COALESCE(rep, rep_name)")
        if "project_name" in claw_cols:
            cur.execute("UPDATE clawbacks SET project = COALESCE(project, project_name)")
        if "created_at" in claw_cols:
            cur.execute("UPDATE clawbacks SET ts = COALESCE(ts, created_at)")
        conn.commit()


def add_payment(rep, amount, paid_by):
    with db_connect() as conn:
        conn.execute(
            "INSERT INTO payments(rep, amount, paid_by, ts) VALUES (?, ?, ?, ?)",
            (rep, float(amount), paid_by, datetime.utcnow().isoformat(timespec="seconds") + "Z"),
        )
        conn.commit()

def delete_latest_payment(rep_name: str):
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.cursor()
        cur.execute(
            """
            DELETE FROM payments
            WHERE id = (
                SELECT id
                FROM payments
                WHERE rep=?
                ORDER BY id DESC
                LIMIT 1
            )
            """,
            (rep_name,)
        )
        conn.commit()


def add_clawback(rep_name: str, project_name: str, clawback_amount: float, created_by: str):

    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.cursor()

        cur.execute(
            """
            INSERT INTO clawbacks (rep, project, amount, ts)
            VALUES (?, ?, ?, ?)
            """,
            (
                rep_name,
                project_name,
                float(clawback_amount),
                datetime.utcnow().isoformat(timespec="seconds") + "Z"
            )
        )

        conn.commit()
def clear_clawbacks(rep_name: str):

    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.cursor()

        cur.execute(
            "DELETE FROM clawbacks WHERE rep=?",
            (rep_name,)
        )

        conn.commit()

def quarter_months(quarter):
    return {"Q1": ("01", "02", "03"), "Q2": ("04", "05", "06"), "Q3": ("07", "08", "09"), "Q4": ("10", "11", "12")}.get(quarter, ())


def _where_rep_quarter(rep=None, quarter="All"):
    where, params = [], []
    if rep:
        where.append("rep=?")
        params.append(rep)
    if quarter != "All":
        months = quarter_months(quarter)
        if months:
            where.append("substr(ts,6,2) IN (%s)" % ",".join(["?"] * len(months)))
            params.extend(months)
    return where, params


def get_total_paid(rep=None, quarter="All"):
    with db_connect() as conn:
        where, params = _where_rep_quarter(rep, quarter)
        sql = "SELECT COALESCE(SUM(amount),0) FROM payments"
        if where:
            sql += " WHERE " + " AND ".join(where)
        row = conn.execute(sql, tuple(params)).fetchone()
        return float(row[0] or 0.0)


def get_total_clawback(rep=None, quarter="All"):
    with db_connect() as conn:
        where, params = _where_rep_quarter(rep, quarter)
        sql = "SELECT COALESCE(SUM(amount),0) FROM clawbacks"
        if where:
            sql += " WHERE " + " AND ".join(where)
        row = conn.execute(sql, tuple(params)).fetchone()
        return float(row[0] or 0.0)


def get_payment_history(rep=None, limit=None):
    with db_connect() as conn:
        conn.row_factory = sqlite3.Row
        params = []
        sql = "SELECT ts, rep, amount, paid_by FROM payments"
        if rep:
            sql += " WHERE rep=?"
            params.append(rep)
        sql += " ORDER BY id DESC"
        if limit:
            sql += " LIMIT ?"
            params.append(limit)
        return [dict(r) for r in conn.execute(sql, tuple(params)).fetchall()]


def get_clawbacks(rep=None, limit=None):
    with db_connect() as conn:
        conn.row_factory = sqlite3.Row
        params = []
        sql = "SELECT ts, rep, project, amount, created_by FROM clawbacks"
        if rep:
            sql += " WHERE rep=?"
            params.append(rep)
        sql += " ORDER BY id DESC"
        if limit:
            sql += " LIMIT ?"
            params.append(limit)
        return [dict(r) for r in conn.execute(sql, tuple(params)).fetchall()]


def get_recent_activity(limit=5):
    payments = [{"kind": "payment", **r} for r in get_payment_history(limit=limit)]
    clawbacks = [{"kind": "clawback", **r} for r in get_clawbacks(limit=limit)]
    rows = payments + clawbacks
    rows.sort(key=lambda r: r.get("ts", ""), reverse=True)
    return rows[:limit]

# =====================================================
# DATA FETCH + COMMISSION LOGIC
# =====================================================
@st.cache_data(show_spinner=False, ttl=600)
def fetch_opportunities(dataverse_token):
    headers = {
        "Authorization": f"Bearer {dataverse_token}",
        "Accept": "application/json",
        "Prefer": 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
    }
    start = f"{CURRENT_YEAR}-01-01T00:00:00Z"
    end = f"{CURRENT_YEAR}-12-31T23:59:59Z"
    url = (
        f"{DATAVERSE_URL}/api/data/v9.2/opportunities"
        "?$select=opportunityid,name,actualvalue,actualclosedate,_ownerid_value,_nw_workload_value,"
        f"{CSP_TCV_COL},{CSP_START_COL},{CSP_END_COL}"
        f"&$filter=statecode eq 1 and actualclosedate ge {start} and actualclosedate le {end}"
    )
    r = requests.get(url, headers=headers, timeout=60)
    r.raise_for_status()
    df = pd.DataFrame(r.json().get("value", []))
    if df.empty:
        return df
    df["Opportunity Id"] = df.get("opportunityid")
    df["Client Name"] = df.get("name")
    df["Deal Value"] = pd.to_numeric(df.get("actualvalue"), errors="coerce").fillna(0.0)
    df["Sales Rep"] = df.get("_ownerid_value@OData.Community.Display.V1.FormattedValue", "").fillna("").astype(str)
    df["Workload"] = df.get("_nw_workload_value@OData.Community.Display.V1.FormattedValue", "").fillna("").astype(str)
    df["actualclosedate"] = pd.to_datetime(df.get("actualclosedate"), errors="coerce")
    for col in [CSP_TCV_COL, CSP_START_COL, CSP_END_COL]:
        if col not in df.columns:
            df[col] = pd.NA
    df[CSP_TCV_COL] = pd.to_numeric(df[CSP_TCV_COL], errors="coerce")
    df[CSP_START_COL] = pd.to_datetime(df[CSP_START_COL], errors="coerce")
    df[CSP_END_COL] = pd.to_datetime(df[CSP_END_COL], errors="coerce")
    return df[df["Sales Rep"].str.strip() != ""].copy()


@st.cache_data(show_spinner=False)
def load_staffing_rates():
    try:
        rates = pd.read_excel("rates.xlsx", engine="openpyxl")
        rates["Bill Rate"] = pd.to_numeric(rates["Bill Rate"], errors="coerce")
        rates["Pay Rate"] = pd.to_numeric(rates["Pay Rate"], errors="coerce")
        return rates
    except Exception:
        return pd.DataFrame()


def derive_category(workload, client_name):
    txt = f"{workload} {client_name}".lower()
    if "nintex" in txt:
        return "Product"
    if "business apps" in txt or "consult" in txt:
        return "Consulting"
    if "csp" in txt:
        return "CSP"
    if "staffing" in txt or "staff" in txt:
        return "Staffing 2 (10-19% markup)"
    if "govern" in txt:
        return "Govern 365"
    return "Consulting"


def match_rep_name(df, target_name):
    if df is None or df.empty:
        return target_name
    target = str(target_name).lower().strip()
    parts = [p for p in target.split() if p]
    candidates = []
    for rep in sorted(df["Sales Rep"].dropna().unique()):
        rep_l = str(rep).lower()
        if all(p in rep_l for p in parts):
            candidates.append(rep)
    return candidates[0] if candidates else target_name


def lookup_staffing_markup(client_name, rates_df):
    if not client_name or rates_df is None or rates_df.empty:
        return None
    cn = str(client_name).strip()
    parts = [p.strip() for p in cn.split("-")]
    person_part = parts[1] if len(parts) >= 2 else cn
    person_part = re.sub(r"\b(renewal|renew|extension|support|project|phase)\b", "", person_part, flags=re.I)
    person_part = re.sub(r"\b(20\d{2})\b", "", person_part)
    person_part = re.sub(r"\s+", " ", person_part).strip()
    search_cols = ["Resource", "Project Name"]
    for col in search_cols:
        if col in rates_df.columns:
            keys = [person_part, parts[0] if parts else "", cn]
            for key in [k for k in keys if k]:
                m = rates_df[rates_df[col].astype(str).str.contains(key, case=False, na=False)]
                if not m.empty:
                    row = m.iloc[0]
                    bill = pd.to_numeric(row.get("Bill Rate"), errors="coerce")
                    pay = pd.to_numeric(row.get("Pay Rate"), errors="coerce")
                    if pd.notna(bill) and pd.notna(pay) and pay != 0:
                        return (bill - pay) / pay
    return None


def staffing_tier(markup_ratio):
    try:
        m = float(markup_ratio)
    except Exception:
        return pd.NA
    if m >= 0.20:
        return "Staffing 1 (>=20% markup)"
    if m >= 0.10:
        return "Staffing 2 (10-19% markup)"
    return "Staffing 3 (<10% markup)"


def compute_csp_years(start_date, end_date):
    if pd.isna(start_date) or pd.isna(end_date):
        return 1
    days = (end_date - start_date).days
    return max(1, int(math.ceil(days / 365.0))) if days > 0 else 1


def build_rep_df(all_df, rep_display_name):
    if all_df is None or all_df.empty:
        return pd.DataFrame(columns=["Client Name", "Final Revenue Category", "Deal Value", "Quota Credit", "Workload", "actualclosedate"])
    rep_df = all_df[all_df["Sales Rep"] == rep_display_name].copy()
    if rep_df.empty:
        return pd.DataFrame(columns=["Client Name", "Final Revenue Category", "Deal Value", "Quota Credit", "Workload", "actualclosedate"])
    rep_df["actualclosedate"] = pd.to_datetime(rep_df["actualclosedate"], errors="coerce")
    rep_df["Close Month"] = rep_df["actualclosedate"].dt.strftime("%B")
    rep_df["Derived Revenue Category"] = rep_df.apply(lambda r: derive_category(r.get("Workload", ""), r.get("Client Name", "")), axis=1)
    rates_df = load_staffing_rates()
    rep_df["Markup"] = rep_df.apply(lambda r: lookup_staffing_markup(r.get("Client Name", ""), rates_df) if "staff" in str(r.get("Workload", "")).lower() else None, axis=1)
    rep_df["Markup_num"] = pd.to_numeric(rep_df["Markup"], errors="coerce")
    rep_df["Final Revenue Category"] = rep_df["Derived Revenue Category"]
    mask_staff = rep_df["Workload"].astype(str).str.contains("staff", case=False, na=False)
    rep_df.loc[mask_staff, "Final Revenue Category"] = rep_df.loc[mask_staff, "Markup_num"].apply(staffing_tier).fillna(rep_df.loc[mask_staff, "Final Revenue Category"])
    rep_df["Eligibility %"] = rep_df["Final Revenue Category"].map(ELIGIBILITY_MAP).fillna(0.0)
    rep_df["Quota Credit"] = rep_df["Deal Value"] * rep_df["Eligibility %"]
    rep_df["TCV"] = pd.to_numeric(rep_df.get(CSP_TCV_COL, 0), errors="coerce").fillna(0.0)
    rep_df[CSP_START_COL] = pd.to_datetime(rep_df.get(CSP_START_COL), errors="coerce")
    rep_df[CSP_END_COL] = pd.to_datetime(rep_df.get(CSP_END_COL), errors="coerce")

    def csp_bonus(r):
        if str(r.get("Final Revenue Category", "")).strip() != "CSP":
            return 0.0
        years = compute_csp_years(r.get(CSP_START_COL), r.get(CSP_END_COL))
        tcv = float(r.get("TCV", 0.0) or 0.0)
        return tcv * 0.0025 if years >= 3 and tcv > 0 else 0.0

    rep_df["CSP Bonus (Year1 Only)"] = rep_df.apply(csp_bonus, axis=1)
    return rep_df


def filter_df_by_quarter(rep_df, quarter):
    if quarter == "All" or rep_df.empty or "actualclosedate" not in rep_df.columns:
        return rep_df
    qmap = {"Q1": [1, 2, 3], "Q2": [4, 5, 6], "Q3": [7, 8, 9], "Q4": [10, 11, 12]}
    months = qmap.get(quarter, [])
    out = rep_df.copy()
    out["actualclosedate"] = pd.to_datetime(out["actualclosedate"], errors="coerce")
    return out[out["actualclosedate"].dt.month.isin(months)]


def compute_summary(rep_df, rep_name, quarter="All"):
    cfg = rep_config.get(rep_name, {"rate": 0.0, "quota": 0.0})
    rate = float(cfg.get("rate", 0.0) or 0.0)
    quota = float(cfg.get("quota", 0.0) or 0.0)

    total_credit = float(rep_df["Quota Credit"].sum()) if not rep_df.empty and "Quota Credit" in rep_df.columns else 0.0
    attainment = total_credit / quota if quota else 0.0

    if attainment < 0.2:
        payout_factor = 0.0
    elif attainment >= 0.8:
        payout_factor = 1.0
    else:
        payout_factor = (attainment - 0.2) / 0.5

    payout_status = "LOCKED" if payout_factor == 0 else ("RAMPING" if payout_factor < 1 else "UNLOCKED")

    # Base eligible commission = quota credit × CRD rate.
    # IMPORTANT: Multi-year bonus is kept separate so it is NOT double-added.
    base = total_credit * rate
    bonus = float(pd.to_numeric(rep_df.get("CSP Bonus (Year1 Only)", pd.Series(dtype=float)), errors="coerce").fillna(0.0).sum()) if not rep_df.empty else 0.0
    clawback = get_total_clawback(rep_name, quarter)

    eligible = max(base - abs(clawback), 0.0)
    total_eligible_with_bonus = max(base + bonus - abs(clawback), 0.0)
    payable = total_eligible_with_bonus * payout_factor
    paid = get_total_paid(rep_name, quarter)
    remaining = max(payable - paid, 0.0)

    return {
        "quota": quota,
        "total_quota_credit": total_credit,
        "attainment": attainment,
        "threshold": quota * 0.8 if quota else 0.0,
        "gap_to_full": max(0.0, quota - total_credit) if quota else 0.0,
        "payout_factor": payout_factor,
        "payout_status": payout_status,
        "base_eligible_comm": base,
        "multi_year_bonus_comm": bonus,
        "total_clawback": clawback,
        "eligible_comm": eligible,
        "total_eligible_with_bonus": total_eligible_with_bonus,
        "paid_comm": payable,
        "manual_paid": get_total_paid(rep_name),
        "remaining_comm": max(payable - get_total_paid(rep_name), 0.0),
        "rate": rate,
        }

# =====================================================
# SIDEBAR + UI
# =====================================================
def render_sidebar(signed_in_upn, role):
    with st.sidebar:
        st.markdown(
            """
            <div class="sidebar-logo">
              <div class="logo-icon"><div class="logo-bars"><span></span><span></span><span></span></div></div>
              <div class="logo-text">Commission<br>Calculator - 2026</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button("🏠  Dashboard", use_container_width=True, key="nav_dashboard_main"):
            set_view("dashboard", None)

        # CRD navigation is handled from the Admin Dashboard table using the View Details buttons.
        # The old sidebar CRD dropdown is intentionally removed to avoid duplicate navigation.
        if role != "ADMIN":
            if st.button("👤  My Details", use_container_width=True, key="nav_my_details"):
                set_view("detail", st.session_state.get("selected_rep"))

        if st.button("💵  Payments", use_container_width=True, key="nav_payments"):
            set_view("payments", st.session_state.get("selected_rep"))
        if st.button("📋  Clawbacks", use_container_width=True, key="nav_clawbacks"):
            set_view("clawbacks", st.session_state.get("selected_rep"))
        if st.button("📊  Reports", use_container_width=True, key="nav_reports"):
            st.session_state["view"] = "reports"
            st.rerun()
            set_view("reports", st.session_state.get("selected_rep"))
        if st.button("⚙️  Settings", use_container_width=True, key="nav_settings"):
            set_view("settings", st.session_state.get("selected_rep"))

        st.markdown(
            f"""
            <div class="signed-box">
                <div class="signed-title">Signed in as</div>
                <div class="signed-email">{signed_in_upn}</div>
                <div class="signed-role">Role: {role}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("↪  Logout", use_container_width=True, key="logout_btn"):
            for k in ["graph_token", "dataverse_token", "signed_in_upn", "view", "selected_rep"]:
                st.session_state.pop(k, None)
            st.session_state["logged_out"] = True
            clear_query_params()
            st.rerun()


def render_kpi_card(title, value, subtitle, icon, icon_class):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-head"><div class="kpi-icon {icon_class}">{icon}</div><div class="kpi-title">{title}</div></div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_page_header(title, subtitle, show_back=False):
    left, right = st.columns([6.8, 1.25])
    with left:
        st.markdown(
            f'<h1 class="page-title">{title}</h1><div class="page-subtitle">{subtitle}</div>',
            unsafe_allow_html=True,
        )
    with right:
        if show_back:
            
            if st.button("← Back to Dashboard", use_container_width=True, key=f"back_btn_{title}"):
                set_view("dashboard", None)
        else:
            st.download_button(
                "⇩ Export Summary",
                data="Commission Calculator 2026 Summary\n",
                file_name="commission_summary.csv",
                mime="text/csv",
                use_container_width=False,
                key=f"export_btn_{title}_{st.session_state.get('view','dashboard')}"
            )
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

# =====================================================
# DASHBOARDS
# =====================================================
def render_admin_dashboard(df_all):
    render_page_header(f"Admin Dashboard - CRD Summary ({CURRENT_YEAR})", "Click a CRD name to open the detailed commission view.", show_back=False)
    quarter = st.selectbox("Filter Dashboard by Quarter", ["All", "Q1", "Q2", "Q3", "Q4"], index=0, key="admin_quarter_filter_main")

    rows, totals = [], {"quota": 0.0, "credit": 0.0, "eligible": 0.0, "bonus": 0.0}
    for rep in CRD_REPS:
        rep_display = match_rep_name(df_all, rep)
        rep_df = filter_df_by_quarter(build_rep_df(df_all, rep_display), quarter)
        summary = compute_summary(rep_df, rep, quarter)
        rows.append((rep, summary))
        totals["quota"] += summary["quota"]
        totals["credit"] += summary["total_quota_credit"]
        totals["eligible"] += summary["eligible_comm"]
        totals["bonus"] += summary["multi_year_bonus_comm"]

    avg_attainment = totals["credit"] / totals["quota"] if totals["quota"] else 0.0
    total_paid = get_total_paid(quarter=quarter)
    total_clawback = get_total_clawback(quarter=quarter)

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: render_kpi_card("Total Annual Quota", fmt_money(totals["quota"]), "Across 4 CRDs", "🎯", "kpi-purple")
    with c2: render_kpi_card("Total Quota Credit", fmt_money(totals["credit"]), "YTD", "📈", "kpi-green")
    with c3: render_kpi_card("Total Eligible Comm YTD", fmt_money(totals["eligible"]), "After Bonus & Clawbacks", "$", "kpi-blue")
    with c4: render_kpi_card("Avg Attainment", fmt_pct(avg_attainment), "Across 4 CRDs", "◔", "kpi-gold")
    with c5: render_kpi_card("Total Commission Paid", fmt_money(total_paid), "YTD", "💼", "kpi-red")

    st.markdown('\n    <style>\n      .native-admin-table { background:white; border:1px solid #D5E0EF; border-radius:10px; overflow:hidden; box-shadow:0 8px 20px rgba(19,38,68,.06); margin-top:18px; }\n      .native-admin-header { background:#062B66; color:white; font-weight:900; padding:14px 10px; text-align:center; border-right:1px solid rgba(255,255,255,.10); }\n      .native-admin-cell { background:white; padding:14px 10px; min-height:58px; border-right:1px solid #DCE4EF; border-bottom:1px solid #DCE4EF; display:flex; align-items:center; justify-content:center; color:#061A3A; font-weight:650; }\n      .native-admin-name { justify-content:flex-start; color:#0052CC; font-weight:900; padding-left:18px; }\n      .status-pill { display:inline-block; min-width:98px; padding:8px 12px; border-radius:7px; font-weight:900; font-size:13px; text-align:center; }\n      .status-ramping { background:#E2F6E6; color:#14712E; border:1px solid #BDEBC8; }\n      .status-locked { background:#FFE5E7; color:#C1121F; border:1px solid #FFC8CE; }\n      .status-unlocked { background:#DBF3FF; color:#075A9C; border:1px solid #B8E4FF; }\n      div[data-testid="stButton"] button[kind="secondary"] { font-weight:900; }\n    </style>\n    ', unsafe_allow_html=True)

    st.markdown('<div class="native-admin-table">', unsafe_allow_html=True)
    hcols = st.columns([1.25, 1.35, 1.35, 1.35, 1.25, 1.1, 1.25], gap="small")
    headers = ["CRD Name", "Annual Quota Goal", "Total Quota Credit", "Eligible Comm YTD", "YTD Attainment %", "Payout Status", "Action"]
    for c, h in zip(hcols, headers):
        with c:
            st.markdown(f'<div class="native-admin-header">{h}</div>', unsafe_allow_html=True)

    for rep, summary in rows:
        status = summary["payout_status"]
        status_class = "status-locked" if status == "LOCKED" else ("status-unlocked" if status == "UNLOCKED" else "status-ramping")
        cols = st.columns([1.25, 1.35, 1.35, 1.35, 1.25, 1.1, 1.25], gap="small")
        with cols[0]:
            st.markdown(f'<div class="native-admin-cell native-admin-name">{rep}</div>', unsafe_allow_html=True)
        with cols[1]:
            st.markdown(f'<div class="native-admin-cell">{fmt_money(summary["quota"])}</div>', unsafe_allow_html=True)
        with cols[2]:
            st.markdown(f'<div class="native-admin-cell">{fmt_money(summary["total_quota_credit"])}</div>', unsafe_allow_html=True)
        with cols[3]:
            st.markdown(f'<div class="native-admin-cell">{fmt_money(summary["eligible_comm"])}</div>', unsafe_allow_html=True)
        with cols[4]:
            st.markdown(f'<div class="native-admin-cell">{fmt_pct(summary["attainment"])}</div>', unsafe_allow_html=True)
        with cols[5]:
            st.markdown(f'<div class="native-admin-cell"><span class="status-pill {status_class}">{status}</span></div>', unsafe_allow_html=True)
        with cols[6]:
            if st.button("👁 View Details", key=f"view_details_{rep}", use_container_width=True):
                st.session_state["selected_rep"] = rep
                st.session_state["view"] = "detail"
                clear_query_params()
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown(
        '<div class="info-note"><span class="info-dot">i</span><b>Note:</b> Values are calculated based on data available up to the current date.</div>',
        unsafe_allow_html=True,
    )

    recent = get_recent_activity(limit=5)
    if not recent:
        recent_html = '<div class="recent-row"><div class="circle-info">i</div><div>No payment or clawback activity yet</div><div></div><div></div></div>'
    else:
        recent_parts = []
        for r in recent:
            is_payment = r["kind"] == "payment"
            circle = "circle-ok" if is_payment else "circle-bad"
            icon = "✓" if is_payment else "−"
            text = "Payment added for" if is_payment else "Clawback added for"
            money_class = "money-green" if is_payment else "money-red"
            when = str(r.get("ts", "")).replace("T", " ")[:16]
            recent_parts.append(f'<div class="recent-row"><div class="{circle}">{icon}</div><div>{text} {r.get("rep", "")}</div><div class="{money_class}">{fmt_money(r.get("amount", 0))}</div><div class="small-muted">{when}</div></div>')
        recent_html = "".join(recent_parts)

    st.markdown(
        dedent(f"""
        <div class="bottom-grid">
          <div class="card">
            <div class="box-title">Payment & Clawback Summary (YTD)</div>
            <div class="box-line"></div>
            <div class="summary-head"><div>Metric</div><div>Amount (USD)</div></div>
            <div class="summary-row row-green"><div>💵 &nbsp; Total Eligible Commission (After Clawbacks)</div><div>{fmt_money(totals['eligible'])}</div></div>
            <div class="summary-row row-soft"><div>📊 &nbsp; Total OTC</div><div>{fmt_money(sum(OTC_CONFIG.values()))}</div></div>
            <div class="summary-row row-soft"><div>🏆 &nbsp; Total Multi-Year Bonus</div><div>{fmt_money(totals['bonus'])}</div></div>
            <div class="summary-row row-red"><div>🔁 &nbsp; Total Clawbacks</div><div>{fmt_money(total_clawback)}</div></div>
            <div class="summary-row row-soft"><div>💳 &nbsp; Total Commission Paid</div><div>{fmt_money(total_paid)}</div></div>
          </div>
          <div class="card">
            <div style="display:flex;justify-content:space-between;align-items:center;"><div class="box-title">Recent Payments & Clawbacks</div><div class="view-all">View All</div></div>
            <div class="box-line"></div>{recent_html}
          </div>
        </div>
        """),
        unsafe_allow_html=True,
    )


def get_dynamic_explanation(summary):
    attainment = float(summary.get("attainment", 0.0) or 0.0)
    gap = float(summary.get("gap_to_full", 0.0) or 0.0)
    if attainment < 0.2:
        return (
            "<b style='color:#d32f2f;'>Your commission is LOCKED</b><br>"
            "• You are below 20% attainment<br>"
            "• No commission payout yet<br>"
            "• Focus on closing deals to cross 20%<br>"
        )
    if attainment < 0.8:
        return (
            "<b style='color:#f57c00;'>You are in RAMPING zone</b><br>"
            "• Partial commission is being paid<br>"
            "• Every deal increases payout %<br>"
            f"• You need ${gap:,.0f} more to unlock full payout<br>"
        )
    return (
        "<b style='color:#2e7d32;'>Full commission UNLOCKED</b><br>"
        "• You crossed 80% of quota<br>"
        "• You earn 100% commission on all deals<br>"
        "• Keep pushing for maximum earnings<br>"
    )
@st.cache_data(show_spinner=False)
def fetch_clawback_projects_last_3_years(dataverse_token):
    start_year = CURRENT_YEAR - 3
    start = f"{start_year}-01-01T00:00:00Z"
    end = f"{CURRENT_YEAR}-12-31T23:59:59Z"

    headers = {
        "Authorization": f"Bearer {dataverse_token}",
        "Accept": "application/json",
        "Prefer": 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
    }

    url = (
        f"{DATAVERSE_URL}/api/data/v9.2/opportunities"
        "?$select=opportunityid,name,actualvalue,actualclosedate,_ownerid_value"
        f"&$filter=statecode eq 1 and actualclosedate ge {start} and actualclosedate le {end}"
    )

    r = requests.get(url, headers=headers)
    r.raise_for_status()

    df = pd.DataFrame(r.json().get("value", []))

    if df.empty:
        return df

    df["Opportunity Id"] = df["opportunityid"]
    df["Project Name"] = df["name"].fillna("")
    df["CRM Amount"] = pd.to_numeric(df["actualvalue"], errors="coerce").fillna(0.0)
    df["Sales Rep"] = df.get(
        "_ownerid_value@OData.Community.Display.V1.FormattedValue", ""
    ).fillna("").astype(str)
    df["Close Date"] = pd.to_datetime(df["actualclosedate"], errors="coerce")

    return df.sort_values("Close Date", ascending=False)
def get_crm_projects_for_clawback(token, rep_name):
    try:
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json"
        }

        url = (
            f"{DATAVERSE_URL}"
            "/api/data/v9.2/opportunities?"
            "$select=name,estimatedvalue,createdon"
        )

        response = requests.get(url, headers=headers)

        if response.status_code != 200:
            return {}

        data = response.json().get("value", [])

        projects = {}

        for row in data:
            name = row.get("name", "Unknown Project")
            value = float(row.get("estimatedvalue", 0.0) or 0.0)
            created = row.get("createdon", "")[:10]

            label = f"{name} | {created} | ${value:,.2f}"

            projects[label] = {
                "deal_value": value
            }

        return projects

    except Exception as e:
        print("CRM Clawback Fetch Error:", e)
        return {}

def render_detail_view(df_all, selected_rep, role, signed_in_upn, dataverse_token):
    if not selected_rep:
        selected_rep = CRD_REPS[0]
        st.session_state["selected_rep"] = selected_rep
    render_page_header(f"{selected_rep} - Commission Dashboard ({CURRENT_YEAR})", "Detailed commission summary and deal-level records.", show_back=(role == "ADMIN"))
    rep_display = match_rep_name(df_all, selected_rep)
    rep_df = build_rep_df(df_all, rep_display)
    quarter = st.selectbox("Filter My Dashboard by Quarter", ["All", "Q1", "Q2", "Q3", "Q4"], index=0, key=f"detail_quarter_{selected_rep}")
    rep_df = filter_df_by_quarter(rep_df, quarter)
    summary = compute_summary(rep_df, selected_rep, quarter)

    c1, c2, c3, c4 = st.columns(4)
    with c1: render_kpi_card("Annual Quota Goal", fmt_money(summary["quota"]), "Assigned quota", "🎯", "kpi-purple")
    with c2: render_kpi_card("Total Quota Credit", fmt_money(summary["total_quota_credit"]), "YTD", "📈", "kpi-green")
    with c3: render_kpi_card("Eligible Comm YTD", fmt_money(summary["eligible_comm"]), "After clawbacks", "$", "kpi-blue")
    with c4: render_kpi_card("YTD Attainment", fmt_pct(summary["attainment"]), summary["payout_status"], "◔", "kpi-gold")

    if role == "ADMIN":
        st.subheader("Admin Actions")
        colp1, colp2, colp3 = st.columns([1.1, 1.1, 1.4])
        with colp1:
            new_payment = st.number_input("New Payment Amount", min_value=0.0, step=1000.0, key=f"payment_{selected_rep}")
            if st.button("Lock Payment", key=f"lock_payment_{selected_rep}"):
                if new_payment > 0:
                    add_payment(selected_rep, new_payment, signed_in_upn)
                    st.success("Payment saved successfully.")
                    st.rerun()
                else:
                    st.warning("Enter a payment amount greater than 0.")

    if role == "ADMIN":
        if st.button("↩️ Undo Last Payment", key=f"undo_pay_{selected_rep}"):
            delete_latest_payment(selected_rep)
            st.success("Last payment removed successfully.")
            st.rerun()
    
    with colp2:
        

        clawback_projects = get_crm_projects_for_clawback(
            dataverse_token,
            selected_rep
        )

    project_options = ["-- Select Project --"] + list(clawback_projects.keys())

    selected_project_key = st.selectbox(
        "Clawback Project",
        project_options,
        key=f"cb_project_{selected_rep}"
    )

    calculated_clawback = 0.0
    cb_project = ""

    if selected_project_key != "-- Select Project --":

        project = clawback_projects[selected_project_key]

        deal_value = float(project.get("deal_value", 0.0) or 0.0)
        cb_project = selected_project_key

        st.caption(f"CRM Project Amount: {fmt_money(deal_value)}")

        invoiced_value = st.number_input(
            "Actual / Invoiced Value",
            min_value=0.0,
            value=0.0,
            step=1000.0,
            key=f"invoiced_value_{selected_rep}"
        )

        calculated_clawback = max(
            (deal_value - invoiced_value) * float(summary.get("rate", 0.0)),
            0.0
        )

        st.text_input(
            "Clawback Amount",
            value=fmt_money(calculated_clawback),
            disabled=True,
            key=f"cb_amount_display_{selected_rep}"
        )
    if st.button("Save Clawback", key=f"save_cb_{selected_rep}"):

        if calculated_clawback > 0:

            add_clawback(
                selected_rep,
                cb_project,
                float(calculated_clawback),
                signed_in_upn
            )

            st.success("Clawback saved successfully.")
            st.rerun()

        else:
            st.warning("Clawback amount is zero.")


    else:
        st.info("Please select a clawback project.")
  
        with colp3:
            st.write("**Saved Totals**")
            st.write(f"Total Paid: {fmt_money(get_total_paid(selected_rep, quarter))}")
            st.write(f"Total Clawback: {fmt_money(get_total_clawback(selected_rep, quarter))}")

    if role == "ADMIN":
        if st.button("🗑️ Clear Test Clawbacks", key=f"clear_cb_{selected_rep}"):

            clear_clawbacks(selected_rep)

            st.success("Test clawbacks cleared successfully.")
            st.rerun()

        saved_clawback = abs(get_total_clawback(selected_rep))

        final_eligible_comm = max(
            summary["base_eligible_comm"]
            + summary["multi_year_bonus_comm"]
            - saved_clawback,
            0.0
        )

        payable_after_clawback = (
            final_eligible_comm
            * summary["payout_factor"]
        )

        remaining_after_payment = max(
            payable_after_clawback
            - get_total_paid(selected_rep),
            0.0
        )
    rows = [
        ("Annual Quota Goal", fmt_money(summary["quota"])),
        ("OTC", fmt_money(OTC_CONFIG.get(selected_rep, 0))),
        ("Total Quota Credit", fmt_money(summary["total_quota_credit"])),
        ("YTD Attainment %", fmt_pct(summary["attainment"])),
        ("Comm Payout Threshold Value", fmt_money(summary["threshold"])),
        ("Gap To Full Payout", fmt_money(summary["gap_to_full"])),
        ("Ramp Start %", "20.00%"),
        ("Ramp End % / Gate", "80.00%"),
        ("Payout Factor", fmt_pct(summary["payout_factor"])),
        ("Payout Status", summary["payout_status"]),
        ("Eligible Comm YTD", fmt_money(summary["base_eligible_comm"])),
        ("Multi-Year Bonus(>=3yr CSP)", fmt_money(summary["multi_year_bonus_comm"])),
        ("Clawback Adjustment", f"-{fmt_money(saved_clawback)}"),
        ("Final Eligible Comm YTD", fmt_money(final_eligible_comm)),
        ("Total Comm Payable YTD", fmt_money(payable_after_clawback)),
        ("Paid Amount (Manual)", fmt_money(get_total_paid(selected_rep))),
        ("Total Commission Remaining YTD", fmt_money(remaining_after_payment)),
    ]

    table_rows = "".join([f"<tr><td>{m}</td><td class='value'>{v}</td></tr>" for m, v in rows])
    explanation = get_dynamic_explanation(summary)
    st.markdown(
        f"""
        <div class="summary-logic-grid">
          <table class="summary-table compact">
            <thead><tr><th>Metric</th><th>Value</th></tr></thead>
            <tbody>{table_rows}</tbody>
          </table>
          <div class="logic-card">
            <h4>📌 Commission Logic</h4>
            {explanation}
            <b>Ramp Logic</b>
            <ul>
              <li>Start: 20% → payout begins (prorated)</li>
              <li>End: 80% → prorated payout</li>
              <li>End: >80% → full payout</li>
            </ul>
            <b>Payout Factor</b>
            <ul><li>Commission payout starts once this is &gt;= 100%</li></ul>
            <b>CSP Bonus (One time)</b>
            <ul><li>0.25% extra for ≥3 year deals</li></ul>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.subheader("Detailed Deals")
    if rep_df.empty:
        st.info(f"No deals found for this CRD in {CURRENT_YEAR}.")
        return
    display_df = rep_df.copy()
    display_df["Close Date"] = pd.to_datetime(display_df["actualclosedate"], errors="coerce").dt.strftime("%d-%b-%Y")
    rate = float(rep_config.get(selected_rep, {}).get("rate", 0.0) or 0.0)
    display_df["Potential Commission"] = display_df["Quota Credit"] * rate
    display_df["Grand Total Potential Commission"] = display_df["Potential Commission"] + display_df["CSP Bonus (Year1 Only)"]
    for col in ["Deal Value", "TCV", "Quota Credit", "Potential Commission", "CSP Bonus (Year1 Only)", "Grand Total Potential Commission"]:
        if col in display_df.columns:
            display_df[col] = display_df[col].map(fmt_money)
    cols = ["Close Date", "Client Name", "Workload", "Final Revenue Category", "Deal Value", "TCV", "Quota Credit", "Potential Commission", "CSP Bonus (Year1 Only)", "Grand Total Potential Commission"]
    cols = [c for c in cols if c in display_df.columns]
    display_df = display_df.sort_values("Close Date", ascending=False)
    if "actualclosedate" in display_df.columns:
        display_df = display_df.sort_values(
            by="actualclosedate",
            ascending=False
        )
    st.dataframe(display_df[cols], use_container_width=True, hide_index=True)


def render_records_page(title, records):
    render_page_header(title, "Persistent ledger records visible to all admins.", show_back=True)
    df = pd.DataFrame(records)
    if df.empty:
        st.info("No records available yet.")
    else:
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.download_button(f"Download {title}", df.to_csv(index=False).encode("utf-8"), file_name=f"{title.lower()}.csv", mime="text/csv")
def render_reports_page(df_all):
    st.markdown(
        f"""
        <h1 class="page-title">Commission Report - {CURRENT_YEAR}</h1>
        <div class="page-subtitle">FY{CURRENT_YEAR} quota attainment and commission report.</div>
        <div class="divider"></div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns([1, 1])

    with col1:
        report_rep = st.selectbox(
            "Select CRD",
            CRD_REPS,
            key="report_rep_select"
        )

    with col2:
        report_quarter = st.selectbox(
            "Select Quarter",
            ["All", "Q1", "Q2", "Q3", "Q4"],
            key="report_quarter_select"
        )

    rep_display = match_rep_name(df_all, report_rep)
    rep_df = build_rep_df(df_all, rep_display)
    rep_df = filter_df_by_quarter(rep_df, report_quarter)

    summary = compute_summary(rep_df, report_rep)

    if rep_df.empty:
        st.info("No report data available for this selection.")
        return

    report_df = rep_df.copy()

    report_df["Close Date"] = pd.to_datetime(
        report_df["actualclosedate"],
        errors="coerce"
    ).dt.strftime("%d-%b-%Y")

    report_df["Commission"] = (
        pd.to_numeric(report_df["Quota Credit"], errors="coerce").fillna(0)
        * float(rep_config.get(report_rep, {}).get("rate", 0))
    )

    report_df = report_df.sort_values(
        by="actualclosedate",
        ascending=False
    )

    final_cols = [
        "Client Name",
        "Close Date",
        "Workload",
        "Final Revenue Category",
        "Deal Value",
        "Quota Credit",
        "Commission",
    ]

    final_cols = [c for c in final_cols if c in report_df.columns]

    display_report = report_df[final_cols].copy()

    for col in ["Deal Value", "Quota Credit", "Commission"]:
        if col in display_report.columns:
            display_report[col] = display_report[col].apply(fmt_money)

    st.markdown("### Detailed Commission Report")
    st.dataframe(
        display_report,
        use_container_width=True,
        hide_index=True
    )

    st.markdown("### Report Summary")

    summary_report = pd.DataFrame([
        {"Metric": "Annual Quota Goal", "Amount": fmt_money(summary["quota"])},
        {"Metric": "Total Quota Credit", "Amount": fmt_money(summary["total_quota_credit"])},
        {"Metric": "Eligible Commission", "Amount": fmt_money(summary["base_eligible_comm"])},
        {"Metric": "Multi-Year Bonus", "Amount": fmt_money(summary["multi_year_bonus_comm"])},
        {"Metric": "Final Eligible Commission", "Amount": fmt_money(
             summary["base_eligible_comm"]
            + summary["multi_year_bonus_comm"]
            - abs(get_total_clawback(report_rep))
        )},
        {"Metric": "Clawback Adjustment", "Amount": f"-{fmt_money(abs(get_total_clawback(report_rep)))}"},
        {"Metric": "Commission Paid", "Amount": fmt_money(get_total_paid(report_rep))},
        {"Metric": "YTD Attainment %", "Amount": fmt_pct(summary["attainment"])},
        {"Metric": "Payout Status", "Amount": summary["payout_status"]},
    ])

    st.dataframe(
        summary_report,
        use_container_width=True,
        hide_index=True
    )

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        display_report.to_excel(writer, index=False, sheet_name="Detailed Report")
        summary_report.to_excel(writer, index=False, sheet_name="Summary")

    st.download_button(
        label="⬇ Export Report to Excel",
        data=output.getvalue(),
        file_name=f"{report_rep}_Commission_Report_{report_quarter}_{CURRENT_YEAR}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# =====================================================
# MAIN
# =====================================================
def main():
    inject_css()
    ensure_database()

    if st.session_state.get("logged_out"):
        st.markdown('<h1 class="page-title">Commission Calculator - 2026</h1><div class="page-subtitle">You are logged out.</div>', unsafe_allow_html=True)
        if st.button("Sign in with Microsoft"):
            st.session_state.pop("logged_out", None)
            st.rerun()
        st.stop()

    graph_token = get_token_for_scopes(GRAPH_SCOPES, "graph_token")
    signed_in_upn = get_logged_in_user_upn(graph_token)

    dataverse_token = get_token_for_scopes(
        DATAVERSE_SCOPES,
        "dataverse_token"
    )

    # =====================================================
    # ADD THIS BLOCK HERE
    # =====================================================

    if "dataverse_token" not in st.session_state:

        st.session_state["dataverse_token"] = get_token_for_scopes(
            DATAVERSE_SCOPES,
            "dataverse_token"
        )

    dataverse_token = st.session_state["dataverse_token"]

    # ====================================================


    if signed_in_upn in ROLE_CONFIG["admin"]:
        role = "ADMIN"
    elif signed_in_upn in ROLE_CONFIG["users"]:
        role = "USER"
        st.session_state["selected_rep"] = ROLE_CONFIG["users"][signed_in_upn]
        st.session_state["view"] = "detail"
    else:
        st.error("You are not authorized to access this application.")
        st.stop()

    if "view" not in st.session_state:
        st.session_state["view"] = "dashboard"
    if "selected_rep" not in st.session_state:
        st.session_state["selected_rep"] = None

    if requested_dashboard():
        st.session_state["selected_rep"] = None
        st.session_state["view"] = "dashboard"
        clear_query_params()

    requested_crd = get_requested_crd()
    if requested_crd in CRD_REPS:
        st.session_state["selected_rep"] = requested_crd
        st.session_state["view"] = "detail"

    render_sidebar(signed_in_upn, role)

    dataverse_token = get_token_for_scopes(DATAVERSE_SCOPES, "dataverse_token")
    try:
        df = fetch_opportunities(dataverse_token)
    except Exception as e:
        st.error(f"Could not load Dataverse opportunities: {e}")
        st.stop()

    if df.empty:
        st.warning(f"No WON opportunities found for {CURRENT_YEAR}.")
        st.stop()

    if st.session_state["view"] == "dashboard":
        render_admin_dashboard(df)

    elif st.session_state["view"] == "reports":
        render_reports_page(df)

    else:
        render_detail_view(df, st.session_state.get("selected_rep"), role, signed_in_upn, dataverse_token)


if __name__ == "__main__":
    main()
