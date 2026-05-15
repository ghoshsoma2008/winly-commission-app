import streamlit as st
import pandas as pd
import msal
import requests
import os
import json
from datetime import datetime, date
import math
from urllib.parse import quote_plus

# =====================================================
# APP CONFIG
# =====================================================
st.set_page_config(page_title="Commission Calculator - 2026", layout="wide")
st.title("Commission Calculator - 2026")

CURRENT_YEAR = 2026

TENANT_ID = "85f66ea0-8fe4-48b9-a1a7-8633937d534a"
CLIENT_ID = "c2859476-c750-4a48-8bdd-b66e5dbaa732"
DATAVERSE_URL = "https://netwoveninc.crm.dynamics.com"
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

GRAPH_SCOPES = ["User.Read"]
DATAVERSE_SCOPES = [f"{DATAVERSE_URL}/user_impersonation"]

# =====================================================
# ROLE CONFIG (SECURE ALLOW LIST BY SIGNED-IN UPN)
# =====================================================
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

# Admin landing page shows only these 4 CRDs
CRD_REPS = ["Nicholas Simas", "Angira Dey", "Mandeep Nagpal", "Chris Wilkinson"]

# =====================================================
# COMMISSION CONFIG (RATE + QUOTA)
# =====================================================
rep_config = {
    "Angira Dey": {"rate": 0.006250, "quota": 8_000_000},
    "Nicholas Simas": {"rate": 0.008, "quota": 7_500_000},
    "Chris Wilkinson": {"rate": 0.007200, "quota": 5_000_000},
    "Mandeep Nagpal": {"rate": 0.0087600, "quota": 1_500_000},
}

# OTC (One-Time Credit) per CRD
OTC_CONFIG = {
    "Angira Dey": 50000,
    "Nicholas Simas": 60000,
    "Chris Wilkinson": 36000,
    "Mandeep Nagpal": 8420,
}

# =====================================================
# ELIGIBILITY RULES
# =====================================================
ELIGIBILITY_MAP = {
    "CSP": 0.25,
    "Consulting": 1.00,
    "Staffing 1 (>=20% markup)": 1.00,
    "Staffing 2 (10-19% markup)": 0.70,
    "Staffing 3 (<10% markup)": 0.00,
    "Govern 365": 1.00,
}

# =====================================================
# CSP MULTI-YEAR (LOGIC ONLY)
# =====================================================
CSP_TCV_COL = "nw_totalcontractvaluenw"
CSP_START_COL = "new_startdate"
CSP_END_COL = "new_enddate"

# =====================================================
# MSAL PERSISTENT CACHE
# =====================================================
CACHE_PATH = os.path.join(os.getcwd(), ".msal_cache.json")


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
        err = (flow or {}).get("error", "unknown_error")
        desc = (flow or {}).get("error_description", str(flow))
        st.error(f"Failed to start device login flow: {err}\n\n{desc}")
        st.stop()

    st.info(flow.get("message", "Complete the device login shown above."))
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        err = result.get("error", "unknown_error")
        desc = result.get("error_description", str(result))
        st.error(f"Authentication failed: {err}\n\n{desc}")
        st.stop()

    _save_cache(cache)
    return result["access_token"]


def get_token_for_scopes(app, cache, scopes):
    accounts = app.get_accounts()
    if accounts:
        silent = app.acquire_token_silent(scopes, account=accounts[0])
        if silent and "access_token" in silent:
            _save_cache(cache)
            return silent["access_token"]
    return acquire_token_device_flow(app, cache, scopes)


def get_logged_in_user_upn(graph_token):
    headers = {"Authorization": f"Bearer {graph_token}"}
    r = requests.get("https://graph.microsoft.com/v1.0/me", headers=headers)
    r.raise_for_status()
    return r.json()["userPrincipalName"].lower()


st.markdown("""
<style>

/* ✅ ONLY affect CRD detail page section */
.crd-detail-ui [data-testid="stNumberInput"] {
    max-width: 320px !important;
}

.crd-detail-ui input {
    padding: 6px 8px !important;
    height: 34px !important;
    font-size: 14px;
}

/* ✅ Shrink expander to text width */
.crd-detail-ui div.streamlit-expander {
    width: fit-content !important;
    max-width: fit-content !important;
    border-radius: 8px;
}

/* ✅ Reduce expander header padding */
.crd-detail-ui div.streamlit-expanderHeader {
    padding: 4px 10px !important;
    font-weight: 600;
}

/* ✅ Remove annoying full-width divider look */
.crd-detail-ui hr {
    display: none;
}

</style>
""", unsafe_allow_html=True)




# =====================================================
# FETCH DATAVERSE OPPORTUNITIES (Won, 2026 only)
# =====================================================
@st.cache_data(show_spinner=False)
def fetch_all_won_opportunities(dataverse_token):
    
    headers = {
        "Authorization": f"Bearer {dataverse_token}",
        "Accept": "application/json",
        "Prefer": 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
    }

    url = (
        f"{DATAVERSE_URL}/api/data/v9.2/opportunities"
        "?$select=opportunityid,name,actualvalue,actualclosedate,_ownerid_value,_nw_workload_value"
        "&$filter=statecode eq 1"   # ✅ ONLY WON (NO DATE FILTER)
    )

    r = requests.get(url, headers=headers)
    r.raise_for_status()

    rows = r.json().get("value", [])
    df = pd.DataFrame(rows)

    if df.empty:
        return df

    df["Opportunity Id"] = df.get("opportunityid")
    df["Client Name"] = df.get("name")
    df["Deal Value"] = pd.to_numeric(df.get("actualvalue"), errors="coerce").fillna(0.0)

    df["Sales Rep"] = df.get("_ownerid_value@OData.Community.Display.V1.FormattedValue", "").fillna("").astype(str)

    df["actualclosedate"] = pd.to_datetime(df.get("actualclosedate"), errors="coerce")

    return df



@st.cache_data(show_spinner=False)
def fetch_opportunities(dataverse_token):
    headers = {
        "Authorization": f"Bearer {dataverse_token}",
        "Accept": "application/json",
        "Prefer": 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
    }

    start = f"{CURRENT_YEAR}-01-01T00:00:00Z"
    end = f"{CURRENT_YEAR}-12-31T23:59:59Z"

    # IMPORTANT: use '&$filter=' not '&$filter='
    url = (
        f"{DATAVERSE_URL}/api/data/v9.2/opportunities"
        "?$select=opportunityid,name,actualvalue,actualclosedate,_ownerid_value,_nw_workload_value,"
        f"{CSP_TCV_COL},{CSP_START_COL},{CSP_END_COL}"
        f"&$filter=statecode eq 1 and actualclosedate ge {start} and actualclosedate le {end}"
    )

    r = requests.get(url, headers=headers)
    r.raise_for_status()

    payload = r.json()
    rows = payload.get("value", [])
    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df["Opportunity Id"] = df.get("opportunityid")
    df["Client Name"] = df.get("name")
    df["Deal Value"] = pd.to_numeric(df.get("actualvalue"), errors="coerce").fillna(0.0)

    df["Sales Rep"] = df.get("_ownerid_value@OData.Community.Display.V1.FormattedValue", "").fillna("").astype(str)
    df["Workload"] = df.get("_nw_workload_value@OData.Community.Display.V1.FormattedValue", "").fillna("").astype(str)

    # CSP fields
    if CSP_TCV_COL in df.columns:
        df[CSP_TCV_COL] = pd.to_numeric(df[CSP_TCV_COL], errors="coerce")
    else:
        df[CSP_TCV_COL] = pd.NA

    if CSP_START_COL in df.columns:
        df[CSP_START_COL] = pd.to_datetime(df[CSP_START_COL], errors="coerce")
    else:
        df[CSP_START_COL] = pd.NaT

    if CSP_END_COL in df.columns:
        df[CSP_END_COL] = pd.to_datetime(df[CSP_END_COL], errors="coerce")
    else:
        df[CSP_END_COL] = pd.NaT

    df = df[df["Sales Rep"].str.strip() != ""].copy()
    return df


@st.cache_data(show_spinner=False)
def load_staffing_rates():
    try:
        path = "rates.xlsx"
        df = pd.read_excel(path, engine="openpyxl")
        df["Bill Rate"] = pd.to_numeric(df["Bill Rate"], errors="coerce")
        df["Pay Rate"] = pd.to_numeric(df["Pay Rate"], errors="coerce")
        return df
    except Exception as e:
        st.warning("Excel file not loaded. Using empty data.")
        return pd.DataFrame()


import re


def lookup_staffing_markup(client_name: str, rates_df: pd.DataFrame):
    if not client_name or rates_df is None or rates_df.empty:
        return None

    cn = str(client_name).strip()
    parts = [p.strip() for p in cn.split("-")]
    person_part = parts[1] if len(parts) >= 2 else cn

    person_part = re.sub(r"\b(renewal|renew|extension|support|project|phase)\b", "", person_part, flags=re.I)
    person_part = re.sub(r"\b(20\d{2})\b", "", person_part)
    person_part = re.sub(r"\s+", " ", person_part).strip()

    if "Resource" in rates_df.columns and person_part:
        m = rates_df[rates_df["Resource"].astype(str).str.contains(person_part, case=False, na=False)]
        if not m.empty:
            row = m.iloc[0]
            bill = pd.to_numeric(row.get("Bill Rate"), errors="coerce")
            pay = pd.to_numeric(row.get("Pay Rate"), errors="coerce")
            if pd.notna(bill) and pd.notna(pay) and pay != 0:
                return (bill - pay) / pay

    if "Project Name" in rates_df.columns:
        company_part = parts[0] if len(parts) >= 2 else ""
        candidates = [person_part, company_part, cn]
        for key in [k for k in candidates if k]:
            m = rates_df[rates_df["Project Name"].astype(str).str.contains(key, case=False, na=False)]
            if not m.empty:
                row = m.iloc[0]
                bill = pd.to_numeric(row.get("Bill Rate"), errors="coerce")
                pay = pd.to_numeric(row.get("Pay Rate"), errors="coerce")
                if pd.notna(bill) and pd.notna(pay) and pay != 0:
                    return (bill - pay) / pay

    return None


# =====================================================
# RULES + HELPERS
# =====================================================
def derive_category(workload, client_name):
    txt = f"{workload} {client_name}".lower()
    if "nintex" in txt:
        return "Product"
    elif "business apps" in txt:
        return "Consulting"
    elif "csp" in txt:
        return "CSP"
    elif "consult" in txt:
        return "Consulting"
    elif "staffing" in txt:
        return "Staffing 2 (10-19% markup)"
    elif "govern" in txt:
        return "G365"

    # ✅ DEFAULT FOR EVERYTHING ELSE
    return "Consulting"



def match_rep_name(df: pd.DataFrame, target_name: str) -> str:
    if df is None or df.empty:
        return target_name
    target = target_name.lower().strip()
    parts = [p for p in target.split(" ") if p]
    candidates = []
    for rep in sorted(df["Sales Rep"].dropna().unique()):
        rep_l = str(rep).lower()
        if all(p in rep_l for p in parts):
            candidates.append(rep)
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) > 1:
        for c in candidates:
            if str(c).lower().strip() == target:
                return c
        return candidates[0]
    return target_name


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


def parse_markup(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return pd.NA
    if isinstance(x, str):
        s = x.strip()
        if s.endswith("%"):
            s = s[:-1].strip()
            return float(s) / 100.0
        return float(s)
    return float(x)


def staffing_tier(markup_ratio):
    if markup_ratio is None or (isinstance(markup_ratio, float) and pd.isna(markup_ratio)):
        return pd.NA
    m = float(markup_ratio)
    if m >= 0.20:
        return "Staffing 1 (>=20% markup)"
    if m >= 0.10:
        return "Staffing 2 (10-19% markup)"
    return "Staffing 3 (<10% markup)"


def compute_csp_years(start_date, end_date):
    if pd.isna(start_date) or pd.isna(end_date):
        return 1
    try:
        days = (end_date - start_date).days
        if days <= 0:
            return 1
        return max(1, int(math.ceil(days / 365.0)))
    except Exception:
        return 1


def build_rep_df(all_df: pd.DataFrame, rep_display_name: str) -> pd.DataFrame:
    rep_df = all_df[all_df["Sales Rep"] == rep_display_name].copy()
    # Ensure TCV column exists for CSP deals
    if CSP_TCV_COL in rep_df.columns:
        rep_df["TCV"] = pd.to_numeric(rep_df[CSP_TCV_COL], errors="coerce").fillna(0.0)
    else:
        rep_df["TCV"] = 0.0
    if rep_df.empty:
        return pd.DataFrame(columns=["Client Name", "Final Revenue Category", "Deal Value", "Quota Credit", "Workload"])

    # Add Close Month column for filtering
    rep_df["actualclosedate"] = pd.to_datetime(rep_df["actualclosedate"], errors="coerce")
    rep_df["Close Month"] = rep_df["actualclosedate"].dt.strftime('%B')

    rep_df["Derived Revenue Category"] = rep_df.apply(
        lambda r: derive_category(r.get("Workload", ""), r.get("Client Name", "")), axis=1
    )

    rates_df = load_staffing_rates()

    # Markup for staffing
    rep_df["Markup"] = rep_df.apply(
        lambda r: lookup_staffing_markup(r.get("Client Name", ""), rates_df)
        if "staff" in str(r.get("Workload", "")).lower()
        else None,
        axis=1,
    )

    rep_df["Markup_num"] = rep_df["Markup"].apply(parse_markup)

    # Final category (staffing tiers override)
    rep_df["Final Revenue Category"] = rep_df["Derived Revenue Category"]
    mask_staff = rep_df["Workload"].astype(str).str.contains("staff", case=False, na=False)
    rep_df.loc[mask_staff, "Final Revenue Category"] = (
        rep_df.loc[mask_staff, "Markup_num"].apply(staffing_tier).fillna(rep_df.loc[mask_staff, "Final Revenue Category"])
    )

    # Eligibility + quota credit
    rep_df["Eligibility %"] = rep_df["Final Revenue Category"].map(ELIGIBILITY_MAP).fillna(0.0)
    rep_df["Quota Credit"] = rep_df["Deal Value"] * rep_df["Eligibility %"]

    # CSP multi-year fields
    for col, default in [(CSP_TCV_COL, pd.NA), (CSP_START_COL, pd.NaT), (CSP_END_COL, pd.NaT)]:
        if col not in rep_df.columns:
            rep_df[col] = default

    rep_df[CSP_TCV_COL] = pd.to_numeric(rep_df[CSP_TCV_COL], errors="coerce")
    rep_df[CSP_START_COL] = pd.to_datetime(rep_df[CSP_START_COL], errors="coerce")
    rep_df[CSP_END_COL] = pd.to_datetime(rep_df[CSP_END_COL], errors="coerce")

    def _csp_years_row(r):
        if str(r.get("Final Revenue Category", "")).strip() != "CSP":
            return 1
        arr = float(pd.to_numeric(r.get("Deal Value", 0), errors="coerce") or 0.0)
        tcv = r.get(CSP_TCV_COL, pd.NA)
        if arr <= 0 or pd.isna(tcv) or float(tcv) <= 0:
            return 1
        if pd.isna(r.get(CSP_START_COL)) or pd.isna(r.get(CSP_END_COL)):
            return 1
        return compute_csp_years(r.get(CSP_START_COL), r.get(CSP_END_COL))

    rep_df["CSP Years"] = rep_df.apply(_csp_years_row, axis=1)

    rep_df["CSP Bonus (Year1 Only)"] = rep_df.apply(
        lambda r: float(r.get(CSP_TCV_COL) or 0.0) * 0.0025
        if (str(r.get("Final Revenue Category", "")).strip() == "CSP" and pd.notna(r.get(CSP_TCV_COL)) and int(r.get("CSP Years", 1) or 1) >= 3)
        else 0.0,
        axis=1,
    )

    return rep_df


def compute_summary(rep_df: pd.DataFrame, rep_name: str):
    cfg = rep_config.get(rep_name, {"rate": 0.0, "quota": 0.0})
    rate = float(cfg.get("rate", 0.0) or 0.0)
    quota = float(cfg.get("quota", 0.0) or 0.0)

    total_quota_credit = float(rep_df["Quota Credit"].sum()) if not rep_df.empty else 0.0
    attainment = (total_quota_credit / quota) if quota else 0.0

    if attainment < 0.2:
        payout_factor = 0.0
    elif attainment >= 0.8:
        payout_factor = 1.0
    else:
        payout_factor = (attainment - 0.2) / 0.5

    payout_status = "LOCKED" if payout_factor == 0 else ("RAMPING" if payout_factor < 1 else "UNLOCKED")

    base_eligible_comm = total_quota_credit * rate if rate > 0 else 0.0

    bonus_comm = 0.0
    if rate > 0 and (not rep_df.empty) and ("CSP Bonus (Year1 Only)" in rep_df.columns):
        bonus_comm = float(
    pd.to_numeric(rep_df["CSP Bonus (Year1 Only)"], errors="coerce")
      .fillna(0.0)
      .sum()
)

    eligible_comm = base_eligible_comm + bonus_comm

    paid_comm = eligible_comm * payout_factor
    remaining_comm = eligible_comm - paid_comm

    return {
        "quota": quota,
        "total_quota_credit": total_quota_credit,
        "attainment": attainment,
        "threshold": quota * 0.8 if quota else 0.0,
        "gap_to_full": max(0.0, quota - total_quota_credit) if quota else 0.0,
        "payout_factor": payout_factor,
        "payout_status": payout_status,
        "eligible_comm": eligible_comm,
        "paid_comm": paid_comm,
        "remaining_comm": remaining_comm,
        "multi_year_bonus_comm": bonus_comm,
        "rate": rate,
    }


# =====================================================
# STYLES (borders + hyperlink look + summary table)
# =====================================================
st.markdown(
    """
    <style>
      table.dashboard-table { width: 100%; border-collapse: collapse; margin-top: 12px; font-size: 14px; }
      table.dashboard-table th { background:#1F4E79; color:#fff; font-weight:700; padding:10px; border:1px solid #D0D5DD; text-align:left; white-space:nowrap; }
      table.dashboard-table td { padding:10px; border:1px solid #D0D5DD; background:#fff; vertical-align:middle; }
      table.dashboard-table tbody tr:hover td { background:#F8FAFC; }
      a.crd-anchor { color:#1a73e8; font-weight:600; text-decoration:none; }
      a.crd-anchor:hover { text-decoration:underline; color:#1257b7; }

      table.summary-table { width:520px; border-collapse:collapse; margin-top:6px; font-size:14px; }
      table.summary-table th { background:#1F4E79; color:#fff; font-weight:700; padding:10px; border:1px solid #D0D5DD; text-align:left; }
      table.summary-table td { padding:8px 10px; border:1px solid #E5E7EB; background:#fff; }
      table.summary-table td.metric { width:45%; white-space:nowrap; font-weight:600; }
      table.summary-table td.value  { width:55%; }
    </style>
    """,
    unsafe_allow_html=True,
)


# =====================================================
# QUERY PARAM HELPERS
# =====================================================
def get_query_params():
    try:
        return dict(st.query_params)
    except Exception:
        return st.experimental_get_query_params()


def clear_query_params():
    try:
        st.query_params.clear()
    except Exception:
        st.experimental_set_query_params()


# =====================================================
# SESSION STATE
# =====================================================
if "view" not in st.session_state:
    st.session_state["view"] = "dashboard"
if "selected_rep" not in st.session_state:
    st.session_state["selected_rep"] = None

# =====================================================
# PERSISTENT COMMISSION PAYMENT LEDGER (ADMIN LOCK)
# =====================================================
LEDGER_PATH = os.path.join(os.getcwd(), 'commission_payments_ledger.json')

def _load_payment_ledger():
    """Load append-only payment ledger from local JSON file."""
    if os.path.exists(LEDGER_PATH):
        try:
            with open(LEDGER_PATH, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}
    return {}

def _save_payment_ledger(ledger: dict):
    try:
        with open(LEDGER_PATH, 'w', encoding='utf-8') as f:
            json.dump(ledger, f, indent=2)
    except Exception:
        pass

def _get_paid_total(ledger: dict, rep_name: str) -> float:
    entries = ledger.get(rep_name, []) or []
    total = 0.0
    for e in entries:
        try:
            total += float(e.get('amount', 0.0) or 0.0)
        except Exception:
            pass
    return float(total)



# =====================================================
# LOGIN + DATA LOAD
# =====================================================
app, cache = get_msal_app()
graph_token = get_token_for_scopes(app, cache, GRAPH_SCOPES)
signed_in_upn = get_logged_in_user_upn(graph_token)
st.caption(f"Signed in as: {signed_in_upn}")

if signed_in_upn in ROLE_CONFIG["admin"]:
    role = "ADMIN"
elif signed_in_upn in ROLE_CONFIG["users"]:
    role = "USER"
    st.session_state["selected_rep"] = ROLE_CONFIG["users"][signed_in_upn]
    st.session_state["view"] = "detail"
else:
    st.error("You are not authorized to access this application.")
    st.stop()

dataverse_token = get_token_for_scopes(app, cache, DATAVERSE_SCOPES)
df = fetch_opportunities(dataverse_token)
if df.empty:
    st.warning(f"No WON opportunities found for {CURRENT_YEAR}.")
    st.stop()



# =====================================================
# NAVIGATION
# =====================================================
params = get_query_params()
requested = None
if "crd" in params:
    v = params["crd"]
    if isinstance(v, list):
        v = v[0] if v else None
    requested = v

if role == "ADMIN" and requested and requested in CRD_REPS:
    st.session_state["selected_rep"] = requested
    st.session_state["view"] = "detail"

def get_dynamic_explanation(summary):
    attainment = summary.get("attainment", 0)
    status = summary.get("payout_status", "")
    gap = summary.get("gap_to_full", 0)

    if attainment < 0.2:
        return (
            "<b style='color:#d32f2f;'>Your commission is LOCKED</b><br>"
            "• You are below 20% attainment<br>"
            "• No commission payout yet<br>"
            "• Focus on closing deals to cross 20%<br>"
        )
    elif attainment < 0.8:
        return (
            "<b style='color:#f57c00;'>You are in RAMPING zone</b><br>"
            "• Partial commission is being paid<br>"
            "• Every deal increases payout %<br>"
            f"• You need ${gap:,.0f} more to unlock full payout<br>"
        )
    else:
        return (
            "<b style='color:#2e7d32;'>Full commission UNLOCKED</b><br>"
            "• You crossed 80% of quota<br>"
            "• You earn 100% commission on all deals<br>"
            "• Keep pushing for maximum earnings<br>"
        )

    
# =====================================================
# ✅ ADMIN DASHBOARD (WITH CLICKABLE LINK)
# =====================================================

def render_admin_dashboard(df_all: pd.DataFrame):

    st.subheader(f"Admin Dashboard - CRD Summary ({CURRENT_YEAR})")
    st.caption("Click a CRD name to open the detailed commission view.")

    rows_html = []

    for crd in CRD_REPS:

        rep_display = match_rep_name(df_all, crd)
        rep_df = build_rep_df(df_all, rep_display)
        summary = compute_summary(rep_df, crd)

        # ✅ ✅ THIS IS THE BLUE CLICKABLE LINK (FROM V3)
        href = f"?crd={quote_plus(crd)}"

        rows_html.append(
            "<tr>"
            f"<td><a class='crd-anchor' href='{href}'>{crd}</a></td>"
            f"<td>{fmt_money(summary['quota'])}</td>"
            f"<td>{fmt_money(summary['total_quota_credit'])}</td>"
            f"<td>{fmt_money(summary['eligible_comm'])}</td>"
            f"<td>{fmt_pct(summary['attainment'])}</td>"
            f"<td>{summary['payout_status']}</td>"
            "</tr>"
        )

    table_html = (
        "<table class='dashboard-table'>"
        "<thead><tr>"
        "<th>CRD Name</th>"
        "<th>Annual Quota Goal</th>"
        "<th>Total Quota Credit</th>"
        "<th>Eligible Comm YTD</th>"
        "<th>YTD Attainment %</th>"
        "<th>Payout Status</th>"
        "</tr></thead>"
        "<tbody>" + "".join(rows_html) + "</tbody></table>"
    )

    st.markdown(table_html, unsafe_allow_html=True)


# =====================================================
# RENDER: ADMIN DASHBOARD
# =====================================================
def render_detail_view(df_all: pd.DataFrame):
    df_all_years = fetch_all_won_opportunities(dataverse_token)
    selected_rep = st.session_state.get("selected_rep")
    rep_display = match_rep_name(df_all, selected_rep)

    # ✅ THIS LINE IS MISSING (THIS FIXES YOUR ERROR)
    rep_df = build_rep_df(df_all, rep_display)


    if not selected_rep:
        st.warning("Please select a CRD from the dashboard.")
        return
    
     # ✅ Load ledger
    if "payment_ledger" not in st.session_state:
        st.session_state["payment_ledger"] = _load_payment_ledger()

    ledger = st.session_state["payment_ledger"]

    if "manual_paid" not in st.session_state:
        st.session_state["manual_paid"] = {}

    otc_value = OTC_CONFIG.get(selected_rep, 0.0)

    manual_paid_value = _get_paid_total(ledger, selected_rep)
    st.session_state["manual_paid"][selected_rep] = float(manual_paid_value)

    st.subheader(f"{selected_rep} - Commission Summary ({CURRENT_YEAR})")

    rep_display = match_rep_name(df_all, selected_rep)
    rep_df_all = df_all_years[df_all_years["Sales Rep"] == rep_display].copy()


    month_options = ["All Months","January","February","March","April","May","June",
                     "July","August","September","October","November","December"]

    selected_months = st.multiselect("Filter By Month(s)", month_options, default=["All Months"])

    if "All Months" not in selected_months:
        rep_df = rep_df[rep_df["Close Month"].isin(selected_months)]

    summary = compute_summary(rep_df, selected_rep)
    rep_display = match_rep_name(df_all, selected_rep)

    # ✅ THIS IS MISSING (ADD THIS)
    rep_df = build_rep_df(df_all, rep_display)
    total_clawback_adjustment = 0.0

    


    # =============================
    # CRD DETAIL UI BLOCK
    # =============================
    st.markdown('<div class="crd-detail-ui">', unsafe_allow_html=True)

    # ✅ Admin payment locking
    if role == "ADMIN":

        st.markdown("#### Commission Paid (Admin Only) — Lock to prevent double payment")

        new_payment = st.number_input(
            "Enter new payment amount (will be added cumulatively)",
            min_value=0.0,
            value=0.0,
            step=1000.0,
            key=f"new_paid_input_{selected_rep}",
        )

        st.caption(f"Total Paid (Locked): {fmt_money(manual_paid_value)}")

    # ✅ PAYMENT HISTORY
    st.markdown("### 📜 Payment History")

    rep_ledger = ledger.get(selected_rep, [])
    history_rows = []

    if rep_ledger:

        for entry in rep_ledger:
            ts = entry.get("ts", "")
            amt = entry.get("amount", 0.0)
            user = entry.get("by", "")

            history_rows.append(
            f"<tr>"
            f"<td>{ts}</td>"
            f"<td>{fmt_money(amt)}</td>"
            f"<td>{user}</td>"
            f"</tr>"
            )

    history_html = (
        "<table class='summary-table'>"
        "<thead><tr>"
        "<th>Date</th><th>Amount</th><th>By</th>"
        "</tr></thead>"
        "<tbody>" + "".join(history_rows) + "</tbody></table>"
    )

    st.markdown(history_html, unsafe_allow_html=True)

    if not rep_ledger:
        st.info("No payment history available yet")

    # ✅ ✅ ALWAYS SHOW BUTTONS (IMPORTANT)
    col_pay_1, col_pay_2 = st.columns([1, 2])

    with col_pay_1:
        lock_clicked = st.button("Lock Payment", key=f"lock_pay_{selected_rep}")

    with col_pay_2:
        st.caption("Adds this payment to the ledger and locks it permanently.")

    # ✅ CLEAR BUTTON (NOW ALWAYS VISIBLE)
    if st.button("Clear Payment Log", key=f"clear_pay_{selected_rep}"):

        ledger[selected_rep] = []

        _save_payment_ledger(ledger)

        st.session_state["payment_ledger"] = ledger

        st.success("Payment log cleared successfully ✅")

        st.rerun()

    if lock_clicked:

            total_commission_actual_tmp = summary["eligible_comm"] + summary.get("multi_year_bonus_comm", 0.0)
            remaining_before = float(total_commission_actual_tmp) - float(manual_paid_value)

            if float(new_payment) <= 0:
                st.warning("Enter a payment amount greater than 0.")

            elif float(new_payment) - remaining_before > 0.000001:
                st.error(f"Payment exceeds remaining payable. Remaining: {fmt_money(remaining_before)}")

            else:
                entry = {
                    "ts": datetime.utcnow().isoformat(timespec="seconds") + "Z",
                    "amount": float(new_payment),
                    "by": signed_in_upn,
                }

                ledger.setdefault(selected_rep, []).append(entry)
                _save_payment_ledger(ledger)

                st.session_state["payment_ledger"] = ledger
                st.success("Payment locked and added to total paid.")
                st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

    # =====================================================
    # 🔁 REVERSALS
    # =====================================================
    if role == "ADMIN":

        st.markdown("### 🔁 Reversals (Claw-back)")

        for i in range(3):

            deal_options = rep_df_all["Opportunity Id"].dropna().astype(str).tolist()

            oppty_map = dict(zip(
                rep_df_all["Opportunity Id"].astype(str),
                rep_df_all["Client Name"] + " (" + rep_df_all["actualclosedate"].dt.year.astype(str) + ")"
            ))

            options = ["-- Select Project --"] + list(oppty_map.keys())

            selected_oppty = st.selectbox(
                f"Select Project {i+1}",
                options=options,
                format_func=lambda x: oppty_map.get(x, "-- Select Project --"),
                key=f"rev_proj_{selected_rep}_{i}"
            )

            if selected_oppty == "-- Select Project --":
                st.info("Please select a project to proceed")
                continue

            filtered_rows = rep_df_all[
                rep_df_all["Opportunity Id"].astype(str) == str(selected_oppty)
            ]

            row_data = filtered_rows.iloc[0]

            project_name = row_data["Client Name"]
            win_rev = float(row_data["Deal Value"])

            st.write(f"**Project:** {project_name}")
            st.write(f"**Win Revenue:** {fmt_money(win_rev)}")

            paused_bal = st.number_input(
                "Paused Balance",
                min_value=0.0,
                key=f"paused_{selected_rep}_{i}"
            )

            delta = paused_bal - win_rev
            clawback = delta * summary.get("rate", 0.0)

            total_clawback_adjustment += clawback

            st.metric("Delta", fmt_money(delta))
            st.metric("Claw-back", fmt_money(clawback))

            st.divider()

    # =====================================================
    # ✅ SUMMARY TABLE (ORIGINAL UI RESTORED)
    # =====================================================
    # ✅ Summary Table (existing layout)

    # ✅ Always calculate (for both admin & users)

    total_commission_actual = (
        summary["eligible_comm"]
        + summary.get("multi_year_bonus_comm", 0.0)
        + total_clawback_adjustment
    )
    total_paid = manual_paid_value  # ✅ already from ledger
    adjusted_remaining = total_commission_actual - total_paid

    screenshot_rows = [

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
        ("Eligible Comm YTD", fmt_money(summary["eligible_comm"])),
        ("Multi-Year Bonus Commission (>=3yr CSP)", fmt_money(summary.get("multi_year_bonus_comm", 0.0))),
        ("Clawback Adjustment", fmt_money(total_clawback_adjustment)),
        ("Total Eligible Comm YTD", fmt_money(total_commission_actual)),
        ("Total Comm Payable YTD", fmt_money(summary["paid_comm"])),
        ("Commission Paid (Manual Entry)", fmt_money(manual_paid_value)),
        ("Total Commission Remaining YTD", fmt_money(adjusted_remaining)),
    ]

    col1, col2 = st.columns([3, 2])

    with col1:
        rows = "".join([f"<tr><td class='metric'>{m}</td><td class='value'>{v}</td></tr>" for m, v in screenshot_rows])
        st.markdown(
            f"<table class='summary-table'><thead><tr>"
            f"<th style='width:45%'>Metric</th><th style='width:55%'>Value</th>"
            f"</tr></thead><tbody>{rows}</tbody></table>",
            unsafe_allow_html=True,
        )

    with col2:
        explanation = get_dynamic_explanation(summary)
        st.markdown(
            f"""
            <div style="border:1px solid #D0D5DD; border-radius:10px; padding:16px; background:#F9FAFB;">
              <h4 style="color:#1F4E79;">📌 Commission Logic</h4>
              {explanation}
              <b>Ramp Logic</b>
              <ul>
                <li>Start: 20% → payout begins (prorated)</li>
                <li>End: 80% → prorated payout </li>
                <li>End: >80% → full payout</li>
              </ul>
              <b>Payout Factor</b>
              <ul><li>Commission Payout starts once this is >=100%</li></ul>
              <b>CSP Bonus (One time)</b>
              <ul> <li>0.25% extra for ≥3 year deals</li></ul>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.subheader("Detailed Deals")
    if rep_df.empty:    
        st.info(f"No deals found for this CRD in {CURRENT_YEAR}.")


    # =====================================================
    # ✅ DETAILED TABLE (CLEAN)
    # =====================================================

    display_df = rep_df.copy()
    display_df = display_df.sort_values(by="actualclosedate", ascending=False)
    display_df["Close Date"] = display_df["actualclosedate"].dt.strftime("%d-%b-%Y")
    display_df["Deal Value"] = display_df["Deal Value"].map(fmt_money)
    display_df["Quota Credit"] = display_df["Quota Credit"].map(fmt_money)

    detail_cols = [
        "Close Date",
        "Client Name",
        "Workload",
        "Final Revenue Category",
        "Deal Value",
        "Quota Credit",
    ]

    detail_cols = [c for c in detail_cols if c in display_df.columns]
    
    st.dataframe(display_df[detail_cols], use_container_width=True)


# =====================================================
# ROUTER
# =====================================================
if role == "ADMIN":
    if st.session_state["view"] == "dashboard":
        render_admin_dashboard(df)
    else:
        render_detail_view(df)
else:
    render_detail_view(df)
