import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import json
import os
from datetime import datetime

st.set_page_config(
    page_title="TR Billing vs Cost Analysis",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Login credentials (override via Railway environment variables) ────────────
_ADMIN_USER = os.environ.get("ADMIN_USER", "admin")
_ADMIN_PASS = os.environ.get("ADMIN_PASS", "TRAdmin2026")
_SALES_USER = os.environ.get("SALES_USER", "sales")
_SALES_PASS = os.environ.get("SALES_PASS", "TRSales2026")

COST_COLS = [
    "Salary/Payroll", "Bench Payroll", "Marketing & BD",
    "HR, Admin, Mngt", "Software", "JobBoards",
    "Infrastructure", "Rent", "Travel", "Meals",
    "Fees & Charges", "Other Expenses",
]
EXCEL_PATH        = os.path.expanduser("~/Downloads/TR Billing V_S Cost Analysis -Feb 2026.xlsx")
BILLING_SHEETS    = {
    "Clients-Jan":      "Jan 2026",
    "Feb-25":           "Feb 2026",
    "Clients- March26": "Mar 2026",
}
_APP_DIR  = os.path.dirname(os.path.abspath(__file__))
# DATA_DIR: writable directory for all runtime data.
# Set DATA_DIR env var in production (e.g. Railway volume mount path).
# Falls back to the app directory for local development.
_DATA_DIR = os.environ.get("DATA_DIR", _APP_DIR)
os.makedirs(_DATA_DIR, exist_ok=True)

DATA_PATH = os.path.join(_DATA_DIR, "data.json")
EMP_PATH  = os.path.join(_DATA_DIR, "employees.json")

NON_BILLING_DEPTS = ["Marketing", "HR", "Admin", "Management"]
EMP_FIELDS = ["Name", "Designation", "Salary ($)", "Join Date", "Notes"]

EXPENSE_PATH   = os.path.join(_DATA_DIR, "expenses.json")
FX_RATES_PATH  = os.path.join(_DATA_DIR, "fx_rates.json")
CONFIG_PATH    = os.path.join(_DATA_DIR, "config.json")
UPLOADS_DIR    = os.path.join(_DATA_DIR, "uploads")

def _get_active_excel():
    """Return path to the most recently uploaded Excel (used for costs/salaries/expenses).
    Falls back to any xlsx found in uploads/, then EXCEL_PATH."""
    try:
        if os.path.exists(CONFIG_PATH):
            with open(CONFIG_PATH) as f:
                p = json.load(f).get("active_excel_path", "")
            if p and os.path.exists(p):
                return p
            # Stored path may be absolute from a different machine; try relative to app dir
            if p:
                rel = os.path.join(_APP_DIR, os.path.basename(p))
                if os.path.exists(rel):
                    return rel
                rel2 = os.path.join(_APP_DIR, "uploads", os.path.basename(p))
                if os.path.exists(rel2):
                    return rel2
    except Exception:
        pass
    # Auto-discover: pick the first xlsx in uploads/
    import glob as _glob
    for candidate in _glob.glob(os.path.join(_APP_DIR, "uploads", "*.xlsx")):
        return candidate
    return EXCEL_PATH

def _save_active_excel(path):
    cfg = {}
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH) as f:
            cfg = json.load(f)
    cfg["active_excel_path"] = path
    with open(CONFIG_PATH, "w") as f:
        json.dump(cfg, f, indent=2)
EXPENSE_CATS   = [
    "Rent", "Infrastructure", "Software", "Job Boards",
    "Travel", "Meals", "Fees & Charges",
    "Bonus & Incentive", "Marketing", "Other Expenses",
]
_ALL_MONTH_ABBRS = ["Dec","Nov","Oct","Sep","Aug","Jul","Jun","May","Apr","Mar","Feb","Jan"]

# Each abbreviation's possible prefix strings (all lowercase, used for startswith checks)
_MONTH_SEARCH_TERMS = {
    "Jan": ["jan"], "Feb": ["feb"], "Mar": ["mar"],  # "mar" matches "march..." too
    "Apr": ["apr"], "May": ["may"], "Jun": ["jun"],
    "Jul": ["jul"], "Aug": ["aug"], "Sep": ["sep"],
    "Oct": ["oct"], "Nov": ["nov"], "Dec": ["dec"],
}

def _find_sheet(xl_names, prefixes, month_abbr):
    """Find a sheet name whose prefix matches any of `prefixes` and whose suffix
    starts with any search term for `month_abbr`. Ignores spaces and case.
    e.g. 'Expenses-March26', 'Expenses-Feb26', 'Employee Summery- March26' all match."""
    terms = _MONTH_SEARCH_TERMS.get(month_abbr, [month_abbr.lower()])
    for name in xl_names:
        n = name.lower().replace(" ", "")
        for pfx in prefixes:
            p = pfx.lower().replace(" ", "")
            if n.startswith(p):
                suffix = n[len(p):].lstrip("-")
                if any(suffix.startswith(t) for t in terms):
                    return name
    return None

def _find_latest_expense_sheet(xl_names):
    """Return the expense sheet for the most recent month available."""
    for abbr in _ALL_MONTH_ABBRS:
        s = _find_sheet(xl_names, ["expenses"], abbr)
        if s:
            return s
    return None

OVERRIDE_PATH  = os.path.join(_DATA_DIR, "pnl_overrides.json")
SALES_PATH     = os.path.join(_DATA_DIR, "sales_persons.json")

# ── Data persistence ─────────────────────────────────────────────────────────

def load_json_data():
    if os.path.exists(DATA_PATH):
        with open(DATA_PATH) as f:
            return json.load(f)
    return {}

def save_json_data(data):
    with open(DATA_PATH, "w") as f:
        json.dump(data, f, indent=2)

def load_employees():
    if os.path.exists(EMP_PATH):
        with open(EMP_PATH) as f:
            return json.load(f)
    return {dept: [] for dept in NON_BILLING_DEPTS}

def save_employees(data):
    with open(EMP_PATH, "w") as f:
        json.dump(data, f, indent=2)

def load_expenses():
    if os.path.exists(EXPENSE_PATH):
        with open(EXPENSE_PATH) as f:
            return json.load(f)
    return []

def save_expenses(entries):
    with open(EXPENSE_PATH, "w") as f:
        json.dump(entries, f, indent=2)

def load_overrides():
    if os.path.exists(OVERRIDE_PATH):
        with open(OVERRIDE_PATH) as f:
            return json.load(f)
    return {}

def save_overrides(data):
    with open(OVERRIDE_PATH, "w") as f:
        json.dump(data, f, indent=2)

def load_fx_rates():
    defaults = {m: 91.0 for m in BILLING_SHEETS.values()}
    if os.path.exists(FX_RATES_PATH):
        with open(FX_RATES_PATH) as f:
            defaults.update(json.load(f))
    return defaults

def save_fx_rates(rates):
    with open(FX_RATES_PATH, "w") as f:
        json.dump(rates, f, indent=2)

def get_fx_rate(month):
    return load_fx_rates().get(month, _INR_RATE)

def load_sales_persons():
    if os.path.exists(SALES_PATH):
        with open(SALES_PATH) as f:
            return json.load(f)
    return {}

def save_sales_persons(data):
    with open(SALES_PATH, "w") as f:
        json.dump(data, f, indent=2)

_WANTED = [
    "Sr. No.", "Client Name", "Seat Name", "Billing ($)", "Billing (Rs.)",
    *COST_COLS, "Balance",
]

def get_actual_cost_totals_usd(bust=0):
    """
    Build cost totals in $ from actual live sources so that:
      sum(all cols) = Billing Payroll + Non-Billing Payroll + Total Expenses.

    Column mapping:
      Salary/Payroll  → total billing employee salaries (Employee Summery)
      Bench Payroll   → Payroll - Non Bill from Employee Summary sheet
      Marketing & BD  → Marketing dept payroll + Marketing overhead expenses
      HR, Admin, Mngt → HR + Admin + Management dept payroll
      Software        → Expenses sheet "Software"
      JobBoards       → Expenses sheet "Job Boards"
      Infrastructure  → Expenses sheet "Infrastructure"
      Rent            → Expenses sheet "Rent"
      Travel          → Expenses sheet "Travel"
      Meals           → Expenses sheet "Meals"
      Fees & Charges  → Expenses sheet "Fees & Charges"
      Other Expenses  → Expenses sheet "Other Expenses" + "Bonus & Incentive"
    """
    emp_data      = load_employees()
    active_xl     = _get_active_excel()
    xl_exp        = load_excel_expenses(active_xl, _bust=bust)
    sal_lookup    = load_billing_salaries(active_xl, _bust=bust)
    bench_payroll = load_bench_salaries(active_xl, _bust=bust)

    marketing_payroll     = sum(r.get("Salary ($)", 0) or 0 for r in emp_data.get("Marketing", []))
    hr_admin_mgmt_payroll = sum(
        r.get("Salary ($)", 0) or 0
        for dept in ("HR", "Admin", "Management")
        for r in emp_data.get(dept, [])
    )
    billing_payroll = sum(sal_lookup.values())

    exp_by_cat: dict = {}
    for r in xl_exp:
        cat = str(r.get("Category", ""))
        exp_by_cat[cat] = exp_by_cat.get(cat, 0) + float(r.get("Amount", 0))

    return {
        "Salary/Payroll":   billing_payroll,
        "Bench Payroll":    bench_payroll,
        "Marketing & BD":   marketing_payroll + exp_by_cat.get("Marketing", 0),
        "HR, Admin, Mngt":  hr_admin_mgmt_payroll,
        "Software":         exp_by_cat.get("Software", 0),
        "JobBoards":        exp_by_cat.get("Job Boards", 0),
        "Infrastructure":   exp_by_cat.get("Infrastructure", 0),
        "Rent":             exp_by_cat.get("Rent", 0),
        "Travel":           exp_by_cat.get("Travel", 0),
        "Meals":            exp_by_cat.get("Meals", 0),
        "Fees & Charges":   exp_by_cat.get("Fees & Charges", 0),
        "Other Expenses":   exp_by_cat.get("Other Expenses", 0) + exp_by_cat.get("Bonus & Incentive", 0),
    }


_CAT_MAP = {
    "Software & Tech": "Software",
    "Job Boards":       "Job Boards",
    "Other Expense":    "Other Expenses",
}

def _parse_expense_sheet(df):
    """Normalise an expense sheet into Category / Description / Amount records.

    Handles two layouts:
      3-col: Category | Description | Amount($)
      4-col: Category | Description | Amount(INR) | Amount($)   ← Jan format
    Finds the $ amount column by name when possible, else uses the last column.
    """
    df = df.copy()
    # Find the $ amount column: prefer a column whose header contains "$"
    amount_col = next(
        (c for c in df.columns if "$" in str(c)),
        df.columns[-1]   # fallback: last column
    )
    # Find category and description columns (first two)
    cat_col  = df.columns[0]
    desc_col = df.columns[1]

    out = pd.DataFrame({
        "Category":    df[cat_col].astype(str).str.strip(),
        "Description": df[desc_col].astype(str).str.strip(),
        "Amount":      pd.to_numeric(df[amount_col], errors="coerce").fillna(0),
    })
    out = out[out["Amount"] > 0].copy()
    out["Category"] = out["Category"].map(lambda c: _CAT_MAP.get(c, c))
    return out.to_dict(orient="records")

@st.cache_data(ttl=0)
def load_excel_expenses(path, _bust=0):
    """Load expenses from the most-recent month's sheet (used for cost totals)."""
    try:
        xl_names = pd.ExcelFile(path).sheet_names
        sheet = _find_latest_expense_sheet(xl_names)
        if sheet is None:
            return []
        return _parse_expense_sheet(pd.read_excel(path, sheet_name=sheet))
    except Exception:
        return []

@st.cache_data(ttl=0)
def load_expenses_for_month(path, month, _bust=0):
    """Load expenses for a specific month.
    Tries `path` first; if month sheet not found there, falls back to EXCEL_PATH."""
    abbr = month.split()[0][:3]  # e.g. "Mar"
    search_paths = [path] if path == EXCEL_PATH else [path, EXCEL_PATH]
    for p in search_paths:
        try:
            xl_names = pd.ExcelFile(p).sheet_names
            sheet = _find_sheet(xl_names, ["expenses"], abbr)
            if sheet:
                return _parse_expense_sheet(pd.read_excel(p, sheet_name=sheet))
        except Exception:
            continue
    return []


def load_excel_sheet(path, sheet):
    raw = pd.read_excel(path, sheet_name=sheet)
    # Drop helper/extra columns not in our standard set
    raw = raw.drop(columns=[c for c in raw.columns if c not in _WANTED], errors="ignore")
    # Add any missing columns as 0
    for col in _WANTED:
        if col not in raw.columns:
            raw[col] = None
    df = raw[_WANTED].copy()
    df = df.dropna(subset=["Sr. No."])
    df["Sr. No."] = pd.to_numeric(df["Sr. No."], errors="coerce")
    df = df.dropna(subset=["Sr. No."])
    df["Client Name"] = df["Client Name"].ffill()
    for col in ["Billing ($)", "Billing (Rs.)", *COST_COLS, "Balance"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df.reset_index(drop=True)

@st.cache_data(ttl=0)
def load_excel(path, _bust=0):
    sheets = {}
    if not os.path.exists(path):
        return sheets
    try:
        xl = pd.ExcelFile(path)
        for raw_name, display_name in BILLING_SHEETS.items():
            if raw_name in xl.sheet_names:
                sheets[display_name] = load_excel_sheet(path, raw_name)
    except Exception as e:
        st.error(f"Could not load Excel file: {e}")
    return sheets

def get_all_data():
    bust    = st.session_state.get("cache_bust", 0)
    sheets  = load_excel(_get_active_excel(), _bust=bust)
    sheets  = {k: v.copy() for k, v in sheets.items()}

    cost_totals_usd = get_actual_cost_totals_usd(bust=bust)
    cost_totals     = {k: v * _INR_RATE for k, v in cost_totals_usd.items()}
    total_cost_usd  = sum(cost_totals_usd.values())
    fx_rates        = load_fx_rates()

    # ── Merge any manually added custom records ────────────────────────────────
    # Excel (BILLING_SHEETS) is authoritative — skip data.json for those months
    excel_months = set(sheets.keys())
    custom = load_json_data()
    for month, rows in custom.items():
        if month in excel_months:
            continue
        sheets[month] = pd.DataFrame(rows)

    # ── Apply saved P&L overrides (Billing corrections + removals) ───────────
    overrides = load_overrides()
    for month, emp_overrides in overrides.items():
        if month not in sheets:
            continue
        df = sheets[month]
        # Remove seats explicitly moved off billing
        remove_list = emp_overrides.get("_remove_from_billing", [])
        if remove_list:
            df = df[~df["Seat Name"].isin(remove_list)].reset_index(drop=True)
        # Apply billing amount overrides
        for emp, vals in emp_overrides.items():
            if emp.startswith("_"):
                continue
            mask = df["Seat Name"] == emp
            if "Billing ($)" in vals:
                df.loc[mask, "Billing ($)"] = vals["Billing ($)"]
        sheets[month] = df

    # ── Inject cost totals + Balance for ALL months (runs last so every
    #    month — whether from Excel, custom JSON, or import — is covered) ───────
    for month, df in sheets.items():
        N = len(df)
        if N == 0:
            continue
        if "Billing ($)" not in df.columns:
            df["Billing ($)"] = 0.0
        fx_rate = fx_rates.get(month, _INR_RATE)
        df["Billing (Rs.)"] = df["Billing ($)"] * fx_rate
        for c in COST_COLS:
            df[c] = cost_totals.get(c, 0) / N
        df["Balance"] = df["Billing ($)"] - (total_cost_usd / N)
        sheets[month] = df

    return sheets

# ── Formatting helpers ────────────────────────────────────────────────────────

def fmt_inr(v):
    v = int(v)
    if abs(v) >= 1_00_000:
        return f"₹{v/1_00_000:.2f}L"
    if abs(v) >= 1_000:
        return f"₹{v/1_000:.1f}K"
    return f"₹{v}"

def color_balance(val):
    color = "green" if val >= 0 else "red"
    return f"color: {color}; font-weight: bold"

def _disp(usd_val, month="Jan 2026", dec=0):
    """Format a $ value in the active display currency ($ or ₹)."""
    if not st.session_state.get("show_inr", False):
        return f"${usd_val:,.{dec}f}"
    inr = usd_val * get_fx_rate(month)
    if abs(inr) >= 1_00_000:
        return f"₹{inr/1_00_000:.2f}L"
    if abs(inr) >= 1_000:
        return f"₹{inr/1_000:.1f}K"
    return f"₹{inr:,.{dec}f}"

def _currency_label(month="Jan 2026"):
    if st.session_state.get("show_inr", False):
        return f"₹  ($1 = ₹{get_fx_rate(month):.0f})"
    return "$"

# ── Pages ─────────────────────────────────────────────────────────────────────

def page_dashboard(data):
    months = list(data.keys())
    if not months:
        st.warning("No data available.")
        return

    # ── Hero header + month selector ───────────────────────────────────────────
    col_hero, col_month = st.columns([3, 1])
    with col_hero:
        st.markdown("""
<div class="dash-hero">
  <div class="dash-hero-title">Financial Overview</div>
  <div class="dash-hero-sub">Revenue, costs &amp; profitability at a glance</div>
</div>""", unsafe_allow_html=True)
    with col_month:
        st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
        sel_month = st.selectbox("", months, index=len(months) - 1, label_visibility="collapsed")

    df         = data[sel_month].copy()
    total_rev  = df["Billing ($)"].sum()
    total_cost = sum(get_actual_cost_totals_usd().values())
    total_bal  = df["Balance"].sum()
    profitable = int((df["Balance"] > 0).sum())
    loss_seats = int((df["Balance"] <= 0).sum())
    margin     = (total_bal / total_rev * 100) if total_rev else 0
    n_seats    = len(df)

    show_inr    = st.session_state.get("show_inr", False)
    rate        = get_fx_rate(sel_month) if show_inr else 1.0
    pfx         = "₹" if show_inr else "$"
    bal_color   = "#00d4aa" if total_bal >= 0 else "#ff5a00"
    bal_rgb     = "0,212,170" if total_bal >= 0 else "255,90,0"
    margin_color = "#00d4aa" if margin >= 0 else "#ff7a30"
    trend_arrow  = "▲" if total_bal >= 0 else "▼"

    fv = lambda v: _disp(v, sel_month)

    # ── KPI cards ──────────────────────────────────────────────────────────────
    st.markdown(f"""
<div class="dash-kpi-grid">
  <div class="dash-kpi-card">
    <div class="dash-kpi-icon" style="background:rgba(0,212,170,0.15);color:#00d4aa;">$</div>
    <div class="dash-kpi-body">
      <div class="dash-kpi-label">Total Revenue</div>
      <div class="dash-kpi-value" style="color:#00d4aa">{fv(total_rev)}</div>
      <div class="dash-kpi-sub">{n_seats} active seats</div>
    </div>
  </div>
  <div class="dash-kpi-card">
    <div class="dash-kpi-icon" style="background:rgba(162,155,254,0.15);color:#a29bfe;">≡</div>
    <div class="dash-kpi-body">
      <div class="dash-kpi-label">Total Cost</div>
      <div class="dash-kpi-value" style="color:#a29bfe">{fv(total_cost)}</div>
      <div class="dash-kpi-sub">Fully loaded</div>
    </div>
  </div>
  <div class="dash-kpi-card" style="border-color:rgba({bal_rgb},0.28)">
    <div class="dash-kpi-icon" style="background:rgba({bal_rgb},0.15);color:{bal_color};">{trend_arrow}</div>
    <div class="dash-kpi-body">
      <div class="dash-kpi-label">Net Balance</div>
      <div class="dash-kpi-value" style="color:{bal_color}">{fv(total_bal)}</div>
      <div class="dash-kpi-sub" style="color:{margin_color}">{margin:.1f}% margin</div>
    </div>
  </div>
  <div class="dash-kpi-card">
    <div class="dash-kpi-icon" style="background:rgba(184,244,88,0.12);color:#b8f458;">✓</div>
    <div class="dash-kpi-body">
      <div class="dash-kpi-label">Profitable Seats</div>
      <div class="dash-kpi-value" style="color:#b8f458">{profitable}</div>
      <div class="dash-kpi-sub">of {n_seats} total</div>
    </div>
  </div>
  <div class="dash-kpi-card" style="border-color:rgba(255,90,0,0.22)">
    <div class="dash-kpi-icon" style="background:rgba(255,90,0,0.15);color:#ff5a00;">!</div>
    <div class="dash-kpi-body">
      <div class="dash-kpi-label">Loss Seats</div>
      <div class="dash-kpi-value" style="color:#ff5a00">{loss_seats}</div>
      <div class="dash-kpi-sub" style="color:#cc4400">needs attention</div>
    </div>
  </div>
</div>
<div style="height:22px"></div>
""", unsafe_allow_html=True)

    # ── Revenue vs Cost  |  Cost Breakdown ────────────────────────────────────
    col_left, col_right = st.columns([3, 2])

    with col_left:
        st.markdown("""<div class="dash-section-hdr">
  <span class="dash-section-title">Revenue vs Cost by Client</span>
  <span class="dash-section-sub">sorted by revenue</span>
</div>""", unsafe_allow_html=True)
        client_rev  = df.groupby("Client Name")["Billing ($)"].sum() * rate
        client_cost = df.groupby("Client Name")[COST_COLS].sum().sum(axis=1) / _INR_RATE * rate
        rev_col, cost_col = f"Revenue ({pfx})", f"Cost ({pfx})"
        client_df = (
            pd.DataFrame({rev_col: client_rev, cost_col: client_cost})
            .reset_index()
            .sort_values(rev_col, ascending=True)
        )
        fig = px.bar(
            client_df, y="Client Name", x=[rev_col, cost_col],
            barmode="group",
            orientation="h",
            color_discrete_sequence=["#00d4aa", "#7b5ea7"],
            labels={"value": f"Amount ({pfx})", "variable": ""},
        )
        fig.update_layout(
            legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1),
            yaxis=dict(tickfont=dict(size=11)),
        )
        _chart_layout(fig, height=400, xaxis_tickprefix=pfx)
        st.plotly_chart(fig, use_container_width=True)

    with col_right:
        st.markdown("""<div class="dash-section-hdr">
  <span class="dash-section-title">Cost Breakdown</span>
  <span class="dash-section-sub">by category</span>
</div>""", unsafe_allow_html=True)
        cost_totals_disp = (
            df[COST_COLS].sum() / _INR_RATE * rate
        )
        cost_totals_disp = cost_totals_disp[cost_totals_disp > 0].sort_values(ascending=True)
        fig2 = px.bar(
            x=cost_totals_disp.values,
            y=cost_totals_disp.index,
            orientation="h",
            color=cost_totals_disp.values,
            color_continuous_scale=["#7b5ea7", "#00d4aa", "#b8f458"],
            labels={"x": f"Amount ({pfx})", "y": ""},
        )
        fig2.update_traces(
            text=[f"{pfx}{v:,.0f}" for v in cost_totals_disp.values],
            textposition="outside",
            textfont=dict(color="#7ecab0", size=10),
        )
        fig2.update_layout(coloraxis_showscale=False, showlegend=False)
        _chart_layout(fig2, height=400, xaxis_tickprefix=pfx)
        st.plotly_chart(fig2, use_container_width=True)

    # ── Seat balance chart ─────────────────────────────────────────────────────
    st.markdown("""<div class="dash-section-hdr">
  <span class="dash-section-title">Seat-level Balance</span>
  <span class="dash-section-sub">sorted by performance — green = profitable</span>
</div>""", unsafe_allow_html=True)
    bal_col_name = f"Balance ({pfx})"
    df[bal_col_name] = df["Balance"] * rate
    fig3 = px.bar(
        df.sort_values(bal_col_name),
        x="Seat Name", y=bal_col_name,
        color=bal_col_name,
        color_continuous_scale=["#ff5a00", "#ffd166", "#00d4aa"],
        color_continuous_midpoint=0,
        labels={bal_col_name: bal_col_name},
    )
    _chart_layout(fig3, xaxis_tickangle=-45, height=300, yaxis_tickprefix=pfx)
    st.plotly_chart(fig3, use_container_width=True)


def page_compare(data):
    st.title("📊 Month-over-Month Comparison")

    months = list(data.keys())
    if len(months) < 2:
        st.info("Need at least 2 months of data for comparison.")
        return

    show_inr  = st.session_state.get("show_inr", False)
    fx_rates  = load_fx_rates()
    cur_sym   = "₹" if show_inr else "$"

    summary = []
    for m, df in data.items():
        rate = fx_rates.get(m, _INR_RATE) if show_inr else 1.0
        summary.append({
            "Month":    m,
            "Revenue":  df["Billing ($)"].sum() * rate,
            "Cost":     df[COST_COLS].sum().sum() / _INR_RATE * rate,
            "Balance":  df["Balance"].sum() * rate,
            "Seats":    len(df),
        })
    sdf = pd.DataFrame(summary)

    c1, c2 = st.columns(2)
    with c1:
        fig = px.line(sdf, x="Month", y=["Revenue", "Cost", "Balance"],
                      markers=True, color_discrete_sequence=["#00d4aa", "#7b5ea7", "#b8f458"])
        _chart_layout(fig, title=f"Revenue / Cost / Balance Trend ({cur_sym})",
                      height=350, yaxis_tickprefix=cur_sym)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig2 = px.bar(sdf, x="Month", y="Balance",
                      color="Balance",
                      color_continuous_scale=["#ff5a00", "#00d4aa"],
                      color_continuous_midpoint=0)
        _chart_layout(fig2, title=f"Net Balance by Month ({cur_sym})",
                      height=350, yaxis_tickprefix=cur_sym)
        st.plotly_chart(fig2, use_container_width=True)

    st.subheader("Seat-level comparison")
    m1 = st.selectbox("Month A", months, index=0)
    m2 = st.selectbox("Month B", months, index=min(1, len(months)-1))

    rate1 = fx_rates.get(m1, _INR_RATE) if show_inr else 1.0
    rate2 = fx_rates.get(m2, _INR_RATE) if show_inr else 1.0

    df1 = data[m1][["Seat Name", "Billing ($)", "Balance"]].copy()
    df1[f"Revenue_{m1}"] = df1["Billing ($)"] * rate1
    df1[f"Balance_{m1}"] = df1["Balance"] * rate1
    df1 = df1[["Seat Name", f"Revenue_{m1}", f"Balance_{m1}"]]

    df2 = data[m2][["Seat Name", "Billing ($)", "Balance"]].copy()
    df2[f"Revenue_{m2}"] = df2["Billing ($)"] * rate2
    df2[f"Balance_{m2}"] = df2["Balance"] * rate2
    df2 = df2[["Seat Name", f"Revenue_{m2}", f"Balance_{m2}"]]

    merged = df1.merge(df2, on="Seat Name", how="outer").fillna(0)
    merged["Balance_Δ"] = merged[f"Balance_{m2}"] - merged[f"Balance_{m1}"]

    fig3 = px.bar(merged, x="Seat Name", y="Balance_Δ",
                  color="Balance_Δ",
                  color_continuous_scale=["#ff5a00", "#ffd166", "#00d4aa"],
                  color_continuous_midpoint=0,
                  title=f"Balance change: {m1} → {m2} ({cur_sym})")
    _chart_layout(fig3, xaxis_tickangle=-45, height=380, yaxis_tickprefix=cur_sym)
    st.plotly_chart(fig3, use_container_width=True)


def page_data_entry():
    st.title("➕ Add / Edit Data")
    st.write("Enter a new billing record. It will be saved and merged with the Excel data.")

    custom = load_json_data()

    with st.form("entry_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        month      = c1.text_input("Month (e.g. June)", placeholder="June")
        client     = c2.text_input("Client Name")
        seat       = c3.text_input("Seat Name")

        c4, c5 = st.columns(2)
        billing_usd = c4.number_input("Billing ($)", min_value=0.0, step=50.0)
        billing_inr = c5.number_input("Billing (Rs.)", min_value=0.0, step=1000.0)

        st.subheader("Costs (Rs.)")
        cost_vals = {}
        cols = st.columns(4)
        for i, col_name in enumerate(COST_COLS):
            cost_vals[col_name] = cols[i % 4].number_input(col_name, min_value=0.0, step=100.0, key=col_name)

        submitted = st.form_submit_button("Save Record", type="primary")

    if submitted:
        if not month or not client or not seat:
            st.error("Month, Client Name, and Seat Name are required.")
        else:
            total_cost = sum(cost_vals.values())
            balance    = billing_inr - total_cost
            record     = {
                "Sr. No.": 1,
                "Client Name": client,
                "Seat Name": seat,
                "Billing ($)": billing_usd,
                "Billing (Rs.)": billing_inr,
                **cost_vals,
                "Balance": balance,
            }
            if month not in custom:
                custom[month] = []
            # auto-number
            record["Sr. No."] = len(custom[month]) + 1
            custom[month].append(record)
            save_json_data(custom)
            st.success(f"Saved! Balance for this seat: {fmt_inr(balance)}")
            st.cache_data.clear()

    if custom:
        st.divider()
        st.subheader("Custom records (saved locally)")
        for m, rows in custom.items():
            with st.expander(f"{m} ({len(rows)} records)"):
                df = pd.DataFrame(rows)
                df.index = range(1, len(df) + 1)
                st.dataframe(df, use_container_width=True)
                if st.button(f"Delete all {m} records", key=f"del_{m}"):
                    del custom[m]
                    save_json_data(custom)
                    st.cache_data.clear()
                    st.rerun()


_SHARED_COST_COLS = [c for c in COST_COLS if c != "Salary/Payroll"]
_INR_RATE = 85  # Rs per $


@st.cache_data(ttl=0)
def _load_employee_sheet(path, month=None, _bust=0):
    """Return the Employee Summary DataFrame for `month` (e.g. 'Jan 2026').
    If month is given, tries the matching sheet first, then falls back to
    the most recent sheet available."""
    try:
        xl_names = pd.ExcelFile(path).sheet_names

        def _read_sheet(s):
            df = pd.read_excel(path, sheet_name=s)
            # Category = first col, Name = second col (header names vary across months)
            cat_col  = df.columns[0]
            name_col = df.columns[1]
            # Amount: prefer column whose header contains "$", else fall back to last col
            # This handles: 3-col (March: Amount $) and 4-col (Jan/Feb: Amount INR | Amount $)
            amount_col = next((c for c in df.columns if "$" in str(c)), df.columns[-1])
            return pd.DataFrame({
                "Category": df[cat_col].astype(str).str.strip(),
                "Name":     df[name_col].astype(str).str.strip(),
                "Amount":   pd.to_numeric(df[amount_col], errors="coerce").fillna(0),
            })

        # Try month-specific sheet first
        if month:
            abbr = month.split()[0][:3]
            s = _find_sheet(xl_names, ["employee summary", "employee summery"], abbr)
            if s:
                return _read_sheet(s)

        # Fall back to most recent available
        for abbr in _ALL_MONTH_ABBRS:
            s = _find_sheet(xl_names, ["employee summary", "employee summery"], abbr)
            if s:
                return _read_sheet(s)

        return pd.DataFrame(columns=["Category", "Name", "Amount"])
    except Exception:
        return pd.DataFrame(columns=["Category", "Name", "Amount"])


def load_billing_salaries(path, month=None, _bust=0):
    """Load billing employee $ salaries from the Employee Summary sheet for `month`.

    Excludes any employees listed in _remove_from_billing for this month.
    """
    df = _load_employee_sheet(path, month=month, _bust=_bust)
    bdf = df[df["Category"] == "Payroll - Billing"].copy()
    if month:
        removed = load_overrides().get(month, {}).get("_remove_from_billing", [])
        if removed:
            removed_lower = [n.strip().lower() for n in removed]
            bdf = bdf[~bdf["Name"].str.strip().str.lower().isin(removed_lower)]
    return {str(r["Name"]).strip().lower(): float(r["Amount"]) for _, r in bdf.iterrows()}


def load_bench_employees(path, month=None, _bust=0):
    """Return a DataFrame of individual bench (Payroll - Non Bill) employees for `month`.

    Merges Excel data with any 'Bench ($)' overrides from pnl_overrides.json:
    overrides update existing rows or add new ones.
    """
    df = _load_employee_sheet(path, month=month, _bust=_bust)
    bdf = df[df["Category"] == "Payroll - Non Bill"].copy()
    bdf = bdf[["Name", "Amount"]].rename(columns={"Amount": "Payroll ($)"})
    bdf = bdf[~bdf["Name"].str.lower().isin(["nan", "none", "", "total"])].reset_index(drop=True)

    if month:
        bench_ovr = {
            name: vals["Bench ($)"]
            for name, vals in load_overrides().get(month, {}).items()
            if "Bench ($)" in vals
        }
        if bench_ovr:
            for name, amt in bench_ovr.items():
                mask = bdf["Name"].str.strip().str.lower() == name.strip().lower()
                if mask.any():
                    bdf.loc[mask, "Payroll ($)"] = amt
                else:
                    bdf = pd.concat(
                        [bdf, pd.DataFrame({"Name": [name], "Payroll ($)": [float(amt)]})],
                        ignore_index=True,
                    )

    bdf = bdf.sort_values("Name").reset_index(drop=True)
    bdf.index = range(1, len(bdf) + 1)
    return bdf


def load_bench_salaries(path, month=None, _bust=0):
    """Return total bench payroll $ for `month`, using overrides when available."""
    return load_bench_employees(path, month=month, _bust=_bust)["Payroll ($)"].sum()


def _match_salary(seat_name, salary_lookup):
    """
    Match a billing seat name to a salary in the Employee Summary lookup.
    Tries: exact → strip/case → prefix → word-overlap → fuzzy similarity.
    Returns (salary_usd, matched_key) or (None, None).
    """
    from difflib import SequenceMatcher

    seat = seat_name.strip().lower()

    # 1. Exact match (handles trailing-space names like "Mudassar ")
    if seat in salary_lookup:
        return salary_lookup[seat], seat

    # 2. Key starts with seat name  ("rochelle" → "rochelle extross")
    for key in salary_lookup:
        if key.startswith(seat + " ") or key == seat:
            return salary_lookup[key], key

    # 3. Seat name starts with key  (longer alias for a short key)
    for key in salary_lookup:
        if seat.startswith(key + " ") or seat == key:
            return salary_lookup[key], key

    # 4. First/last-name word overlap (e.g. "Veerendra/Veer" aliases)
    seat_words = [w for w in seat.replace("/", " ").split() if len(w) > 3]
    for key in salary_lookup:
        key_words = key.split()
        if any(sw in key_words for sw in seat_words):
            return salary_lookup[key], key

    # 5. Fuzzy similarity ≥ 0.82 — catches typos like "Sandip Raj" ↔ "Sandeep Raj"
    best_ratio, best_key = 0.0, None
    for key in salary_lookup:
        ratio = SequenceMatcher(None, seat, key).ratio()
        if ratio > best_ratio:
            best_ratio, best_key = ratio, key
    if best_ratio >= 0.82:
        return salary_lookup[best_key], best_key

    return None, None


def _compute_pnl(df_month, salary_lookup=None):
    """
    Per-employee P&L engine (rock-solid formula):
      • Salary: matched from Employee Summery ($ values provided by user);
                falls back to Salary/Payroll col ÷ INR rate if unmatched.
      • All other costs → sum(col across ALL billing employees) ÷ N ÷ INR rate
        This works whether the Excel stores raw totals or pre-divided equal values.
      • Net = Billing ($) − Total Cost ($)
    Returns a clean DataFrame with one row per employee, all values in $.
    """
    N = len(df_month)
    if N == 0:
        return pd.DataFrame()

    if salary_lookup is None:
        salary_lookup = {}

    # Pre-compute shared cost totals once (sum → divide by N → convert to $)
    shared_per_emp = {
        c: pd.to_numeric(df_month[c], errors="coerce").fillna(0).sum() / N / _INR_RATE
        for c in _SHARED_COST_COLS
    }

    rows = []
    unmatched = []
    for _, row in df_month.iterrows():
        seat_name = str(row.get("Seat Name", "") or "")
        billing   = float(pd.to_numeric(row.get("Billing ($)"), errors="coerce") or 0)

        # Salary: prefer Employee Summery $ value; fall back to billing sheet Rs ÷ INR rate
        matched_sal, matched_key = _match_salary(seat_name, salary_lookup)
        if matched_sal is not None:
            salary       = matched_sal
            salary_source = "Employee Summery"
        else:
            salary_rs    = float(pd.to_numeric(row.get("Salary/Payroll"), errors="coerce") or 0)
            salary       = salary_rs / _INR_RATE
            salary_source = "Excel (Rs÷85)"
            unmatched.append(seat_name)

        total_cost = salary + sum(shared_per_emp.values())
        net        = billing - total_cost
        margin     = (net / billing * 100) if billing else 0

        rows.append({
            "Client":          row.get("Client Name", ""),
            "Employee":        seat_name,
            "Billing ($)":     billing,
            "Salary ($)":      salary,
            "Salary Source":   salary_source,
            **{c: shared_per_emp[c] for c in _SHARED_COST_COLS},
            "Total Cost ($)":  total_cost,
            "Net ($)":         net,
            "Margin %":        margin,
        })

    result = pd.DataFrame(rows)
    result.index = range(1, len(result) + 1)
    result._unmatched = unmatched  # carry unmatched list for UI display
    return result


def _apply_overrides_and_recalc(result, month_overrides):
    """Apply saved overrides to Billing/Salary columns and recompute derived cols."""
    for emp, vals in month_overrides.items():
        mask = result["Employee"] == emp
        if not mask.any():
            continue
        for col, val in vals.items():
            if col in result.columns:
                result.loc[mask, col] = val
    # Recompute derived columns from (possibly overridden) Billing + Salary
    result["Total Cost ($)"] = result["Salary ($)"] + result[_SHARED_COST_COLS].sum(axis=1)
    result["Net ($)"]        = result["Billing ($)"] - result["Total Cost ($)"]
    result["Margin %"]       = result.apply(
        lambda r: (r["Net ($)"] / r["Billing ($)"] * 100) if r["Billing ($)"] else 0, axis=1
    )
    return result


def page_detail(data):
    st.title("🔍 Per-Employee P&L")

    months = list(data.keys())
    if not months:
        st.warning("No data available.")
        return

    sel_month = st.selectbox("Month", months, index=len(months) - 1)
    full_df   = data[sel_month].copy()
    N_total   = len(full_df)

    if N_total == 0:
        st.warning("No billing employees found for this month.")
        return

    # Load salary lookup and overrides
    bust          = st.session_state.get("bust", 0)
    salary_lookup = load_billing_salaries(_get_active_excel(), month=sel_month, _bust=bust)
    all_overrides = load_overrides()
    month_ovr     = all_overrides.get(sel_month, {})

    # Compute base P&L, then apply any saved overrides
    result = _compute_pnl(full_df, salary_lookup=salary_lookup)
    result = _apply_overrides_and_recalc(result, month_ovr)

    # Warn about unmatched salary names (fallback employees)
    unmatched = getattr(result, "_unmatched", [])
    unmatched_without_override = [n for n in unmatched if not (n in month_ovr and "Salary ($)" in month_ovr[n])]
    if unmatched_without_override:
        with st.expander(f"⚠️ {len(unmatched_without_override)} employee(s) missing from Employee Summary sheet"):
            st.caption("These employees have no salary entry in the Excel 'Employee Summary' sheet. "
                       "Add them to the sheet and refresh, or enter their salary directly in the table below.")
            for name in unmatched:
                ovr_note = " ✅ salary overridden" if name in month_ovr and "Salary ($)" in month_ovr[name] else ""
                st.write(f"  • **{name}**{ovr_note}")

    # ── Filters ────────────────────────────────────────────────────────────────
    f1, f2 = st.columns(2)
    clients    = ["All Clients"] + sorted(full_df["Client Name"].dropna().unique().tolist())
    sel_client = f1.selectbox("Filter by Client", clients)
    search     = f2.text_input("Search employee name")

    display_cols = ["Client", "Employee", "Billing ($)", "Salary ($)"] + _SHARED_COST_COLS + ["Total Cost ($)", "Net ($)", "Margin %"]
    view = result[display_cols].copy()
    if sel_client != "All Clients":
        view = view[view["Client"] == sel_client]
    if search:
        view = view[view["Employee"].str.contains(search, case=False, na=False)]
    view = view.reset_index(drop=True)
    view.index = range(1, len(view) + 1)

    # ── KPIs (always from full unfiltered result) ──────────────────────────────
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Billing Employees", N_total)
    k2.metric("Total Billing",     _disp(result["Billing ($)"].sum(),    sel_month))
    k3.metric("Total Cost",        _disp(result["Total Cost ($)"].sum(), sel_month))
    net_total = result["Net ($)"].sum()
    k4.metric("Net P&L",           _disp(net_total, sel_month),
              delta=f"{'▲' if net_total >= 0 else '▼'}")
    k5.metric("Avg Margin",        f"{result['Margin %'].mean():.1f}%")

    if month_ovr:
        st.caption(f"✏️ {len(month_ovr)} override(s) active for {sel_month}")

    st.divider()

    # ── Table — ₹ read-only view OR $ editable view ────────────────────────────
    show_inr = st.session_state.get("show_inr", False)

    if show_inr:
        fx_rate  = get_fx_rate(sel_month)
        inr_view = view.copy()
        money_cols = ["Billing ($)", "Salary ($)"] + _SHARED_COST_COLS + ["Total Cost ($)", "Net ($)"]
        rename_map = {
            "Billing ($)":    "Billing (₹)",
            "Salary ($)":     "Salary (₹)",
            "Total Cost ($)": "Total Cost (₹)",
            "Net ($)":        "Net (₹)",
        }
        for c in money_cols:
            inr_view[c] = (inr_view[c] * fx_rate).round(0)
        inr_view = inr_view.rename(columns=rename_map)

        inr_col_cfg = {
            "Client":          st.column_config.TextColumn("Client",          disabled=True),
            "Employee":        st.column_config.TextColumn("Employee",        disabled=True),
            "Billing (₹)":     st.column_config.NumberColumn("Billing (₹)",   format="₹%.0f", disabled=True),
            "Salary (₹)":      st.column_config.NumberColumn("Salary (₹)",    format="₹%.0f", disabled=True),
            "Total Cost (₹)":  st.column_config.NumberColumn("Total Cost (₹)",format="₹%.0f", disabled=True),
            "Net (₹)":         st.column_config.NumberColumn("Net (₹)",       format="₹%.0f", disabled=True),
            "Margin %":        st.column_config.NumberColumn("Margin %",      format="%.1f%%", disabled=True),
            **{c: st.column_config.NumberColumn(c, format="₹%.0f", disabled=True) for c in _SHARED_COST_COLS},
        }
        st.caption(f"Read-only ₹ view  ($1 = ₹{fx_rate:.0f}). Switch to $ mode to edit values.")
        st.dataframe(inr_view, column_config=inr_col_cfg, use_container_width=True, height=520)
        # Still run a no-op data_editor in $ (hidden) so save/clear buttons have an `edited` var
        edited = view.copy()
    else:
        st.markdown("**Edit Billing ($) or Salary ($) directly in the table, then click Save.**")
        col_cfg = {
            "Client":        st.column_config.TextColumn("Client",        disabled=True),
            "Employee":      st.column_config.TextColumn("Employee",      disabled=True),
            "Billing ($)":   st.column_config.NumberColumn("Billing ($)", format="$%.2f", step=0.01, min_value=0.0),
            "Salary ($)":    st.column_config.NumberColumn("Salary ($)",  format="$%.2f", step=0.01, min_value=0.0),
            "Total Cost ($)":st.column_config.NumberColumn("Total Cost ($)", format="$%.2f", disabled=True),
            "Net ($)":       st.column_config.NumberColumn("Net ($)",     format="$%.2f", disabled=True),
            "Margin %":      st.column_config.NumberColumn("Margin %",    format="%.1f%%", disabled=True),
            **{c: st.column_config.NumberColumn(c, format="$%.2f", disabled=True) for c in _SHARED_COST_COLS},
        }
        edited = st.data_editor(
            view,
            column_config=col_cfg,
            use_container_width=True,
            height=520,
            key=f"pnl_editor_{sel_month}_{sel_client}_{search}",
        )

    # Recalculate Net/Total/Margin live from edited values so user sees updated numbers
    edited["Total Cost ($)"] = edited["Salary ($)"] + edited[_SHARED_COST_COLS].sum(axis=1)
    edited["Net ($)"]        = edited["Billing ($)"] - edited["Total Cost ($)"]
    edited["Margin %"]       = edited.apply(
        lambda r: (r["Net ($)"] / r["Billing ($)"] * 100) if r["Billing ($)"] else 0, axis=1
    )

    # ── Save / Clear buttons ───────────────────────────────────────────────────
    btn_col1, btn_col2, btn_col3 = st.columns([2, 2, 6])

    if btn_col1.button("💾 Save Changes", type="primary"):
        # Detect which cells changed vs. original computed view
        new_overrides = dict(month_ovr)
        changed = 0
        for _, row in edited.iterrows():
            emp  = row["Employee"]
            orig = view[view["Employee"] == emp]
            if orig.empty:
                continue
            orig_row   = orig.iloc[0]
            emp_changes = {}
            if abs(row["Billing ($)"] - orig_row["Billing ($)"]) > 0.005:
                emp_changes["Billing ($)"] = round(float(row["Billing ($)"]), 2)
            if abs(row["Salary ($)"]  - orig_row["Salary ($)"])  > 0.005:
                emp_changes["Salary ($)"]  = round(float(row["Salary ($)"]),  2)
            if emp_changes:
                new_overrides[emp] = {**new_overrides.get(emp, {}), **emp_changes}
                changed += 1
        if changed:
            all_overrides[sel_month] = new_overrides
            save_overrides(all_overrides)
            st.success(f"Saved {changed} override(s) for {sel_month}.")
            st.rerun()
        else:
            st.info("No changes detected.")

    if btn_col2.button("🗑 Clear Overrides"):
        if sel_month in all_overrides:
            del all_overrides[sel_month]
            save_overrides(all_overrides)
            st.success(f"Cleared all overrides for {sel_month}.")
            st.rerun()
        else:
            st.info("No overrides to clear.")

    # ── Download ───────────────────────────────────────────────────────────────
    csv = edited.to_csv(index=False).encode()
    st.download_button("⬇ Download CSV", csv, "per_employee_pnl.csv", "text/csv", key="dl_pnl")

    # ── Chart ──────────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("Net P&L per Employee")
    chart_df        = edited.sort_values("Net ($)").copy()
    chart_df["Color"] = chart_df["Net ($)"].apply(lambda v: "Profit" if v >= 0 else "Loss")
    fig = px.bar(
        chart_df, x="Employee", y="Net ($)",
        color="Color",
        color_discrete_map={"Profit": "#00d4aa", "Loss": "#ff5a00"},
        hover_data={"Client": True, "Billing ($)": True, "Total Cost ($)": True, "Margin %": ":.1f"},
        labels={"Net ($)": "Net P&L ($)", "Employee": ""},
        height=420,
    )
    _chart_layout(fig, xaxis_tickangle=-45, showlegend=True, legend_title_text="", yaxis_tickprefix="$")
    fig.add_hline(y=0, line_dash="dash", line_color="rgba(0,212,170,0.4)", line_width=1)
    st.plotly_chart(fig, use_container_width=True)


def page_billing_clients(data):
    st.title("🏢 Billing Clients")

    months = list(data.keys())
    if not months:
        st.warning("No data available.")
        return

    sel_month = st.selectbox("Select Month", months, index=len(months) - 1, key="bc_month")
    fx_rate   = get_fx_rate(sel_month)

    sales_map = load_sales_persons()
    unique_df = data[sel_month][["Client Name", "Seat Name", "Billing ($)"]].copy()
    unique_df["Billing (₹)"] = (unique_df["Billing ($)"] * fx_rate).round(0)
    unique_df["Sales Person"] = unique_df["Client Name"].map(lambda c: sales_map.get(c, ""))

    # ── Filters ──────────────────────────────────────────────────────────────
    col_f1, col_f2 = st.columns(2)
    client_opts = ["All Clients"] + sorted(unique_df["Client Name"].dropna().unique().tolist())
    sel_client  = col_f1.selectbox("Filter by Client", client_opts)
    search      = col_f2.text_input("Search recruiter name")

    view = unique_df.copy()
    if sel_client != "All Clients":
        view = view[view["Client Name"] == sel_client]
    if search:
        view = view[view["Seat Name"].str.contains(search, case=False, na=False)]

    # ── KPIs ─────────────────────────────────────────────────────────────────
    lbl = _currency_label(sel_month)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Clients",    view["Client Name"].nunique())
    k2.metric("Total Recruiters", len(view))
    k3.metric(f"Total Billing ($)",  f"${view['Billing ($)'].sum():,.0f}")
    k4.metric(f"Total Billing (₹)",  f"₹{view['Billing (₹)'].sum():,.0f}")

    st.divider()

    tab1, tab2 = st.tabs(["Recruiter Detail", "Client Summary"])

    with tab1:
        st.subheader("All Recruiters & Bill Rates")
        detail = (
            view[["Client Name", "Sales Person", "Seat Name", "Billing ($)", "Billing (₹)"]]
            .rename(columns={
                "Seat Name":   "Recruiter",
                "Billing ($)": "Bill Rate ($)",
                "Billing (₹)": "Bill Rate (₹)",
            })
            .sort_values(["Client Name", "Recruiter"])
            .reset_index(drop=True)
        )

        edited = st.data_editor(
            detail,
            use_container_width=True,
            height=500,
            hide_index=True,
            column_config={
                "Client Name":  st.column_config.TextColumn("Client Name",  disabled=True),
                "Sales Person": st.column_config.TextColumn("Sales Person", help="Click to assign or edit"),
                "Recruiter":    st.column_config.TextColumn("Recruiter",    disabled=True),
                "Bill Rate ($)":st.column_config.NumberColumn("Bill Rate ($)", format="$%.0f", disabled=True),
                "Bill Rate (₹)":st.column_config.NumberColumn("Bill Rate (₹)", format="₹%.0f", disabled=True),
            },
            key="bc_detail_editor",
        )

        # Persist any Sales Person changes
        changed = edited[["Client Name", "Sales Person"]].drop_duplicates("Client Name")
        new_map = {r["Client Name"]: r["Sales Person"] for _, r in changed.iterrows() if r["Client Name"]}
        if new_map != {k: sales_map.get(k, "") for k in new_map}:
            sales_map.update({k: v for k, v in new_map.items() if v != sales_map.get(k, "")})
            save_sales_persons(sales_map)

        total_billing  = detail["Bill Rate ($)"].sum()
        total_billing_inr = detail["Bill Rate (₹)"].sum()
        st.caption(
            f"₹ column uses wire rate: $1 = ₹{fx_rate:.0f} ({sel_month})  ·  "
            f"**Total: ${total_billing:,.0f}  /  ₹{total_billing_inr:,.0f}**"
        )

        csv = detail.to_csv(index=False).encode()
        st.download_button("Download CSV", csv, "billing_clients.csv", "text/csv")

    with tab2:
        st.subheader("Client-wise Summary")
        summary = (
            view.groupby("Client Name", as_index=False)
            .agg(
                Recruiters=("Seat Name", "count"),
                Total_Billing_USD=("Billing ($)", "sum"),
                Total_Billing_INR=("Billing (₹)", "sum"),
                Avg_Bill_Rate_USD=("Billing ($)", "mean"),
            )
            .sort_values("Total_Billing_INR", ascending=False)
        )
        summary.columns = ["Client Name", "# Recruiters", "Total Billing ($)", "Total Billing (₹)", "Avg Bill Rate ($)"]

        total_row_s = pd.DataFrame([{
            "Client Name":        "TOTAL",
            "# Recruiters":       summary["# Recruiters"].sum(),
            "Total Billing ($)":  summary["Total Billing ($)"].sum(),
            "Total Billing (₹)":  summary["Total Billing (₹)"].sum(),
            "Avg Bill Rate ($)":  summary["Total Billing ($)"].sum() / summary["# Recruiters"].sum()
                                  if summary["# Recruiters"].sum() else 0,
        }])
        summary_display = pd.concat([summary, total_row_s], ignore_index=True)
        summary_display.index = range(1, len(summary_display) + 1)
        total_idx_s = len(summary_display)

        def _bold_total_row_s(row):
            style = "font-weight: bold; background-color: #1a1a2e; color: white"
            return [style if row.name == total_idx_s else "" for _ in row]

        def _billing_bar(col):
            data_rows = col.iloc[:-1]
            lo, hi = data_rows.min(), data_rows.max()
            styles = []
            for i, v in enumerate(col):
                if i == len(col) - 1:
                    styles.append("")
                elif hi > lo:
                    # ratio 0 = lowest billing, 1 = highest
                    ratio = (v - lo) / (hi - lo)
                    # dark navy (#0a2850) → medium blue (#1a6bbf)
                    r = int(10 + 16 * ratio)
                    g = int(40 + 67 * ratio)
                    b = int(80 + 111 * ratio)
                    styles.append(f"background-color: rgb({r},{g},{b}); color: white")
                else:
                    styles.append("background-color: rgb(10,40,80); color: white")
            return styles

        styled_s = (
            summary_display.style
            .apply(_bold_total_row_s, axis=1)
            .apply(_billing_bar, subset=["Total Billing ($)"])
            .format({
                "Total Billing ($)":  "${:,.0f}",
                "Total Billing (₹)":  "₹{:,.0f}",
                "Avg Bill Rate ($)":  "${:,.0f}",
            })
        )
        st.dataframe(styled_s, use_container_width=True, height=400)
        st.caption(f"₹ column uses wire rate: $1 = ₹{fx_rate:.0f} ({sel_month})")

        st.subheader("Bill Rate by Client")
        y_col = "Billing ($)"
        fig = px.bar(
            view.sort_values(y_col, ascending=False),
            x="Client Name", y=y_col,
            color="Seat Name",
            labels={y_col: f"Bill Rate ($)", "Seat Name": "Recruiter"},
            height=420,
        )
        _chart_layout(fig, xaxis_tickangle=-35, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)


def page_employee_details(data):
    st.title("👥 Employee Details")

    months = list(data.keys())
    emp_data = load_employees()

    emp_month     = st.selectbox("Select Month", months, index=len(months) - 1, key="ed_month")
    bust          = st.session_state.get("cache_bust", 0)
    salary_lookup = load_billing_salaries(_get_active_excel(), month=emp_month, _bust=bust)
    bench_df      = load_bench_employees(_get_active_excel(), month=emp_month, _bust=bust)

    # Build billed employee list from the selected month only
    _month_df = data[emp_month][["Client Name", "Seat Name", "Billing ($)"]].copy()
    billed_df = (
        _month_df
        .rename(columns={"Seat Name": "Name"})
        .sort_values(["Client Name", "Name"])
        .reset_index(drop=True)
    )
    billed_df["Payroll ($)"] = billed_df["Name"].apply(
        lambda n: _match_salary(n, salary_lookup)[0] or 0.0
    )
    # Apply salary overrides from pnl_overrides.json for this month
    _emp_ovr = load_overrides().get(emp_month, {})
    for _emp, _vals in _emp_ovr.items():
        if _emp.startswith("_"):
            continue
        if "Salary ($)" in _vals:
            _mask = billed_df["Name"] == _emp
            if _mask.any():
                billed_df.loc[_mask, "Payroll ($)"] = _vals["Salary ($)"]

    billed_df["Gross Margin ($)"] = billed_df["Billing ($)"] - billed_df["Payroll ($)"]

    total_non_billing   = sum(len(v) for v in emp_data.values())
    total_nb_salary     = sum(
        r.get("Salary ($)", 0) or 0
        for dept_list in emp_data.values()
        for r in dept_list
    )
    total_bill_payroll  = billed_df["Payroll ($)"].sum()
    total_billing       = billed_df["Billing ($)"].sum()
    total_gross_margin  = billed_df["Gross Margin ($)"].sum()
    total_bench_payroll = bench_df["Payroll ($)"].sum() if not bench_df.empty else 0.0
    bench_pct           = (len(bench_df) / len(billed_df) * 100) if len(billed_df) > 0 else 0.0

    _ED_TABS = ["🟢 Billed Employees", "🟡 Bench Employees", "🔵 Non-Billing Employees"]

    # ── Top KPIs (all st.metric for visual consistency) ───────────────────────
    k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
    k1.metric("Billed Employees",  len(billed_df))
    k2.metric("Bench Employees",   len(bench_df))
    k3.metric("Bench %",           f"{bench_pct:.1f}%")
    k4.metric("Non-Billing Staff", total_non_billing)
    k5.metric("Total Billing",     _disp(total_billing,      emp_month, dec=2))
    k6.metric("Billing Payroll",   _disp(total_bill_payroll, emp_month, dec=2))
    k7.metric("Gross Margin",      _disp(total_gross_margin, emp_month, dec=2))

    # JS: clicking first 4 metric cards triggers the radio selector below
    import streamlit.components.v1 as _components
    _components.html("""
<script>
(function() {
    var NAV = {
        'BILLED EMPLOYEES': 0,
        'BENCH EMPLOYEES':  1,
        'BENCH %':          1,
        'NON-BILLING STAFF':2
    };
    function setup() {
        var labels = parent.document.querySelectorAll('[data-testid="stMetricLabel"]');
        labels.forEach(function(lbl) {
            var key = lbl.innerText.trim().toUpperCase();
            if (!(key in NAV)) return;
            var card = lbl.closest('[data-testid="metric-container"]');
            if (!card || card._navAttached) return;
            card._navAttached = true;
            card.classList.add('kpi-nav-card');
            card.style.cursor = 'pointer';
            card.addEventListener('click', function() {
                var radios = parent.document.querySelectorAll(
                    '[data-testid="stRadio"] input[type="radio"]');
                if (radios[NAV[key]]) radios[NAV[key]].click();
            });
        });
    }
    setTimeout(setup, 200);
    setTimeout(setup, 800);
    new MutationObserver(setup).observe(parent.document.body,
        {childList: true, subtree: true});
})();
</script>
""", height=0)

    st.divider()

    active_section = st.radio(
        "Section",
        _ED_TABS,
        horizontal=True,
        label_visibility="collapsed",
        key="ed_section",
    )

    st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)

    # ── Billed Employees ──────────────────────────────────────────────────────
    if active_section == _ED_TABS[0]:
        col_f1, col_f2 = st.columns(2)
        client_opts = ["All Clients"] + sorted(billed_df["Client Name"].dropna().unique().tolist())
        sel_client  = col_f1.selectbox("Filter by Client", client_opts, key="emp_client")
        search      = col_f2.text_input("Search by name", key="emp_search")

        view = billed_df.copy()
        if sel_client != "All Clients":
            view = view[view["Client Name"] == sel_client]
        if search:
            view = view[view["Name"].str.contains(search, case=False, na=False)]

        view = view.reset_index(drop=True)
        view.index = range(1, len(view) + 1)

        base_fmt = {
            "Billing ($)":      "${:,.2f}",
            "Payroll ($)":      "${:,.2f}",
            "Gross Margin ($)": "${:,.2f}",
        }
        if st.session_state.get("show_inr", False):
            fx_rate = get_fx_rate(emp_month)
            view["Billing (₹)"]      = (view["Billing ($)"]      * fx_rate).round(0)
            view["Payroll (₹)"]      = (view["Payroll ($)"]       * fx_rate).round(0)
            view["Gross Margin (₹)"] = (view["Gross Margin ($)"]  * fx_rate).round(0)
            base_fmt.update({
                "Billing (₹)":      "₹{:,.0f}",
                "Payroll (₹)":      "₹{:,.0f}",
                "Gross Margin (₹)": "₹{:,.0f}",
            })

        def _color_margin(val):
            if isinstance(val, (int, float)):
                return "color: #4ade80; font-weight:600" if val >= 0 else "color: #f87171; font-weight:600"
            return ""

        # Append totals row
        num_cols = [c for c in view.columns if c not in ("Client Name", "Name")]
        totals_row = {c: view[c].sum() if c in num_cols else ("" if c == "Client Name" else "TOTAL") for c in view.columns}
        view_with_total = pd.concat([view, pd.DataFrame([totals_row])], ignore_index=True)
        view_with_total.index = list(range(1, len(view) + 1)) + [""]

        def _style_total_row(row):
            is_total = row.name == ""
            if is_total:
                return ["font-weight:bold; border-top: 2px solid #7ecab0"] * len(row)
            return [""] * len(row)

        gm_cols = [c for c in view_with_total.columns if "Gross Margin" in c]
        styled = (
            view_with_total.style
            .format(base_fmt)
            .map(_color_margin, subset=gm_cols)
            .apply(_style_total_row, axis=1)
        )
        st.dataframe(styled, use_container_width=True, height=510)

        csv = view.to_csv(index=False).encode()
        st.download_button("Download CSV", csv, "billed_employees.csv", "text/csv", key="dl_billed")

    # ── Bench Employees ───────────────────────────────────────────────────────
    elif active_section == _ED_TABS[1]:
        st.caption(f"Employees on bench (Payroll - Non Bill) for **{emp_month}** — sourced from Excel Employee Summary sheet.")

        if bench_df.empty:
            st.info("No bench employees found in the Employee Summary sheet for this month.")
        else:
            search_bench = st.text_input("Search by name", key="bench_search")
            view_bench = bench_df.copy()
            if search_bench:
                view_bench = view_bench[view_bench["Name"].str.contains(search_bench, case=False, na=False)]
                view_bench = view_bench.reset_index(drop=True)
                view_bench.index = range(1, len(view_bench) + 1)

            if st.session_state.get("show_inr", False):
                fx_rate = get_fx_rate(emp_month)
                view_bench["Payroll (₹)"] = (view_bench["Payroll ($)"] * fx_rate).round(0)
                fmt_b = {"Payroll ($)": "${:,.2f}", "Payroll (₹)": "₹{:,.0f}"}
            else:
                fmt_b = {"Payroll ($)": "${:,.2f}"}

            st.dataframe(
                view_bench.style.format(fmt_b),
                use_container_width=True,
                height=min(60 + len(view_bench) * 35, 480),
            )

            b1, b2 = st.columns(2)
            b1.metric("Bench Headcount", len(bench_df))
            b2.metric("Total Bench Payroll", f"${total_bench_payroll:,.2f}")

            csv_bench = bench_df.to_csv(index=False).encode()
            st.download_button("Download Bench CSV", csv_bench, f"bench_employees_{emp_month.replace(' ', '_')}.csv", "text/csv", key="dl_bench")

    # ── Non-Billing Employees ─────────────────────────────────────────────────
    elif active_section == _ED_TABS[2]:
        # Tabs: Marketing | HR & Admin | Management
        DISPLAY_TABS   = ["Marketing", "HR & Admin", "Management"]
        tab_marketing, tab_hradmin, tab_mgmt = st.tabs([f"📌 {d}" for d in DISPLAY_TABS])

        def _render_dept_tab(dept, dtab):
            """Render a single-dept tab (Marketing or Management)."""
            with dtab:
                employees = emp_data.get(dept, [])
                dept_salary = sum(r.get("Salary ($)", 0) or 0 for r in employees)
                d1, d2 = st.columns(2)
                d1.metric(f"{dept} Headcount", len(employees))
                d2.metric(f"{dept} Payroll",   f"${dept_salary:,.2f}")

                if employees:
                    df_dept = pd.DataFrame(employees)
                    for col in EMP_FIELDS:
                        if col not in df_dept.columns:
                            df_dept[col] = ""
                    df_dept = df_dept[EMP_FIELDS]
                    df_dept.index = range(1, len(df_dept) + 1)
                    styled_d = df_dept.style.format({"Salary ($)": lambda v: f"${v:,.2f}" if v else "—"})
                    st.dataframe(styled_d, use_container_width=True, height=280)
                    csv_d = df_dept.to_csv(index=False).encode()
                    st.download_button(f"Download {dept} CSV", csv_d, f"{dept.lower()}_employees.csv", "text/csv", key=f"dl_{dept}")
                else:
                    st.info(f"No {dept} employees added yet. Use the form below to add.")

                with st.expander(f"➕ Add {dept} Employee"):
                    with st.form(f"form_{dept}", clear_on_submit=True):
                        r1c1, r1c2, r1c3, r1c4 = st.columns(4)
                        name        = r1c1.text_input("Name *")
                        designation = r1c2.text_input("Designation")
                        salary      = r1c3.number_input("Salary ($)", min_value=0.0, step=10.0, format="%.2f")
                        join_date   = r1c4.text_input("Join Date (e.g. Jan 2025)")
                        notes       = st.text_input("Notes")
                        submitted   = st.form_submit_button("Save Employee", type="primary")
                    if submitted:
                        if not name:
                            st.error("Name is required.")
                        else:
                            emp_data[dept].append({"Name": name, "Designation": designation, "Salary ($)": salary, "Join Date": join_date, "Notes": notes})
                            save_employees(emp_data)
                            st.success(f"Added {name} to {dept}.")
                            st.rerun()

                if employees:
                    with st.expander(f"🗑 Remove a {dept} Employee"):
                        names_list = [r.get("Name", f"Row {i}") for i, r in enumerate(employees)]
                        to_delete  = st.selectbox("Select employee to remove", names_list, key=f"del_{dept}")
                        if st.button(f"Remove {to_delete}", key=f"delbtn_{dept}"):
                            emp_data[dept] = [r for r in employees if r.get("Name") != to_delete]
                            save_employees(emp_data)
                            st.success(f"Removed {to_delete}.")
                            st.rerun()

        _render_dept_tab("Marketing",   tab_marketing)
        _render_dept_tab("Management",  tab_mgmt)

        # ── HR & Admin combined tab ───────────────────────────────────────────
        with tab_hradmin:
            hr_emps    = emp_data.get("HR", [])
            admin_emps = emp_data.get("Admin", [])
            all_emps   = hr_emps + admin_emps
            total_salary = sum(r.get("Salary ($)", 0) or 0 for r in all_emps)

            d1, d2, d3 = st.columns(3)
            d1.metric("HR Headcount",       len(hr_emps))
            d2.metric("Admin Headcount",     len(admin_emps))
            d3.metric("HR & Admin Payroll",  f"${total_salary:,.2f}")

            if all_emps:
                def _tag_dept(rows, label):
                    out = []
                    for r in rows:
                        rec = {**r, "Dept": label}
                        out.append(rec)
                    return out

                combined_rows = _tag_dept(hr_emps, "HR") + _tag_dept(admin_emps, "Admin")
                df_ha = pd.DataFrame(combined_rows)
                show_ha = ["Dept"] + EMP_FIELDS
                for col in show_ha:
                    if col not in df_ha.columns:
                        df_ha[col] = ""
                df_ha = df_ha[show_ha]
                df_ha.index = range(1, len(df_ha) + 1)
                styled_ha = df_ha.style.format({"Salary ($)": lambda v: f"${v:,.2f}" if v else "—"})
                st.dataframe(styled_ha, use_container_width=True, height=300)
                csv_ha = df_ha.to_csv(index=False).encode()
                st.download_button("Download HR & Admin CSV", csv_ha, "hr_admin_employees.csv", "text/csv", key="dl_hradmin")
            else:
                st.info("No HR or Admin employees added yet.")

            # Add form — choose sub-dept
            with st.expander("➕ Add HR / Admin Employee"):
                with st.form("form_hradmin", clear_on_submit=True):
                    sub_dept    = st.selectbox("Department", ["HR", "Admin"])
                    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
                    name        = r1c1.text_input("Name *")
                    designation = r1c2.text_input("Designation")
                    salary      = r1c3.number_input("Salary ($)", min_value=0.0, step=10.0, format="%.2f")
                    join_date   = r1c4.text_input("Join Date (e.g. Jan 2025)")
                    notes       = st.text_input("Notes")
                    submitted   = st.form_submit_button("Save Employee", type="primary")
                if submitted:
                    if not name:
                        st.error("Name is required.")
                    else:
                        emp_data[sub_dept].append({"Name": name, "Designation": designation, "Salary ($)": salary, "Join Date": join_date, "Notes": notes})
                        save_employees(emp_data)
                        st.success(f"Added {name} to {sub_dept}.")
                        st.rerun()

            # Delete — across both HR and Admin
            if all_emps:
                with st.expander("🗑 Remove an HR / Admin Employee"):
                    labels = [f"[{d}] {r.get('Name','?')}" for d, rows in [("HR", hr_emps), ("Admin", admin_emps)] for r in rows]
                    to_delete = st.selectbox("Select employee to remove", labels, key="del_hradmin")
                    if st.button(f"Remove {to_delete}", key="delbtn_hradmin"):
                        dept_tag, emp_name = to_delete.split("] ", 1)
                        dept_tag = dept_tag.lstrip("[")
                        emp_data[dept_tag] = [r for r in emp_data[dept_tag] if r.get("Name") != emp_name]
                        save_employees(emp_data)
                        st.success(f"Removed {emp_name} from {dept_tag}.")
                        st.rerun()


def page_expenses(data):
    st.title("🧾 Expenses")

    bust           = st.session_state.get("cache_bust", 0)
    months         = list(data.keys())
    manual_entries = load_expenses()
    all_months     = sorted(set(months + [e.get("Month", months[0] if months else "Jan 2026")
                                          for e in manual_entries]))

    exp_month  = st.selectbox("Select Month", all_months,
                               index=len(all_months) - 1, key="exp_month_view")
    excel_rows = load_expenses_for_month(_get_active_excel(), exp_month, _bust=bust)

    # ── Build records for the selected month ──────────────────────────────────
    xl_records = [{
        "Month":       exp_month,
        "Category":    r.get("Category", "Other Expenses"),
        "Description": str(r.get("Description", "")),
        "Date":        "",
        "Amount":      float(r.get("Amount", 0)),
        "Source":      "Excel",
    } for r in excel_rows]

    man_records = [{
        "Month":       e.get("Month", exp_month),
        "Category":    e.get("Category", ""),
        "Description": str(e.get("Description", "")),
        "Date":        str(e.get("Date", "")),
        "Amount":      float(e.get("Amount", 0)),
        "Source":      "Manual",
    } for e in manual_entries if e.get("Month", exp_month) == exp_month]

    all_records = xl_records + man_records

    # ── KPIs ──────────────────────────────────────────────────────────────────
    xl_total  = sum(r["Amount"] for r in xl_records)
    man_total = sum(r["Amount"] for r in man_records)
    combined  = xl_total + man_total

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Expenses (Excel)", _disp(xl_total,  exp_month, dec=2))
    k2.metric("Additional (Manual)",    _disp(man_total, exp_month, dec=2))
    k3.metric("Grand Total",            _disp(combined,  exp_month, dec=2))
    k4.metric("Line Items",             len(all_records))

    st.divider()

    if not all_records:
        st.info("No expense data found. Make sure the Excel file is accessible.")
        return

    combined_df = pd.DataFrame(all_records)

    # ── Filters ───────────────────────────────────────────────────────────────
    f1, f2, f3 = st.columns(3)
    cat_opts      = sorted(combined_df["Category"].dropna().unique().tolist())
    cat_filter    = f1.selectbox("Filter by Category", ["All"] + cat_opts, key="ef_cat")
    source_filter = f2.selectbox("Source", ["All", "Excel", "Manual"], key="ef_source")
    search_filter = f3.text_input("Search description", key="ef_search")

    view = combined_df.copy()
    if cat_filter != "All":
        view = view[view["Category"] == cat_filter]
    if source_filter != "All":
        view = view[view["Source"] == source_filter]
    if search_filter:
        view = view[view["Description"].str.contains(search_filter, case=False, na=False)]

    # ── INR conversion helpers ────────────────────────────────────────────────
    show_inr = st.session_state.get("show_inr", False)
    fx_rate  = get_fx_rate(exp_month) if show_inr else 1.0
    amt_fmt  = "₹{:,.2f}" if show_inr else "${:,.2f}"

    # ── Charts ────────────────────────────────────────────────────────────────
    chart_left, chart_right = st.columns(2)

    with chart_left:
        cat_totals = combined_df.groupby("Category")["Amount"].sum().reset_index().sort_values("Amount", ascending=False)
        fig_pie = px.pie(
            cat_totals, values="Amount", names="Category",
            hole=0.4,
            title="Expense Breakdown by Category",
            color_discrete_sequence=["#00d4aa","#7b5ea7","#b8f458","#ff6b6b","#ffd166","#a29bfe","#fd79a8","#55efc4","#fdcb6e","#6c5ce7"],
        )
        _chart_layout(fig_pie, height=360)
        st.plotly_chart(fig_pie, use_container_width=True)

    with chart_right:
        cat_bar = cat_totals.sort_values("Amount", ascending=True)
        fig_bar = px.bar(
            cat_bar, x="Amount", y="Category", orientation="h",
            title="Expense Amounts by Category",
            color="Amount", color_continuous_scale=["#0d2019","#00b894","#b8f458"],
            labels={"Amount": f"Amount"},
        )
        _chart_layout(fig_bar, height=360)
        st.plotly_chart(fig_bar, use_container_width=True)

    # ── Category deep-dive tabs ───────────────────────────────────────────────
    st.subheader("Category Detail")
    present_cats = [c for c in EXPENSE_CATS if c in combined_df["Category"].values]
    other_cats   = sorted(set(combined_df["Category"].unique()) - set(EXPENSE_CATS))
    tab_labels   = present_cats + other_cats
    if tab_labels:
        cat_tabs = st.tabs(tab_labels)
        for cat, ctab in zip(tab_labels, cat_tabs):
            with ctab:
                cat_df = combined_df[combined_df["Category"] == cat].copy()
                cc1, cc2 = st.columns(2)
                cc1.metric(f"Total {cat}", _disp(cat_df['Amount'].sum(), exp_month, dec=2))
                cc2.metric("Entries", len(cat_df))
                show_cols = ["Source", "Description", "Amount"]
                cat_show = cat_df[show_cols].copy()
                cat_show["Amount"] = cat_show["Amount"] * fx_rate
                cat_show.index = range(1, len(cat_show) + 1)
                st.dataframe(cat_show.style.format({"Amount": amt_fmt}), use_container_width=True, height=220)

    # ── Full table ────────────────────────────────────────────────────────────
    st.divider()
    with st.expander("📋 All Expense Line Items"):
        show_all_cols = ["Source", "Category", "Description", "Amount"]
        view_all = view[show_all_cols].copy()
        view_all["Amount"] = view_all["Amount"] * fx_rate
        view_all.index = range(1, len(view_all) + 1)
        st.dataframe(
            view_all.style.format({"Amount": amt_fmt}),
            use_container_width=True, height=400,
        )
        csv_all = view_all.to_csv(index=False).encode()
        st.download_button("Download CSV", csv_all, "expenses.csv", "text/csv", key="dl_exp")

    # ── Manual entry form ─────────────────────────────────────────────────────
    st.divider()
    with st.expander("➕ Add Manual Expense Entry"):
        with st.form("expense_form", clear_on_submit=True):
            c1, c2, c3 = st.columns(3)
            month    = c1.selectbox("Month", all_months if all_months else [exp_month], key="exp_month_sel")
            category = c2.selectbox("Category", EXPENSE_CATS)
            amount   = c3.number_input("Amount ($)", min_value=0.0, step=10.0, format="%.2f")
            c4, c5   = st.columns(2)
            description = c4.text_input("Description / Vendor")
            date_str    = c5.text_input("Date (e.g. 15 Mar 2025)")
            notes    = st.text_input("Notes")
            save_btn = st.form_submit_button("Save Expense", type="primary")

        if save_btn:
            if amount <= 0:
                st.error("Amount must be greater than 0.")
            else:
                manual_entries.append({
                    "Month": month, "Category": category,
                    "Description": description, "Date": date_str,
                    "Amount": amount, "Notes": notes,
                })
                save_expenses(manual_entries)
                st.success(f"Saved ${amount:,.2f} under {category} for {month}.")
                st.rerun()

    # ── Delete manual entry ───────────────────────────────────────────────────
    if manual_entries:
        with st.expander("🗑 Delete a Manual Entry"):
            labels = [
                f"{e['Month']} | {e['Category']} | {e.get('Description','—')} | ${float(e['Amount']):,.2f}"
                for e in manual_entries
            ]
            to_del = st.selectbox("Select entry", labels, key="del_exp_sel")
            if st.button("Delete selected entry", key="del_exp_btn"):
                idx = labels.index(to_del)
                manual_entries.pop(idx)
                save_expenses(manual_entries)
                st.success("Entry deleted.")
                st.rerun()


def page_import():
    st.title("📂 Import Excel File")
    st.write("Upload a new Excel file to add more months of data.")

    uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    if uploaded:
        import tempfile
        uploaded_bytes = uploaded.read()
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp.write(uploaded_bytes)
            tmp_path = tmp.name

        try:
            xl = pd.ExcelFile(tmp_path)
            # Month-suffix → full month name mapping
            _MONTH_MAP = {
                "Jan": "Jan 2026", "Feb": "Feb 2026", "Mar": "Mar 2026",
                "Apr": "Apr 2026", "May": "May 2026", "Jun": "Jun 2026",
                "Jul": "Jul 2026", "Aug": "Aug 2026", "Sep": "Sep 2026",
                "Oct": "Oct 2026", "Nov": "Nov 2026", "Dec": "Dec 2026",
            }
            def _suffix_to_month(raw_suffix):
                """Convert a sheet suffix like ' March26', 'Feb26', 'Jan' → 'Mar 2026'."""
                s = raw_suffix.strip().lower()
                for abbr, terms in _MONTH_SEARCH_TERMS.items():
                    if any(s.startswith(t) for t in terms):
                        return f"{abbr} 2026"
                return _MONTH_MAP.get(raw_suffix.strip(), raw_suffix.strip())

            billing_sheets = {}
            for sheet in xl.sheet_names:
                s_clean = sheet.lower().replace(" ", "")
                if sheet == "Clients":
                    billing_sheets[sheet] = "Jan 2026"
                elif s_clean.startswith("clients-"):
                    suffix = sheet.split("-", 1)[1]
                    billing_sheets[sheet] = _suffix_to_month(suffix)
            if not billing_sheets:
                st.warning("No billing sheets (named 'Clients' or 'Clients-<Month>') found in this file.")
            else:
                st.success(f"Found billing sheets: {', '.join(f'{s} → {m}' for s, m in billing_sheets.items())}")
                # Update the main Excel path and BILLING_SHEETS by writing a config file
                # For now: save these sheets into the app's BILLING_SHEETS via data.json only
                # for months NOT already covered by BILLING_SHEETS
                known_months = set(BILLING_SHEETS.values())
                custom = load_json_data()
                # Remove any stale imported entries first
                for key in list(custom.keys()):
                    if key.startswith("Clients") or key in ("Estimate till Sept", "Raw Data", "May", "March", "Feb-25"):
                        del custom[key]
                imported = 0
                for sheet, month_name in billing_sheets.items():
                    if month_name not in known_months:
                        df = load_excel_sheet(tmp_path, sheet)
                        custom[month_name] = df.to_dict(orient="records")
                        st.info(f"Imported {len(df)} rows → {month_name}")
                        imported += 1
                    else:
                        st.info(f"Skipped {sheet} ({month_name} is loaded from the primary Excel file)")
                save_json_data(custom)
                # Save file persistently so costs/expenses/salaries come from this file
                os.makedirs(UPLOADS_DIR, exist_ok=True)
                saved_path = os.path.join(UPLOADS_DIR, uploaded.name)
                with open(saved_path, "wb") as fp:
                    fp.write(uploaded_bytes)
                _save_active_excel(saved_path)
                st.cache_data.clear()
                if imported:
                    st.success(f"Import complete! {imported} new month(s) added. Costs/expenses will now be read from this file.")
        except Exception as e:
            st.error(f"Import failed: {e}")
        finally:
            os.unlink(tmp_path)


def _page_estimate_forward(data, months):
    # Use the most recent month as the cost baseline
    def _month_sort_key(m):
        _order = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        parts = m.split()
        mon = parts[0][:3] if parts else "Jan"
        yr  = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 2026
        return yr * 100 + (_order.index(mon) if mon in _order else 0)

    last_month = sorted(months, key=_month_sort_key)[-1]

    bust = st.session_state.get("cache_bust", 0)
    actual_costs = get_actual_cost_totals_usd(bust=bust)

    # ── Hero ───────────────────────────────────────────────────────────────────
    col_hero, col_ref = st.columns([3, 1])
    with col_hero:
        st.markdown(f"""<div class="dash-hero">
  <div class="dash-hero-title">Forward Estimate</div>
  <div class="dash-hero-sub">Project revenue &amp; profitability — costs based on {last_month}</div>
</div>""", unsafe_allow_html=True)
    with col_ref:
        st.markdown('<div style="height:36px"></div>', unsafe_allow_html=True)
        st.markdown(f"""<div style="
            background:rgba(0,212,170,0.06);border:1px solid rgba(0,212,170,0.22);
            border-radius:12px;padding:10px 14px;text-align:center;font-size:0.8rem;
            color:#7ecab0;">📌 Base month<br><b style="color:#00d4aa">{last_month}</b></div>""",
            unsafe_allow_html=True)

    # ── Revenue inputs — editable client table ─────────────────────────────────
    st.markdown("""<div class="dash-section-hdr" style="margin-top:10px">
  <span class="dash-section-title">Estimated Revenue</span>
  <span class="dash-section-sub">add or edit clients &amp; billing seats below</span>
</div>""", unsafe_allow_html=True)

    # Pre-fill from last month on first load
    if "est_clients" not in st.session_state:
        last_df = data[last_month].copy()
        prefill = (
            last_df.groupby("Client Name")
            .agg(Seats=("Seat Name", "count"), AvgRate=("Billing ($)", "mean"))
            .reset_index()
            .rename(columns={"Client Name": "Client", "Seats": "# Seats", "AvgRate": "Bill Rate ($/seat)"})
        )
        st.session_state["est_clients"] = prefill[["Client", "# Seats", "Bill Rate ($/seat)"]].to_dict("records")

    seed_df = pd.DataFrame(st.session_state["est_clients"])
    if seed_df.empty:
        seed_df = pd.DataFrame({"Client": [""], "# Seats": [1], "Bill Rate ($/seat)": [0.0]})
    seed_df["# Seats"] = pd.to_numeric(seed_df["# Seats"], errors="coerce").fillna(0).astype(int)
    seed_df["Bill Rate ($/seat)"] = pd.to_numeric(seed_df["Bill Rate ($/seat)"], errors="coerce").fillna(0.0)

    edited_clients = st.data_editor(
        seed_df,
        column_config={
            "Client":              st.column_config.TextColumn("Client Name", width="medium"),
            "# Seats":             st.column_config.NumberColumn("# Seats", min_value=0, step=1, format="%d"),
            "Bill Rate ($/seat)":  st.column_config.NumberColumn("Bill Rate ($/seat)", format="$%.2f", step=10.0, min_value=0.0),
        },
        num_rows="dynamic",
        use_container_width=True,
        height=230,
        key="est_clients_editor",
    )

    # Persist edits so variable cost changes don't reset the table
    valid_rows = edited_clients.dropna(subset=["Client"])
    valid_rows = valid_rows[valid_rows["Client"].astype(str).str.strip() != ""].copy()
    valid_rows["# Seats"] = pd.to_numeric(valid_rows["# Seats"], errors="coerce").fillna(0).astype(int)
    valid_rows["Bill Rate ($/seat)"] = pd.to_numeric(valid_rows["Bill Rate ($/seat)"], errors="coerce").fillna(0.0)
    valid_rows["Total Billing ($)"] = valid_rows["# Seats"] * valid_rows["Bill Rate ($/seat)"]

    est_revenue = float(valid_rows["Total Billing ($)"].sum())
    est_seats   = int(valid_rows["# Seats"].sum())

    # ── Variable cost sliders ──────────────────────────────────────────────────
    st.markdown("""<div class="dash-section-hdr" style="margin-top:18px">
  <span class="dash-section-title">Variable Cost Adjustments</span>
  <span class="dash-section-sub">these change as headcount grows</span>
</div>""", unsafe_allow_html=True)

    vc1, vc2, vc3 = st.columns(3)
    est_payroll = vc1.number_input(
        "Billing Payroll ($)",
        value=float(round(actual_costs.get("Salary/Payroll", 0), 2)),
        min_value=0.0, step=100.0, format="%.2f",
        help=f"Last month actual: ${actual_costs.get('Salary/Payroll', 0):,.2f}",
        key="est_payroll",
    )
    est_infra = vc2.number_input(
        "Infrastructure ($)",
        value=float(round(actual_costs.get("Infrastructure", 0), 2)),
        min_value=0.0, step=50.0, format="%.2f",
        help=f"Last month actual: ${actual_costs.get('Infrastructure', 0):,.2f}",
        key="est_infra",
    )
    est_meals = vc3.number_input(
        "Meals ($)",
        value=float(round(actual_costs.get("Meals", 0), 2)),
        min_value=0.0, step=50.0, format="%.2f",
        help=f"Last month actual: ${actual_costs.get('Meals', 0):,.2f}",
        key="est_meals",
    )

    # Fixed costs = everything else, locked to last month
    VARIABLE_KEYS = {"Salary/Payroll", "Infrastructure", "Meals"}
    fixed_costs    = {k: v for k, v in actual_costs.items() if k not in VARIABLE_KEYS}
    est_fixed_total = sum(fixed_costs.values())

    with st.expander("📌 Fixed Costs — carried from last month (read-only)", expanded=False):
        fc_cols = st.columns(3)
        for i, (k, v) in enumerate(fixed_costs.items()):
            fc_cols[i % 3].metric(k, f"${v:,.2f}")

    # ── Compute estimates ──────────────────────────────────────────────────────
    est_variable_total = est_payroll + est_infra + est_meals
    est_total_cost     = est_variable_total + est_fixed_total
    est_net            = est_revenue - est_total_cost
    est_margin         = (est_net / est_revenue * 100) if est_revenue else 0

    actual_rev    = float(data[last_month]["Billing ($)"].sum())
    actual_cost   = float(sum(actual_costs.values()))
    actual_net    = actual_rev - actual_cost

    rev_delta  = est_revenue - actual_rev
    net_delta  = est_net     - actual_net

    bal_color   = "#00d4aa" if est_net    >= 0 else "#ff5a00"
    bal_rgb     = "0,212,170" if est_net  >= 0 else "255,90,0"
    margin_color = "#00d4aa" if est_margin >= 0 else "#ff7a30"
    rev_d_color  = "#00d4aa" if rev_delta >= 0 else "#ff5a00"
    net_d_color  = "#00d4aa" if net_delta >= 0 else "#ff5a00"

    # ── KPI summary cards ──────────────────────────────────────────────────────
    st.markdown("""<div class="dash-section-hdr" style="margin-top:18px">
  <span class="dash-section-title">Estimated P&amp;L Summary</span>
  <span class="dash-section-sub">vs last month actuals</span>
</div>""", unsafe_allow_html=True)

    st.markdown(f"""<div class="dash-kpi-grid">
  <div class="dash-kpi-card">
    <div class="dash-kpi-icon" style="background:rgba(0,212,170,0.15);color:#00d4aa;">$</div>
    <div class="dash-kpi-body">
      <div class="dash-kpi-label">Est. Revenue</div>
      <div class="dash-kpi-value" style="color:#00d4aa">${est_revenue:,.0f}</div>
      <div class="dash-kpi-sub">{est_seats} seats · {len(valid_rows)} clients</div>
    </div>
  </div>
  <div class="dash-kpi-card">
    <div class="dash-kpi-icon" style="background:rgba(162,155,254,0.15);color:#a29bfe;">≡</div>
    <div class="dash-kpi-body">
      <div class="dash-kpi-label">Est. Total Cost</div>
      <div class="dash-kpi-value" style="color:#a29bfe">${est_total_cost:,.0f}</div>
      <div class="dash-kpi-sub">${est_variable_total:,.0f} variable + ${est_fixed_total:,.0f} fixed</div>
    </div>
  </div>
  <div class="dash-kpi-card" style="border-color:rgba({bal_rgb},0.28)">
    <div class="dash-kpi-icon" style="background:rgba({bal_rgb},0.15);color:{bal_color};">{'▲' if est_net >= 0 else '▼'}</div>
    <div class="dash-kpi-body">
      <div class="dash-kpi-label">Est. Net P&amp;L</div>
      <div class="dash-kpi-value" style="color:{bal_color}">${est_net:,.0f}</div>
      <div class="dash-kpi-sub" style="color:{margin_color}">{est_margin:.1f}% margin</div>
    </div>
  </div>
  <div class="dash-kpi-card">
    <div class="dash-kpi-icon" style="background:rgba(0,212,170,0.12);color:#00d4aa;">↗</div>
    <div class="dash-kpi-body">
      <div class="dash-kpi-label">Revenue vs {last_month}</div>
      <div class="dash-kpi-value" style="color:{rev_d_color}">${rev_delta:+,.0f}</div>
      <div class="dash-kpi-sub">was ${actual_rev:,.0f}</div>
    </div>
  </div>
  <div class="dash-kpi-card">
    <div class="dash-kpi-icon" style="background:rgba(184,244,88,0.12);color:#b8f458;">%</div>
    <div class="dash-kpi-body">
      <div class="dash-kpi-label">Net P&amp;L vs {last_month}</div>
      <div class="dash-kpi-value" style="color:{net_d_color}">${net_delta:+,.0f}</div>
      <div class="dash-kpi-sub">was ${actual_net:,.0f}</div>
    </div>
  </div>
</div>
<div style="height:22px"></div>
""", unsafe_allow_html=True)

    # ── Charts ─────────────────────────────────────────────────────────────────
    col_left, col_right = st.columns([3, 2])

    with col_left:
        st.markdown("""<div class="dash-section-hdr">
  <span class="dash-section-title">Estimated Revenue by Client</span>
  <span class="dash-section-sub">seats × rate</span>
</div>""", unsafe_allow_html=True)
        if not valid_rows.empty:
            chart_rev = valid_rows.sort_values("Total Billing ($)", ascending=True)
            fig_rev = px.bar(
                chart_rev, y="Client", x="Total Billing ($)", orientation="h",
                color="Total Billing ($)",
                color_continuous_scale=["#7b5ea7", "#00d4aa", "#b8f458"],
                text=[f"${v:,.0f}" for v in chart_rev["Total Billing ($)"]],
                labels={"Total Billing ($)": "Billing ($)", "Client": ""},
            )
            fig_rev.update_traces(textposition="outside", textfont=dict(color="#7ecab0", size=10))
            fig_rev.update_layout(coloraxis_showscale=False, showlegend=False)
            _chart_layout(fig_rev, height=max(260, len(chart_rev) * 44 + 60), xaxis_tickprefix="$")
            st.plotly_chart(fig_rev, use_container_width=True)
        else:
            st.info("Add clients above to see the revenue chart.")

    with col_right:
        st.markdown("""<div class="dash-section-hdr">
  <span class="dash-section-title">Cost: Estimate vs Last Month</span>
  <span class="dash-section-sub">by category</span>
</div>""", unsafe_allow_html=True)
        all_est_costs = {"Salary/Payroll": est_payroll, "Infrastructure": est_infra, "Meals": est_meals, **fixed_costs}
        comp_df = pd.DataFrame({
            "Category":        list(all_est_costs.keys()),
            "Estimate ($)":    list(all_est_costs.values()),
            "Last Month ($)":  [actual_costs.get(k, 0) for k in all_est_costs],
        })
        comp_df = comp_df[(comp_df["Estimate ($)"] + comp_df["Last Month ($)"]) > 0]
        comp_df = comp_df.sort_values("Estimate ($)", ascending=True)
        fig_cost = px.bar(
            comp_df, y="Category", x=["Estimate ($)", "Last Month ($)"],
            barmode="group", orientation="h",
            color_discrete_sequence=["#00d4aa", "#7b5ea7"],
            labels={"value": "Amount ($)", "variable": ""},
        )
        fig_cost.update_layout(
            legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1),
        )
        _chart_layout(fig_cost, height=400, xaxis_tickprefix="$")
        st.plotly_chart(fig_cost, use_container_width=True)

    # ── Reset ──────────────────────────────────────────────────────────────────
    st.divider()
    if st.button("🔄 Reset to Last Month Actuals"):
        for k in ("est_clients", "est_clients_editor", "est_payroll", "est_infra", "est_meals"):
            st.session_state.pop(k, None)
        st.rerun()

    # ── Break-Even Projector ───────────────────────────────────────────────────
    st.divider()
    st.markdown("""<div class="dash-section-hdr">
  <span class="dash-section-title">Break-Even Projector</span>
  <span class="dash-section-sub">add X billing seats each month — see when you reach profitability</span>
</div>""", unsafe_allow_html=True)

    # Defaults from current estimate (per-seat averages)
    _avg_bill  = est_revenue / est_seats if est_seats > 0 else 0.0
    _avg_pay   = est_payroll / est_seats if est_seats > 0 else 0.0
    _avg_infra = est_infra   / est_seats if est_seats > 0 else 0.0
    _avg_meals = est_meals   / est_seats if est_seats > 0 else 0.0

    bp1, bp2, bp3, bp4, bp5, bp6 = st.columns(6)
    _brc = st.session_state.get("bep_reset_count", 0)
    bep_seats_mo   = bp1.number_input("New Seats / Month",      value=1,   min_value=1, step=1, format="%d",   key=f"bep_seats_mo_{_brc}")
    bep_bill_rate  = bp2.number_input("Bill Rate / Seat ($)",   value=float(round(_avg_bill,  2)), min_value=0.0, step=50.0,  format="%.2f", key=f"bep_bill_rate_{_brc}")
    bep_pay_seat   = bp3.number_input("Payroll / Seat ($)",     value=float(round(_avg_pay,   2)), min_value=0.0, step=50.0,  format="%.2f", key=f"bep_pay_seat_{_brc}")
    bep_infra_seat = bp4.number_input("Infra / Seat ($)",       value=float(round(_avg_infra, 2)), min_value=0.0, step=10.0,  format="%.2f", key=f"bep_infra_seat_{_brc}")
    bep_meals_seat = bp5.number_input("Meals / Seat ($)",       value=float(round(_avg_meals, 2)), min_value=0.0, step=10.0,  format="%.2f", key=f"bep_meals_seat_{_brc}")
    bep_max_mo     = bp6.number_input("Max Months",             value=24,  min_value=6, max_value=60, step=6,   key=f"bep_max_mo_{_brc}")

    # ── Compute month labels from last_month ──────────────────────────────────
    _MO_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
    def _month_label(base, offset):
        parts = base.split()
        mi = _MO_NAMES.index(parts[0][:3]) if parts[0][:3] in _MO_NAMES else 0
        yr = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 2026
        ni = mi + offset
        return f"{_MO_NAMES[ni % 12]} {yr + ni // 12}"

    # ── Build projection rows ─────────────────────────────────────────────────
    rows_bep = []
    bep_found = None
    for n in range(bep_max_mo + 1):
        added   = n * bep_seats_mo
        t_seats = est_seats + added
        rev     = est_revenue + added * bep_bill_rate
        pay     = est_payroll + added * bep_pay_seat
        infra   = est_infra   + added * bep_infra_seat
        meals   = est_meals   + added * bep_meals_seat
        var_c   = pay + infra + meals
        tot_c   = var_c + est_fixed_total
        net     = rev - tot_c
        margin  = (net / rev * 100) if rev else 0
        label   = _month_label(last_month, n) + (" (now)" if n == 0 else "")
        if bep_found is None and net >= 0:
            bep_found = n
        rows_bep.append({
            "Month":              label,
            "Seats Added":        added,
            "Total Seats":        t_seats,
            "Revenue ($)":        round(rev,   0),
            "Variable Cost ($)":  round(var_c, 0),
            "Fixed Cost ($)":     round(est_fixed_total, 0),
            "Total Cost ($)":     round(tot_c, 0),
            "Net P&L ($)":        round(net,   0),
            "Margin %":           round(margin, 1),
        })

    bep_df = pd.DataFrame(rows_bep)

    # ── Break-even KPI banner ─────────────────────────────────────────────────
    if bep_found == 0:
        bep_label = "Already profitable"
        bep_kpi_color = "#00d4aa"
    elif bep_found is not None:
        bep_label = f"{bep_found} month{'s' if bep_found > 1 else ''}"
        bep_kpi_color = "#b8f458" if bep_found <= 6 else "#ffd166" if bep_found <= 12 else "#ff7a30"
    else:
        bep_label = f"Not reached in {bep_max_mo} months"
        bep_kpi_color = "#ff5a00"

    bep_row = rows_bep[bep_found] if bep_found is not None else None

    bk1, bk2, bk3, bk4 = st.columns(4)
    bk1.metric("Break-Even",             bep_label)
    bk2.metric("Seats at Break-Even",    str(bep_row["Total Seats"]) if bep_row else "—")
    bk3.metric("Revenue at Break-Even",  f"${bep_row['Revenue ($)']:,.0f}" if bep_row else "—")
    bk4.metric("Net P&L at Break-Even",  f"${bep_row['Net P&L ($)']:+,.0f}" if bep_row else "—")

    # ── Line chart ────────────────────────────────────────────────────────────
    fig_bep = go.Figure()
    fig_bep.add_trace(go.Scatter(
        x=bep_df["Month"], y=bep_df["Revenue ($)"],
        mode="lines+markers", name="Revenue",
        line=dict(color="#00d4aa", width=2.5), marker=dict(size=5),
    ))
    fig_bep.add_trace(go.Scatter(
        x=bep_df["Month"], y=bep_df["Total Cost ($)"],
        mode="lines+markers", name="Total Cost",
        line=dict(color="#7b5ea7", width=2.5), marker=dict(size=5),
    ))
    fig_bep.add_trace(go.Scatter(
        x=bep_df["Month"], y=bep_df["Net P&L ($)"],
        mode="lines+markers", name="Net P&L",
        line=dict(color="#b8f458", width=2, dash="dot"), marker=dict(size=4),
    ))
    fig_bep.add_hline(y=0, line_dash="dash", line_color="rgba(0,212,170,0.35)", line_width=1)
    if bep_found is not None and bep_found > 0:
        bep_x = bep_df.iloc[bep_found]["Month"]
        fig_bep.add_vline(x=bep_x, line_dash="dash", line_color="rgba(0,212,170,0.45)", line_width=1.5)
        fig_bep.add_annotation(
            x=bep_x, y=bep_row["Revenue ($)"],
            text=f"  Break-even: {bep_label}",
            showarrow=True, arrowhead=2, arrowcolor="#00d4aa",
            font=dict(color="#00d4aa", size=11),
            bgcolor="rgba(0,20,15,0.85)", bordercolor="#00d4aa", borderwidth=1,
            xanchor="left",
        )
    _chart_layout(fig_bep, height=360, yaxis_tickprefix="$")
    fig_bep.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1))
    st.plotly_chart(fig_bep, use_container_width=True)

    # ── Projection table ──────────────────────────────────────────────────────
    def _bep_row_style(row):
        n = bep_df.index.get_loc(row.name)
        if bep_found is not None and n == bep_found and bep_found > 0:
            return ["background:rgba(0,212,170,0.18);color:#00d4aa;font-weight:bold"] * len(row)
        return ["color:#00d4aa" if row["Net P&L ($)"] >= 0 else "color:#ff7a60"] * len(row)

    styled_bep = (
        bep_df.style
        .apply(_bep_row_style, axis=1)
        .format({
            "Revenue ($)":       "${:,.0f}",
            "Variable Cost ($)": "${:,.0f}",
            "Fixed Cost ($)":    "${:,.0f}",
            "Total Cost ($)":    "${:,.0f}",
            "Net P&L ($)":       "${:+,.0f}",
            "Margin %":          "{:.1f}%",
        })
    )
    st.dataframe(styled_bep, use_container_width=True,
                 height=min(600, (len(bep_df) + 1) * 36 + 38))

    if st.button("🔄 Reset Break-Even Inputs"):
        st.session_state["bep_reset_count"] = st.session_state.get("bep_reset_count", 0) + 1
        st.rerun()


def page_estimate(data):
    st.title("🔮 Forward Estimate")

    months = list(data.keys())
    if not months:
        st.warning("No data loaded. Import an Excel file first.")
        return

    _page_estimate_forward(data, months)


def _page_sales_dept_data(data):
    st.title("💼 Sales Dept Data")
    MONTHS = ["Oct-25", "Nov-25", "Dec-25", "Jan-26", "Feb-26", "Mar-26", "Apr-26"]

    _RAW = [
        ("Rent",         "Regus Management",   0.00,     2362.99,  1646.26,  1692.11,  1660.21,  1645.91,  2735.01),
        ("Food & Drink", "Food Exps",          398.22,    661.86,    56.73,   292.96,   711.23,  1545.06,   362.30),
        ("Marketing",    "NATHO",               0.00,   1500.00,     0.00,  2695.00,     0.00,     0.00,     0.00),
        ("Marketing",    "STAFFING IND AN",  2574.00,      0.00,     0.00,     0.00,     0.00,     0.00,     0.00),
        ("Marketing",    "ASC COMMUNICA",       0.00,      0.00,     0.00,     0.00, 17500.00,     0.00,  2500.00),
        ("Marketing",    "MINUTEMAN PRES",      0.00,      0.00,     0.00,     0.00,    78.40,     0.00,     0.00),
        ("Travel",       "Travelling Exps",  2011.58,   1350.77,   533.20,  1650.66,  2992.14,  2590.64,  1957.33),
        ("Shipping Charges", "Shipping Charges",  0.00,      0.00,     0.00,     0.00,     0.00,   287.50,  3080.01),
        ("Software",     "ChatGpt",             0.00,      0.00,    60.00,    60.00,    60.00,    60.00,    44.13),
        ("Software",     "SALESFORCE",        210.00,    125.00,   125.00,   125.00,     0.00,   241.25,     0.00),
        ("Salary",       "Brett Williams",   4027.75,  11114.52, 11114.68, 10416.66, 10416.66, 10416.66, 10416.66),
        ("Salary",       "Shravan",          1618.88,   1618.88,  1618.88,  1618.88,  1618.88,  1618.88,  1618.88),
        ("Salary",       "Dominic",           730.00,    730.00,   730.00,   730.00,   730.00,   730.00,   730.00),
        ("Salary",       "Jen",                 0.00,      0.00,     0.00,  3906.50,  5000.00,  5000.00,  5000.00),
        ("Salary",       "Kenzie",              0.00,      0.00,     0.00,  2604.34,  3333.30,  3333.30,  3333.30),
    ]

    detail_df = pd.DataFrame(_RAW, columns=["Categories", "Particulars"] + MONTHS)

    fmt_dict = {m: "${:,.2f}" for m in MONTHS}

    # ── Brett: Cost vs Revenue table ──────────────────────────────────────────
    st.markdown("""<div class="dash-section-hdr" style="margin-top:4px">
  <span class="dash-section-title">Sales Performance — Brett Williams</span>
  <span class="dash-section-sub">cost vs revenue brought in per month</span>
</div>""", unsafe_allow_html=True)

    def _shortmon_to_key(m):
        mo, yr = m.split("-")
        return f"{mo} 20{yr}"   # "Oct-25" → "Oct 2025"

    # Brett's clients are those assigned to him in sales_persons.json
    sales_map = load_sales_persons()
    brett_clients = {
        client for client, sp in sales_map.items()
        if "brett" in sp.strip().lower()
    }

    perf_rows = []
    for mo in MONTHS:
        cost = float(detail_df[mo].sum())
        data_key = _shortmon_to_key(mo)
        revenue = 0.0
        if data_key in data and brett_clients:
            mdf = data[data_key]
            mask = mdf["Client Name"].isin(brett_clients)
            revenue = float(mdf.loc[mask, "Billing ($)"].sum())
        net = revenue - cost
        perf_rows.append({
            "Month":           mo,
            "Total Cost ($)":  cost,
            "Revenue ($)":     revenue,
            "Net ($)":         net,
        })

    perf_df = pd.DataFrame(perf_rows)

    total_cost    = perf_df["Total Cost ($)"].sum()
    total_revenue = perf_df["Revenue ($)"].sum()
    total_net     = total_revenue - total_cost
    total_row = pd.DataFrame([{
        "Month":           "Total",
        "Total Cost ($)":  total_cost,
        "Revenue ($)":     total_revenue,
        "Net ($)":         total_net,
    }])
    perf_with_total = pd.concat([perf_df, total_row], ignore_index=True)

    def _style_perf(df):
        out = [list([""] * len(df.columns)) for _ in range(len(df))]
        for i, row in df.iterrows():
            if row["Month"] == "Total":
                base = "font-weight:bold; border-top:2px solid #00d4aa; color:#00d4aa"
                net_color = base if row["Net ($)"] >= 0 else base.replace("#00d4aa", "#ff5a00")
                out[i] = [base, base, base, net_color]
            else:
                net_color = "color:#00d4aa" if row["Net ($)"] >= 0 else "color:#ff5a00"
                out[i][3] = net_color
        return pd.DataFrame(out, columns=df.columns)

    styled_perf = (
        perf_with_total.style
        .apply(_style_perf, axis=None)
        .format({
            "Total Cost ($)": "${:,.2f}",
            "Revenue ($)":    "${:,.2f}",
            "Net ($)":        "${:+,.2f}",
        })
        .hide(axis="index")
    )
    st.dataframe(styled_perf, use_container_width=True,
                 height=(len(perf_with_total) + 1) * 36 + 38)

    # ── Category summary ──────────────────────────────────────────────────────
    st.markdown("""<div class="dash-section-hdr" style="margin-top:4px">
  <span class="dash-section-title">Category Summary</span>
  <span class="dash-section-sub">aggregated totals per month</span>
</div>""", unsafe_allow_html=True)

    summary_df = detail_df.groupby("Categories")[MONTHS].sum().reset_index()
    grand_row = {"Categories": "Grand Total", **{m: summary_df[m].sum() for m in MONTHS}}
    summary_with_total = pd.concat([summary_df, pd.DataFrame([grand_row])], ignore_index=True)

    def _style_summary(df):
        out = [list([""] * len(df.columns)) for _ in range(len(df))]
        for i, row in df.iterrows():
            if row["Categories"] == "Grand Total":
                out[i] = ["font-weight:bold; border-top:2px solid #00d4aa; color:#00d4aa"] * len(df.columns)
        return pd.DataFrame(out, columns=df.columns)

    styled_summary = (
        summary_with_total.style
        .apply(_style_summary, axis=None)
        .format(fmt_dict)
        .hide(axis="index")
    )
    st.dataframe(styled_summary, use_container_width=True,
                 height=min(400, (len(summary_with_total) + 1) * 36 + 38))

    # ── Detailed breakdown ────────────────────────────────────────────────────
    st.markdown("""<div class="dash-section-hdr" style="margin-top:18px">
  <span class="dash-section-title">Detailed Breakdown</span>
  <span class="dash-section-sub">by category &amp; particulars</span>
</div>""", unsafe_allow_html=True)

    totals_row = {"Categories": "", "Particulars": "Total", **{m: detail_df[m].sum() for m in MONTHS}}
    detail_with_total = pd.concat([detail_df, pd.DataFrame([totals_row])], ignore_index=True)

    def _style_detail(df):
        out = [list([""] * len(df.columns)) for _ in range(len(df))]
        for i, row in df.iterrows():
            if row["Particulars"] == "Total":
                out[i] = ["font-weight:bold; border-top:2px solid #00d4aa; color:#00d4aa"] * len(df.columns)
        return pd.DataFrame(out, columns=df.columns)

    styled_detail = (
        detail_with_total.style
        .apply(_style_detail, axis=None)
        .format(fmt_dict)
        .hide(axis="index")
    )
    st.dataframe(styled_detail, use_container_width=True,
                 height=min(640, (len(detail_with_total) + 1) * 36 + 38))

    # ── Trend chart ───────────────────────────────────────────────────────────
    st.markdown("""<div class="dash-section-hdr" style="margin-top:18px">
  <span class="dash-section-title">Spend Trend by Category</span>
  <span class="dash-section-sub">stacked bar — monthly view</span>
</div>""", unsafe_allow_html=True)

    _CAT_COLORS = {
        "Salary":            "#7b5ea7",
        "Marketing":         "#00d4aa",
        "Travel":            "#ffd166",
        "Rent":              "#a29bfe",
        "Software":          "#74b9ff",
        "Food & Drink":      "#fd79a8",
        "Shipping Charges":  "#b8f458",
    }

    fig_trend = go.Figure()
    for cat in summary_df["Categories"].tolist():
        row_vals = summary_df.loc[summary_df["Categories"] == cat, MONTHS].iloc[0].tolist()
        fig_trend.add_trace(go.Bar(
            name=cat,
            x=MONTHS,
            y=row_vals,
            marker_color=_CAT_COLORS.get(cat, "#888888"),
        ))
    fig_trend.update_layout(barmode="stack")
    _chart_layout(fig_trend, height=360, yaxis_tickprefix="$")
    fig_trend.update_layout(
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1)
    )
    st.plotly_chart(fig_trend, use_container_width=True)

    # ── KPI row — monthly totals ───────────────────────────────────────────────
    st.markdown("""<div class="dash-section-hdr" style="margin-top:18px">
  <span class="dash-section-title">Monthly Total Spend</span>
  <span class="dash-section-sub">all categories combined</span>
</div>""", unsafe_allow_html=True)

    kpi_cols = st.columns(len(MONTHS))
    for i, mo in enumerate(MONTHS):
        kpi_cols[i].metric(mo, f"${detail_df[mo].sum():,.0f}")


# ── CSS ───────────────────────────────────────────────────────────────────────

def _inject_css():
    st.markdown("""
<style>
/* ── Base & background ────────────────────────────────────────── */
html, body, [class*="css"] {
    font-family: 'Inter', 'Segoe UI', system-ui, sans-serif;
}
.stApp {
    background: radial-gradient(ellipse at 20% 0%, #0d2b22 0%, #080f1c 55%, #060d17 100%) !important;
}
/* Main content area */
.block-container {
    padding-top: 1.8rem !important;
    padding-bottom: 2rem !important;
}

/* ── Page titles ──────────────────────────────────────────────── */
h1 {
    background: linear-gradient(135deg, #00d4aa 0%, #7b5ea7 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    font-weight: 800 !important;
    letter-spacing: -0.5px;
}
h2 { color: #b8f0e0 !important; font-weight: 600 !important; }
h3 { color: #7ecab0 !important; font-weight: 500 !important; }

/* ── Metric cards ─────────────────────────────────────────────── */
[data-testid="metric-container"] {
    background: linear-gradient(145deg, #0f2231 0%, #132d24 100%);
    border: 1px solid rgba(0, 212, 170, 0.2);
    border-radius: 18px;
    padding: 20px 24px 16px !important;
    box-shadow: 0 6px 28px rgba(0, 0, 0, 0.45),
                inset 0 1px 0 rgba(0, 212, 170, 0.07);
    transition: transform 0.22s ease, box-shadow 0.22s ease, border-color 0.22s ease;
    position: relative;
    overflow: hidden;
}
[data-testid="metric-container"]::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, #00d4aa, #7b5ea7, #b8f458);
    border-radius: 18px 18px 0 0;
}
[data-testid="metric-container"]:hover {
    transform: translateY(-4px);
    box-shadow: 0 12px 36px rgba(0, 212, 170, 0.2);
    border-color: rgba(0, 212, 170, 0.45);
}
[data-testid="stMetricLabel"] > div {
    font-size: 0.72rem !important;
    font-weight: 700 !important;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: #5a9a86 !important;
}
[data-testid="stMetricValue"] > div {
    font-size: 1.8rem !important;
    font-weight: 800 !important;
    background: linear-gradient(135deg, #00d4aa 0%, #b8f458 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    line-height: 1.15;
}
[data-testid="stMetricDelta"] {
    font-size: 0.82rem !important;
    font-weight: 600 !important;
}

/* ── Top header bar ───────────────────────────────────────────── */
header[data-testid="stHeader"] {
    background: linear-gradient(180deg, #06150f 0%, #080f1c 100%) !important;
    border-bottom: 1px solid rgba(0, 212, 170, 0.12) !important;
}
[data-testid="stToolbar"] { background: transparent !important; }
[data-testid="stDecoration"] {
    background: linear-gradient(90deg, #00d4aa, #7b5ea7, #b8f458) !important;
    height: 2px !important;
}

/* ── Sidebar ──────────────────────────────────────────────────── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #06150f 0%, #0b1f17 50%, #060e0a 100%) !important;
    border-right: 1px solid rgba(0, 212, 170, 0.15) !important;
}
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] .stMarkdown h1 {
    font-size: 1rem !important;
    color: #00d4aa !important;
    -webkit-text-fill-color: #00d4aa !important;
    background: none !important;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}
/* All general sidebar text */
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] small,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stMarkdown p {
    color: #b8f0e0 !important;
}
/* Hide the "Go to" label text — nav items speak for themselves */
[data-testid="stSidebarContent"] .stRadio > label {
    display: none !important;
}

/* Stack items vertically with tight gap */
[data-testid="stSidebarContent"] .stRadio > div[role="radiogroup"] {
    display: flex !important;
    flex-direction: column !important;
    gap: 3px !important;
}

/* Hide the circular radio dot — we use full-row click area instead */
[data-testid="stSidebarContent"] .stRadio [data-baseweb="radio"] > div:first-child {
    display: none !important;
}

/* Each nav item base style */
[data-testid="stSidebarContent"] .stRadio label[data-baseweb="radio"] {
    width: 100% !important;
    border-radius: 11px !important;
    padding: 10px 14px !important;
    border: 1px solid rgba(0, 212, 170, 0.07) !important;
    border-left: 3px solid transparent !important;
    background: transparent !important;
    color: #5a9a86 !important;
    cursor: pointer !important;
    font-size: 0.87rem !important;
    font-weight: 500 !important;
    transition: all 0.18s ease !important;
    gap: 0 !important;
}

/* Hover */
[data-testid="stSidebarContent"] .stRadio label[data-baseweb="radio"]:hover {
    background: rgba(0, 212, 170, 0.08) !important;
    border-color: rgba(0, 212, 170, 0.22) !important;
    border-left-color: rgba(0, 212, 170, 0.5) !important;
    color: #b8f0e0 !important;
    transform: translateX(3px) !important;
}

/* Selected / active item */
[data-testid="stSidebarContent"] .stRadio label[data-baseweb="radio"]:has(input:checked) {
    background: linear-gradient(135deg, rgba(0,212,170,0.16) 0%, rgba(0,160,128,0.08) 100%) !important;
    border-color: rgba(0, 212, 170, 0.28) !important;
    border-left: 3px solid #00d4aa !important;
    color: #00d4aa !important;
    font-weight: 700 !important;
    box-shadow: 0 2px 14px rgba(0, 212, 170, 0.12),
                inset 0 1px 0 rgba(0, 212, 170, 0.06) !important;
    transform: translateX(2px) !important;
}

/* Click press */
[data-testid="stSidebarContent"] .stRadio label[data-baseweb="radio"]:active {
    transform: translateX(1px) scale(0.99) !important;
}

/* ── Buttons ──────────────────────────────────────────────────── */
.stButton > button {
    background: linear-gradient(135deg, #00b894 0%, #00d4aa 100%) !important;
    color: #040f0b !important;
    border: none !important;
    border-radius: 50px !important;
    font-weight: 700 !important;
    font-size: 0.85rem !important;
    letter-spacing: 0.04em;
    padding: 8px 14px !important;
    white-space: nowrap !important;
    box-shadow: 0 4px 18px rgba(0, 212, 170, 0.4) !important;
    transition: all 0.25s ease !important;
}
.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 24px rgba(0, 212, 170, 0.6) !important;
    filter: brightness(1.08) !important;
}
.stButton > button:active { transform: translateY(0) !important; }

.stDownloadButton > button {
    background: transparent !important;
    border: 1px solid rgba(0, 212, 170, 0.4) !important;
    color: #00d4aa !important;
    border-radius: 50px !important;
    font-weight: 600 !important;
    box-shadow: none !important;
    transition: all 0.2s ease !important;
}
.stDownloadButton > button:hover {
    background: rgba(0, 212, 170, 0.1) !important;
    border-color: #00d4aa !important;
    transform: translateY(-1px) !important;
}

/* ── Tabs ─────────────────────────────────────────────────────── */
.stTabs [data-baseweb="tab-list"] {
    background: rgba(0, 212, 170, 0.05) !important;
    border-radius: 50px !important;
    padding: 5px 6px !important;
    gap: 4px !important;
    border: 1px solid rgba(0, 212, 170, 0.15) !important;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 50px !important;
    color: #5a9a86 !important;
    font-weight: 500 !important;
    font-size: 0.87rem !important;
    padding: 8px 22px !important;
    transition: color 0.2s !important;
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(135deg, #00b894 0%, #00d4aa 100%) !important;
    color: #040f0b !important;
    font-weight: 700 !important;
    box-shadow: 0 3px 12px rgba(0, 212, 170, 0.45) !important;
}

/* ── Form labels (all pages) ──────────────────────────────────── */
.stSelectbox > label,
.stMultiSelect > label,
.stTextInput > label,
.stNumberInput > label,
.stTextArea > label,
.stFileUploader > label,
.stDateInput > label,
[data-testid="stWidgetLabel"],
[data-testid="stWidgetLabel"] p {
    color: #7ecab0 !important;
    font-size: 0.72rem !important;
    font-weight: 700 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.09em !important;
    margin-bottom: 4px !important;
}

/* ── Select / Input ───────────────────────────────────────────── */
.stSelectbox > div > div,
.stMultiSelect > div > div {
    background: #0d2019 !important;
    border: 1px solid rgba(0, 212, 170, 0.25) !important;
    border-radius: 12px !important;
}
.stSelectbox > div > div:focus-within {
    border-color: #00d4aa !important;
    box-shadow: 0 0 0 2px rgba(0, 212, 170, 0.18) !important;
}
/* Selectbox selected value text — force readable color */
div[data-baseweb="select"] span,
div[data-baseweb="select"] div[class],
.stSelectbox [data-baseweb="select"] [aria-selected="true"],
.stSelectbox [data-baseweb="select"] > div > div {
    color: #d4f5ec !important;
}
/* Dropdown list options */
ul[data-baseweb="menu"] li {
    background: #0d2019 !important;
    color: #b8f0e0 !important;
}
ul[data-baseweb="menu"] li:hover {
    background: rgba(0, 212, 170, 0.12) !important;
}
.stTextInput > div > div > input,
.stNumberInput > div > div > input {
    background: #0d2019 !important;
    border: 1px solid rgba(0, 212, 170, 0.25) !important;
    border-radius: 12px !important;
    color: #d4f5ec !important;
}
.stTextInput > div > div > input:focus,
.stNumberInput > div > div > input:focus {
    border-color: #00d4aa !important;
    box-shadow: 0 0 0 2px rgba(0, 212, 170, 0.18) !important;
}

/* ── Expander ─────────────────────────────────────────────────── */
.streamlit-expanderHeader {
    background: rgba(0, 212, 170, 0.06) !important;
    border: 1px solid rgba(0, 212, 170, 0.18) !important;
    border-radius: 12px !important;
    font-weight: 600 !important;
    font-size: 0.88rem !important;
    color: #7ecab0 !important;
    transition: all 0.2s !important;
}
.streamlit-expanderHeader:hover {
    background: rgba(0, 212, 170, 0.1) !important;
    border-color: rgba(0, 212, 170, 0.4) !important;
}
.streamlit-expanderContent {
    background: #071510 !important;
    border: 1px solid rgba(0, 212, 170, 0.12) !important;
    border-top: none !important;
    border-radius: 0 0 12px 12px !important;
}

/* ── Dividers ─────────────────────────────────────────────────── */
hr {
    border: none !important;
    height: 1px !important;
    background: linear-gradient(90deg, transparent, rgba(0,212,170,0.4), transparent) !important;
    margin: 1.4rem 0 !important;
}

/* ── Alerts ───────────────────────────────────────────────────── */
.stSuccess > div {
    background: rgba(0, 212, 170, 0.08) !important;
    border-left: 3px solid #00d4aa !important;
    border-radius: 0 12px 12px 0 !important;
    color: #7ecab0 !important;
}
.stInfo > div {
    background: rgba(123, 94, 167, 0.1) !important;
    border-left: 3px solid #7b5ea7 !important;
    border-radius: 0 12px 12px 0 !important;
    color: #c4b0e0 !important;
}
.stWarning > div {
    background: rgba(184, 244, 88, 0.08) !important;
    border-left: 3px solid #b8f458 !important;
    border-radius: 0 12px 12px 0 !important;
    color: #d6f7a0 !important;
}
.stError > div {
    background: rgba(255, 107, 107, 0.1) !important;
    border-left: 3px solid #ff6b6b !important;
    border-radius: 0 12px 12px 0 !important;
    color: #ffaaaa !important;
}

/* ── Captions ─────────────────────────────────────────────────── */
.stCaption, small { color: #3d7a66 !important; font-size: 0.76rem !important; }

/* ── KPI Nav Cards (Employee Details) — hover effect only ────── */
/* .kpi-nav-card class is added by JS to the 4 clickable metrics  */
.kpi-nav-card {
    cursor: pointer !important;
    transition: transform 0.22s ease, box-shadow 0.22s ease,
                border-color 0.22s ease !important;
}
.kpi-nav-card:hover {
    transform: translateY(-6px) !important;
    box-shadow: 0 14px 36px rgba(0,212,170,0.25) !important;
    border-color: rgba(0,212,170,0.5) !important;
}

/* ── Section radio (Employee Details tab switcher) ────────────── */
.block-container .stRadio > div[role="radiogroup"] {
    display: flex !important;
    flex-direction: row !important;
    gap: 4px !important;
    background: rgba(0,212,170,0.05) !important;
    border: 1px solid rgba(0,212,170,0.15) !important;
    border-radius: 50px !important;
    padding: 5px 6px !important;
}
.block-container .stRadio label[data-baseweb="radio"] {
    border-radius: 50px !important;
    padding: 8px 22px !important;
    font-size: 0.87rem !important;
    font-weight: 500 !important;
    color: #5a9a86 !important;
    cursor: pointer !important;
    transition: all 0.18s ease !important;
    border: none !important;
    background: transparent !important;
}
.block-container .stRadio label[data-baseweb="radio"]:hover {
    color: #b8f0e0 !important;
    background: rgba(0,212,170,0.08) !important;
}
.block-container .stRadio label[data-baseweb="radio"]:has(input:checked) {
    background: linear-gradient(135deg,#00b894 0%,#00d4aa 100%) !important;
    color: #040f0b !important;
    font-weight: 700 !important;
    box-shadow: 0 3px 12px rgba(0,212,170,0.45) !important;
}
.block-container .stRadio [data-baseweb="radio"] > div:first-child {
    display: none !important;
}

/* ── Forms ────────────────────────────────────────────────────── */
[data-testid="stForm"] {
    background: rgba(0, 212, 170, 0.04) !important;
    border: 1px solid rgba(0, 212, 170, 0.18) !important;
    border-radius: 16px !important;
    padding: 20px !important;
}

/* ── Dataframes ───────────────────────────────────────────────── */
/* Only style the container border — theme config handles cell colors */
[data-testid="stDataFrame"] > div {
    border-radius: 12px !important;
    overflow: hidden;
    border: 1px solid rgba(0, 212, 170, 0.12) !important;
}
[data-testid="stDataEditor"] > div {
    border-radius: 12px !important;
    overflow: hidden;
    border: 1px solid rgba(0, 212, 170, 0.12) !important;
}

/* ── Scrollbar ────────────────────────────────────────────────── */
::-webkit-scrollbar { width: 5px; height: 5px; }
::-webkit-scrollbar-track { background: #06110d; }
::-webkit-scrollbar-thumb {
    background: linear-gradient(180deg, #00b894, #7b5ea7);
    border-radius: 4px;
}
::-webkit-scrollbar-thumb:hover { background: #00d4aa; }

/* ── Dashboard hero ───────────────────────────────────────────── */
.dash-hero {
    padding: 32px 0 18px 0;
}
.dash-hero-title {
    font-size: 2rem;
    font-weight: 900;
    line-height: 1.15;
    background: linear-gradient(135deg, #00d4aa 0%, #b8f458 70%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin-bottom: 6px;
}
.dash-hero-sub {
    font-size: 0.85rem;
    color: #5ab896;
    font-weight: 400;
}

/* ── Dashboard KPI grid ───────────────────────────────────────── */
.dash-kpi-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 14px;
    margin-bottom: 4px;
}
@media (max-width: 960px) { .dash-kpi-grid { grid-template-columns: repeat(3, 1fr); } }
@media (max-width: 640px)  { .dash-kpi-grid { grid-template-columns: repeat(2, 1fr); } }

.dash-kpi-card {
    background: linear-gradient(150deg, #0c2030 0%, #0d2820 100%);
    border: 1px solid rgba(0, 212, 170, 0.16);
    border-radius: 18px;
    padding: 18px 16px;
    display: flex;
    align-items: center;
    gap: 14px;
    box-shadow: 0 6px 28px rgba(0,0,0,0.5);
    position: relative;
    overflow: hidden;
    transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
}
.dash-kpi-card::before {
    content: "";
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: linear-gradient(90deg, transparent, rgba(0,212,170,0.5), transparent);
    border-radius: 18px 18px 0 0;
}
.dash-kpi-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 14px 36px rgba(0, 212, 170, 0.14);
    border-color: rgba(0, 212, 170, 0.3);
}
.dash-kpi-icon {
    min-width: 46px;
    height: 46px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.2rem;
    font-weight: 900;
    flex-shrink: 0;
}
.dash-kpi-body { flex: 1; min-width: 0; }
.dash-kpi-label {
    font-size: 0.62rem;
    text-transform: uppercase;
    letter-spacing: 0.12em;
    color: #3d7a66;
    font-weight: 700;
    margin-bottom: 4px;
    white-space: nowrap;
}
.dash-kpi-value {
    font-size: 1.32rem;
    font-weight: 800;
    line-height: 1.1;
    margin-bottom: 4px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.dash-kpi-sub {
    font-size: 0.68rem;
    color: #3a6658;
    font-weight: 500;
}

/* ── Dashboard section headers ────────────────────────────────── */
.dash-section-hdr {
    display: flex;
    align-items: baseline;
    gap: 10px;
    margin: 2px 0 6px 0;
    padding-bottom: 8px;
    border-bottom: 1px solid rgba(0, 212, 170, 0.1);
}
.dash-section-title {
    font-size: 0.78rem !important;
    font-weight: 700 !important;
    color: #7ecab0 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.1em !important;
}
.dash-section-sub {
    font-size: 0.68rem !important;
    color: #3a6658 !important;
    font-weight: 400 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
}
</style>
""", unsafe_allow_html=True)


def _chart_layout(fig, **kwargs):
    """Apply the shared dark-teal theme to any Plotly figure."""
    kwargs.setdefault("title", "")
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(13,32,25,0.5)",
        font=dict(family="Inter, Segoe UI, sans-serif", color="#7ecab0"),
        title_font=dict(size=14, color="#b8f0e0"),
        legend=dict(
            bgcolor="rgba(0,0,0,0)",
            bordercolor="rgba(0,212,170,0.15)",
            borderwidth=1,
            font=dict(color="#7ecab0"),
        ),
        xaxis=dict(
            gridcolor="rgba(0,212,170,0.08)",
            zerolinecolor="rgba(0,212,170,0.15)",
            tickfont=dict(color="#5a9a86"),
        ),
        yaxis=dict(
            gridcolor="rgba(0,212,170,0.08)",
            zerolinecolor="rgba(0,212,170,0.15)",
            tickfont=dict(color="#5a9a86"),
        ),
        margin=dict(t=20, b=10, l=10, r=10),
        **kwargs,
    )
    return fig


# ── Main ──────────────────────────────────────────────────────────────────────

def _inject_logo():
    import base64
    logo_path = os.path.join(_APP_DIR, "download.png")
    if not os.path.exists(logo_path):
        return
    with open(logo_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    st.sidebar.markdown(f"""
        <div style="
            background: rgba(255,255,255,0.07);
            border: 1px solid rgba(255,255,255,0.13);
            backdrop-filter: blur(8px);
            -webkit-backdrop-filter: blur(8px);
            border-radius: 12px;
            padding: 10px 14px;
            margin-bottom: 16px;
            text-align: center;
        ">
            <img src="data:image/png;base64,{b64}"
                 style="width: 100%; max-width: 190px; height: auto; display: block; margin: auto;
                        mix-blend-mode: luminosity; opacity: 0.92;" />
        </div>
    """, unsafe_allow_html=True)


def _inject_day_css():
    st.markdown("""
<style>
/* ── DAY MODE OVERRIDES ───────────────────────────────────────── */
.stApp {
    background: linear-gradient(135deg, #f0faf7 0%, #ffffff 60%, #f5f9ff 100%) !important;
}
header[data-testid="stHeader"] {
    background: linear-gradient(180deg, #e8f7f2 0%, #f0faf7 100%) !important;
    border-bottom: 1px solid rgba(0,180,140,0.18) !important;
}
[data-testid="stToolbar"] { background: transparent !important; }
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #e0f5ee 0%, #edf9f5 60%, #f0faf7 100%) !important;
    border-right: 1px solid rgba(0,180,140,0.2) !important;
}
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] small,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stMarkdown p { color: #1a4a3a !important; }
[data-testid="stSidebarContent"] .stRadio label[data-baseweb="radio"] {
    color: #3a7a62 !important;
    border-color: rgba(0,180,140,0.12) !important;
}
[data-testid="stSidebarContent"] .stRadio label[data-baseweb="radio"]:hover {
    background: rgba(0,180,140,0.1) !important;
    color: #1a4a3a !important;
}
[data-testid="stSidebarContent"] .stRadio label[data-baseweb="radio"]:has(input:checked) {
    background: linear-gradient(135deg,rgba(0,180,140,0.18) 0%,rgba(0,140,110,0.09) 100%) !important;
    color: #007a5a !important;
    border-left: 3px solid #00b894 !important;
}
h2 { color: #1a4a3a !important; }
h3 { color: #2d6b55 !important; }
[data-testid="metric-container"] {
    background: linear-gradient(145deg,#ffffff 0%,#f0faf7 100%) !important;
    border: 1px solid rgba(0,180,140,0.22) !important;
    box-shadow: 0 4px 18px rgba(0,180,140,0.1), inset 0 1px 0 rgba(255,255,255,0.8) !important;
}
[data-testid="metric-container"]:hover {
    box-shadow: 0 10px 28px rgba(0,180,140,0.18) !important;
}
[data-testid="stMetricLabel"] > div { color: #3a8a70 !important; }
.stSelectbox > div > div,
.stMultiSelect > div > div {
    background: #ffffff !important;
    border: 1px solid rgba(0,180,140,0.3) !important;
}
div[data-baseweb="select"] span,
div[data-baseweb="select"] div[class],
.stSelectbox [data-baseweb="select"] > div > div { color: #1a2e26 !important; }
ul[data-baseweb="menu"] li { background: #f0faf7 !important; color: #1a2e26 !important; }
ul[data-baseweb="menu"] li:hover { background: rgba(0,180,140,0.14) !important; }
.stTextInput > div > div > input,
.stNumberInput > div > div > input {
    background: #ffffff !important;
    border: 1px solid rgba(0,180,140,0.3) !important;
    color: #1a2e26 !important;
}
.streamlit-expanderHeader {
    background: rgba(0,180,140,0.07) !important;
    border: 1px solid rgba(0,180,140,0.2) !important;
    color: #2d6b55 !important;
}
.streamlit-expanderContent {
    background: #f6fcf9 !important;
    border-color: rgba(0,180,140,0.14) !important;
}
[data-testid="stForm"] {
    background: rgba(0,180,140,0.04) !important;
    border: 1px solid rgba(0,180,140,0.2) !important;
}
.stTabs [data-baseweb="tab-list"] {
    background: rgba(0,180,140,0.07) !important;
    border: 1px solid rgba(0,180,140,0.18) !important;
}
.stTabs [data-baseweb="tab"] { color: #3a8a70 !important; }
.block-container .stRadio > div[role="radiogroup"] {
    background: rgba(0,180,140,0.07) !important;
    border-color: rgba(0,180,140,0.18) !important;
}
.block-container .stRadio label[data-baseweb="radio"] { color: #3a8a70 !important; }
.block-container .stRadio label[data-baseweb="radio"]:hover { color: #1a4a3a !important; }
hr { background: linear-gradient(90deg,transparent,rgba(0,180,140,0.35),transparent) !important; }
.stSuccess > div { background: rgba(0,180,140,0.08) !important; color: #1a4a3a !important; }
.stInfo    > div { background: rgba(100,80,160,0.07) !important; color: #3a2d6b !important; }
.stWarning > div { background: rgba(160,200,0,0.07) !important;  color: #4a5a00 !important; }
.stError   > div { background: rgba(220,50,50,0.07) !important;  color: #6a1a1a !important; }
.stCaption, small { color: #5aaa8a !important; }
.dash-kpi-card {
    background: linear-gradient(150deg,#ffffff 0%,#f0faf7 100%) !important;
    border: 1px solid rgba(0,180,140,0.18) !important;
    box-shadow: 0 4px 16px rgba(0,180,140,0.08) !important;
}
.dash-kpi-label  { color: #5aaa8a !important; }
.dash-kpi-sub    { color: #5aaa8a !important; }
.dash-section-hdr { border-bottom-color: rgba(0,180,140,0.15) !important; }
.dash-section-title { color: #2d6b55 !important; }
.dash-section-sub   { color: #5aaa8a !important; }
[data-testid="stDataFrame"] > div,
[data-testid="stDataEditor"] > div {
    border: 1px solid rgba(0,180,140,0.18) !important;
    background: #ffffff !important;
}
::-webkit-scrollbar-track { background: #f0faf7; }
</style>
""", unsafe_allow_html=True)


def _page_login():
    st.markdown("""
<style>
[data-testid="stAppViewContainer"] {
    background: radial-gradient(ellipse at 20% 50%, #0a2a20 0%, #060f0d 60%, #000000 100%) !important;
}
[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stSidebar"]       { display: none !important; }
[data-testid="collapsedControl"]{ display: none !important; }

/* Remove Streamlit's default block padding so card hugs content */
[data-testid="stMainBlockContainer"] {
    padding-top: 0 !important;
    padding-bottom: 0 !important;
}
section[data-testid="stMain"] > div { padding-top: 0 !important; }

/* Floating particles */
@keyframes float1 {
  0%,100%{ transform:translateY(0) translateX(0);   opacity:.5; }
  50%    { transform:translateY(-35px) translateX(12px); opacity:.15; }
}
@keyframes float2 {
  0%,100%{ transform:translateY(0) translateX(0);    opacity:.35; }
  50%    { transform:translateY(25px) translateX(-18px); opacity:.1; }
}
@keyframes float3 {
  0%,100%{ transform:translateY(0) translateX(0);   opacity:.25; }
  50%    { transform:translateY(-18px) translateX(8px); opacity:.5; }
}
@keyframes glow-ring {
  0%,100%{ box-shadow:0 0 0 0 rgba(0,180,140,.5),0 0 16px rgba(0,180,140,.2); }
  50%    { box-shadow:0 0 0 9px rgba(0,180,140,0), 0 0 32px rgba(0,180,140,.35); }
}
@keyframes slideUp {
  from{ opacity:0; transform:translateY(28px); }
  to  { opacity:1; transform:translateY(0); }
}
@keyframes shimmer {
  0%  { background-position:-200% center; }
  100%{ background-position: 200% center; }
}

.login-particles { position:fixed;top:0;left:0;width:100%;height:100%;pointer-events:none;z-index:0; }
.particle        { position:absolute;border-radius:50%;background:rgba(0,180,140,.13); }
.p1{ width:160px;height:160px;top:8%; left:4%;  animation:float1 8s  ease-in-out infinite; }
.p2{ width:90px; height:90px; top:58%;left:78%; animation:float2 6s  ease-in-out infinite; }
.p3{ width:55px; height:55px; top:78%;left:14%; animation:float3 7s  ease-in-out infinite; }
.p4{ width:38px; height:38px; top:22%;left:68%; animation:float1 9s  ease-in-out infinite reverse; }
.p5{ width:110px;height:110px;top:42%;left:88%; animation:float2 10s ease-in-out infinite; }

/* Outer centering — full viewport */
.login-page {
  display:flex; justify-content:center; align-items:center;
  height:100vh; position:relative; z-index:1;
}

/* The single card that wraps EVERYTHING including the form */
.login-card {
  width:380px;
  background:linear-gradient(160deg,rgba(13,45,36,.96) 0%,rgba(6,18,14,.98) 100%);
  border:1px solid rgba(0,180,140,.28);
  border-radius:22px;
  padding:32px 36px 28px;
  box-shadow:0 20px 70px rgba(0,0,0,.55),0 0 50px rgba(0,180,140,.06),inset 0 1px 0 rgba(0,180,140,.12);
  animation:slideUp .6s cubic-bezier(.22,1,.36,1) both;
  backdrop-filter:blur(18px);
}
.login-ring {
  width:60px;height:60px;margin:0 auto 14px;border-radius:50%;
  border:2px solid rgba(0,180,140,.6);
  display:flex;align-items:center;justify-content:center;font-size:1.8rem;
  animation:glow-ring 2.5s ease-in-out infinite;
  background:radial-gradient(circle,rgba(0,180,140,.14) 0%,transparent 70%);
}
.login-brand {
  background:linear-gradient(90deg,#00b48c,#00e0b0,#00b48c);
  background-size:200% auto;
  -webkit-background-clip:text;-webkit-text-fill-color:transparent;
  animation:shimmer 3s linear infinite;
  font-size:1.65rem;font-weight:800;text-align:center;
  margin:0 0 4px;letter-spacing:1.5px;
}
.login-sub {
  color:#5aaa8a;font-size:0.72rem;text-align:center;
  margin:0 0 18px;letter-spacing:.5px;text-transform:uppercase;
}
.login-sep { border:none;border-top:1px solid rgba(0,180,140,.12);margin:0 0 18px; }

/* Style the Streamlit form that sits below the card header */
div[data-testid="stForm"] {
  background:transparent !important;border:none !important;padding:0 !important;
}
div[data-testid="stForm"] input {
  background:rgba(0,180,140,.06) !important;
  border:1px solid rgba(0,180,140,.2) !important;
  border-radius:9px !important;color:#d0f0e8 !important;font-size:.9rem !important;
}
div[data-testid="stForm"] input:focus {
  border-color:rgba(0,180,140,.6) !important;
  box-shadow:0 0 0 2px rgba(0,180,140,.14) !important;
}
div[data-testid="stForm"] label {
  color:#7ac8a8 !important;font-size:.7rem !important;
  letter-spacing:.8px !important;text-transform:uppercase !important;font-weight:600 !important;
}
div[data-testid="stForm"] button[kind="primaryFormSubmit"] {
  background:linear-gradient(90deg,#00b48c,#00c89e) !important;
  border:none !important;border-radius:9px !important;
  font-size:.95rem !important;font-weight:700 !important;
  height:44px !important;margin-top:6px !important;
}
div[data-testid="stForm"] button[kind="primaryFormSubmit"]:hover {
  background:linear-gradient(90deg,#00c89e,#00e0b0) !important;
  transform:translateY(-1px) !important;
  box-shadow:0 7px 20px rgba(0,180,140,.32) !important;
}
/* Tighten vertical gaps inside the form */
div[data-testid="stForm"] .stTextInput { margin-bottom:4px !important; }
</style>

<div class="login-particles">
  <div class="particle p1"></div><div class="particle p2"></div>
  <div class="particle p3"></div><div class="particle p4"></div>
  <div class="particle p5"></div>
</div>
<div class="login-page">
  <div class="login-card">
    <div class="login-ring">🔐</div>
    <p class="login-brand">TalentRupt</p>
    <p class="login-sub">Billing &amp; Cost Analysis Platform</p>
    <hr class="login-sep"/>
  </div>
</div>
""", unsafe_allow_html=True)

    # Narrow centered column so inputs aren't full-width
    _, col, _ = st.columns([2, 1.4, 2])
    with col:
        with st.form("login_form"):
            username = st.text_input("Username", placeholder="Enter username")
            password = st.text_input("Password", type="password", placeholder="Enter password")
            submitted = st.form_submit_button("Sign In →", type="primary", use_container_width=True)

    if submitted:
        u = username.strip()
        p = password.strip()
        if u == _ADMIN_USER.strip() and p == _ADMIN_PASS.strip():
            st.session_state["role"] = "admin"
            st.session_state["username"] = u
            st.rerun()
        elif u == _SALES_USER.strip() and p == _SALES_PASS.strip():
            st.session_state["role"] = "sales"
            st.session_state["username"] = u
            st.rerun()
        else:
            st.error("Invalid username or password.")


def main():
    if "role" not in st.session_state:
        _page_login()
        return

    role = st.session_state["role"]

    _inject_css()
    if st.session_state.get("day_mode", False):
        _inject_day_css()
    _inject_logo()
    st.sidebar.title("Navigation")

    col_r, col_theme = st.sidebar.columns([1, 1])
    if col_r.button("🔄 Refresh", use_container_width=True):
        st.session_state["cache_bust"] = st.session_state.get("cache_bust", 0) + 1
        st.rerun()
    theme_label = "☀️ Day" if not st.session_state.get("day_mode", False) else "🌙 Night"
    if col_theme.button(theme_label, key="theme_toggle", use_container_width=True):
        st.session_state["day_mode"] = not st.session_state.get("day_mode", False)
        st.rerun()
    st.sidebar.checkbox("₹ INR", key="show_inr")

    # ── FX Rate management ────────────────────────────────────────────────────
    with st.sidebar.expander("💱 Wire Rates ($ → ₹)"):
        fx_rates = load_fx_rates()
        for m, r in sorted(fx_rates.items()):
            st.caption(f"{m}: $1 = ₹{r:.2f}")
        st.divider()
        with st.form("fx_form", clear_on_submit=True):
            fx_month = st.text_input("Month", placeholder="Feb 2026")
            fx_rate  = st.number_input("₹ per $", min_value=1.0, value=91.0, step=0.01, format="%.2f")
            if st.form_submit_button("Save Rate"):
                if fx_month:
                    fx_rates[fx_month] = fx_rate
                    save_fx_rates(fx_rates)
                    st.success(f"Saved ₹{fx_rate:.2f}/$ for {fx_month}")
                    st.rerun()
    all_pages = {
        "Dashboard":         page_dashboard,
        "Billing Clients":   page_billing_clients,
        "Employee Details":  page_employee_details,
        "Expenses":          page_expenses,
        "Month Comparison":  page_compare,
        "Detailed Records":  page_detail,
        "Estimate":          page_estimate,
        "Sales Dept Data":   _page_sales_dept_data,
        "Add Data":          page_data_entry,
        "Import Excel":      page_import,
    }
    icons = {
        "Dashboard":         "📊",
        "Billing Clients":   "🏢",
        "Employee Details":  "👥",
        "Expenses":          "🧾",
        "Month Comparison":  "📈",
        "Detailed Records":  "🔍",
        "Estimate":          "🔮",
        "Sales Dept Data":   "💼",
        "Add Data":          "➕",
        "Import Excel":      "📂",
    }

    if role == "sales":
        pages = {"Sales Dept Data": all_pages["Sales Dept Data"]}
    else:
        pages = all_pages

    choice = st.sidebar.radio(
        "Go to",
        list(pages.keys()),
        format_func=lambda k: f"{icons[k]} {k}",
        key="page_choice",
    )

    st.sidebar.divider()
    st.sidebar.caption(f"Logged in as **{st.session_state.get('username', role)}**")
    if st.sidebar.button("🚪 Logout", use_container_width=True):
        st.session_state.clear()
        st.rerun()

    data = get_all_data()

    if role == "admin":
        if data:
            st.sidebar.divider()
            st.sidebar.subheader("Summary")
            for m, df in data.items():
                bal = df["Balance"].sum()
                color = "🟢" if bal >= 0 else "🔴"
                st.sidebar.write(f"{color} **{m}**: {_disp(bal, m)}")

    needs_data = ("Dashboard", "Estimate", "Billing Clients", "Employee Details",
                  "Expenses", "Month Comparison", "Detailed Records", "Sales Dept Data")
    if choice in needs_data:
        pages[choice](data)
    else:
        pages[choice]()


if __name__ == "__main__":
    main()
