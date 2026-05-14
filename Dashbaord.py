import os
import shutil
import tempfile
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

os.chdir(os.path.dirname(os.path.abspath(__file__)))


def safe_read_excel(path, **kwargs):
    """
    Windows-safe Excel reader.
    If the file is locked by Excel (PermissionError), copies it to a
    temp file first and reads from there — dashboard works even when
    the source file is open in Excel.
    """
    try:
        return pd.read_excel(path, **kwargs)
    except PermissionError:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            shutil.copy2(path, tmp.name)
            tmp_path = tmp.name
        result = pd.read_excel(tmp_path, **kwargs)
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
        return result


def safe_excel_file(path):
    """
    Windows-safe ExcelFile opener — same lock-bypass logic.
    Returns a (pandas.ExcelFile, sheet_names) tuple.
    """
    try:
        return pd.ExcelFile(path)
    except PermissionError:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            shutil.copy2(path, tmp.name)
            return pd.ExcelFile(tmp.name)

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ICAI Counsellor Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────────────────────────
# CUSTOM CSS  — sidebar labels dark, selected pills readable
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f0f4f8; }
    .block-container { padding-top: 1.5rem; padding-bottom: 1rem; }

    /* KPI Cards */
    [data-testid="metric-container"] {
        background: white;
        border-radius: 14px;
        padding: 18px 20px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.07);
        border-left: 5px solid #4f8ef7;
    }
    [data-testid="metric-container"] > div { color: #1e293b !important; }
    [data-testid="stMetricValue"] { font-size: 28px !important; font-weight: 700 !important; }
    [data-testid="stMetricLabel"] { font-size: 13px !important; color: #64748b !important; }

    /* Title */
    .dash-title {
        background: linear-gradient(135deg, #1e3a8a, #3b82f6);
        color: white;
        padding: 20px 28px;
        border-radius: 16px;
        margin-bottom: 20px;
    }
    .dash-title h1 { color: white !important; margin: 0; font-size: 26px; }
    .dash-title p  { color: #bfdbfe !important; margin: 4px 0 0; font-size: 14px; }

    /* Section headers */
    .section-header {
        background: white;
        border-radius: 10px;
        padding: 10px 18px;
        margin: 16px 0 8px;
        border-left: 5px solid #3b82f6;
        font-weight: 600;
        font-size: 16px;
        color: #1e293b;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    }

    /* Insight box */
    .insight-box {
        background: white;
        border-radius: 12px;
        padding: 16px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        margin-bottom: 12px;
        border: 1px solid #e2e8f0;
    }
    .stDataFrame { border-radius: 12px !important; }

    /* ── SIDEBAR ── dark background, ALL text white/light */
    div[data-testid="stSidebarContent"] {
        background: #1e293b;
    }

    /* labels above widgets */
    div[data-testid="stSidebarContent"] label,
    div[data-testid="stSidebarContent"] .stSelectbox label,
    div[data-testid="stSidebarContent"] .stMultiSelect label,
    div[data-testid="stSidebarContent"] p,
    div[data-testid="stSidebarContent"] span {
        color: #f1f5f9 !important;
        font-size: 13px;
        font-weight: 600;
    }

    /* dropdown / select box text */
    div[data-testid="stSidebarContent"] .stSelectbox div[data-baseweb="select"] *,
    div[data-testid="stSidebarContent"] .stMultiSelect div[data-baseweb="select"] * {
        color: #1e293b !important;
        background-color: #f8fafc !important;
    }

    /* selected pill tags — force dark text so they are readable */
    div[data-testid="stSidebarContent"] span[data-baseweb="tag"] {
        background-color: #1d4ed8 !important;
        color: #ffffff !important;
    }
    div[data-testid="stSidebarContent"] span[data-baseweb="tag"] span {
        color: #ffffff !important;
    }

    /* selectbox current value */
    div[data-testid="stSidebarContent"] div[data-baseweb="select"] > div {
        background-color: #f8fafc !important;
        color: #1e293b !important;
    }

    /* sidebar markdown headers */
    div[data-testid="stSidebarContent"] h2,
    div[data-testid="stSidebarContent"] h3 {
        color: #93c5fd !important;
    }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────
# SESSION STATE — cross-chart filter selection
# ─────────────────────────────────────────────────────────────────
if "selected_counsellor" not in st.session_state:
    st.session_state["selected_counsellor"] = None
if "selected_branch" not in st.session_state:
    st.session_state["selected_branch"] = None


# ─────────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────────
INPUT_FILE          = "Input File.xlsx"
MAPPED_BRANCHES_FILE = "Mapped Branches.xlsx"


@st.cache_data
def load_input_data():
    xl = safe_excel_file(INPUT_FILE)
    frames = []
    for sheet in xl.sheet_names:
        df = safe_read_excel(INPUT_FILE, sheet_name=sheet)
        df.columns = df.columns.str.strip()
        df["Year"] = str(sheet)
        frames.append(df)
    raw = pd.concat(frames, ignore_index=True)

    rename_map = {
        "Councellor ID No": "Counsellor_ID",
        "Councellor ID No.": "Counsellor_ID",
        "Counsellor ID": "Counsellor_ID",
        "Counselor ID": "Counsellor_ID",
        "Branch Name": "Branch_Name",
        "Branch": "Branch_Name",
        "Total Students Attended": "Students_Attended",
        "Students Attended": "Students_Attended",
        "No of Participants": "Students_Attended",
        "Participants": "Students_Attended",
        "Event Date": "Event_Date",
        "Start Time": "Start_Time",
        "End Time": "End_Time",
        "Estimated Grant": "Estimated_Grant",
        "Grant": "Estimated_Grant",
        "Reward": "Estimated_Grant",
        "State": "State",
        "City": "City",
        "Region": "State",
    }
    raw.rename(columns={k: v for k, v in rename_map.items() if k in raw.columns}, inplace=True)

    for col in ["Students_Attended", "Estimated_Grant", "State", "City"]:
        if col not in raw.columns:
            raw[col] = 0 if col in ["Students_Attended", "Estimated_Grant"] else "Unknown"

    for col in ["Students_Attended", "Estimated_Grant"]:
        raw[col] = pd.to_numeric(raw[col], errors="coerce").fillna(0)

    if "Event_Date" in raw.columns:
        raw["Event_Date"] = pd.to_datetime(raw["Event_Date"], errors="coerce")
        raw["Month"] = raw["Event_Date"].dt.month_name()

    raw["Lecture_Duration_Minutes"] = 0
    if "Start_Time" in raw.columns and "End_Time" in raw.columns:
        st_col = pd.to_datetime(raw["Start_Time"], errors="coerce")
        et_col = pd.to_datetime(raw["End_Time"], errors="coerce")
        raw["Lecture_Duration_Minutes"] = ((et_col - st_col).dt.total_seconds() / 60).fillna(0)
        raw["Lecture_Duration_Minutes"] = raw["Lecture_Duration_Minutes"].clip(lower=0)

    # FIX: Remove .0 from Counsellor_ID — keep as integer string
    if "Counsellor_ID" in raw.columns:
        raw["Counsellor_ID"] = pd.to_numeric(raw["Counsellor_ID"], errors="coerce")
        raw["Counsellor_ID"] = raw["Counsellor_ID"].dropna().astype(int).reindex(raw.index)
        raw["Counsellor_ID"] = raw["Counsellor_ID"].fillna(0).astype(int)

    return raw


@st.cache_data
def load_mapped_branches():
    """
    Mapped Branches.xlsx — single sheet, wide format:
        State | City | 2024 | 2025 | Branch Name

    A single branch (e.g. "Kalyan Dombivli") can appear across multiple
    city rows — we SUM registrations per (Branch_Name, Year) after melting.

    Returns long-format df: Branch_Name | Year | Registrations
    """
    df = safe_read_excel(MAPPED_BRANCHES_FILE, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]

    # Rename branch column
    for cand in ["Branch Name", "Branch"]:
        if cand in df.columns:
            df.rename(columns={cand: "Branch_Name"}, inplace=True)
            break

    df = df.dropna(subset=["Branch_Name"])

    # Detect year columns: 4-digit numeric strings e.g. "2024", "2025"
    year_cols = [c for c in df.columns
                 if str(c).strip().isdigit() and len(str(c).strip()) == 4]

    if year_cols:
        melted = df[["Branch_Name"] + year_cols].melt(
            id_vars="Branch_Name",
            var_name="Year",
            value_name="Registrations"
        )
        melted["Registrations"] = pd.to_numeric(melted["Registrations"], errors="coerce").fillna(0)
        # SUM across cities so each (Branch_Name, Year) is unique
        mb = melted.groupby(["Branch_Name", "Year"], as_index=False)["Registrations"].sum()
    else:
        # Fallback: sum all numeric columns per branch
        num_cols = df.select_dtypes(include="number").columns.tolist()
        df["Registrations"] = df[num_cols].sum(axis=1) if num_cols else 0
        df["Year"] = "All"
        mb = df.groupby(["Branch_Name", "Year"], as_index=False)["Registrations"].sum()

    mb["Year"] = mb["Year"].astype(str)   # match raw["Year"] which is also str
    return mb.reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────
# MAIN LOAD
# ─────────────────────────────────────────────────────────────────
try:
    raw = load_input_data()
    mb  = load_mapped_branches()
except FileNotFoundError as e:
    st.error(f"""
    ⚠️ **File Not Found**: {e}

    Please make sure both files exist in the **same folder as this script**:
    - `Input File.xlsx`
    - `Mapped Branches.xlsx`
    """)
    st.stop()


# ─────────────────────────────────────────────────────────────────
# BRANCH REGISTRATIONS  (full, all years — used before year filter)
# mb is long-format: Branch_Name | Year | Registrations
# ─────────────────────────────────────────────────────────────────
# Keep mb_long for year-aware lookups later
mb_long = mb.copy()   # Branch_Name | Year | Registrations


# ─────────────────────────────────────────────────────────────────
# EFFICIENCY SCORE  (merge registrations year-by-year into raw)
# ─────────────────────────────────────────────────────────────────
def normalize(s):
    mx = s.max()
    return s / mx if mx > 0 else s * 0

# Merge on (Branch_Name + Year) so each session row gets its correct year's regs
raw_with_reg = raw.merge(mb_long, on=["Branch_Name", "Year"], how="left")
raw_with_reg["Registrations"] = raw_with_reg["Registrations"].fillna(0)

raw_with_reg["Participation_Rate"] = np.where(
    raw_with_reg["Registrations"] > 0,
    (raw_with_reg["Students_Attended"] / raw_with_reg["Registrations"]) * 100,
    0
).clip(0, 100)

raw_with_reg["Cost_Per_Student"] = np.where(
    raw_with_reg["Students_Attended"] > 0,
    raw_with_reg["Estimated_Grant"] / raw_with_reg["Students_Attended"],
    0
)

raw_with_reg["Efficiency_Score"] = (
    normalize(raw_with_reg["Participation_Rate"]) * 0.35 +
    normalize(raw_with_reg["Students_Attended"])  * 0.25 +
    normalize(raw_with_reg["Lecture_Duration_Minutes"]) * 0.15 +
    (1 - normalize(raw_with_reg["Cost_Per_Student"])) * 0.15 +
    normalize(raw_with_reg["Registrations"]) * 0.10
) * 100
raw_with_reg["Efficiency_Score"] = raw_with_reg["Efficiency_Score"].fillna(0).clip(0, 100)


# ─────────────────────────────────────────────────────────────────
# SIDEBAR FILTERS
# FIX: Removed State & City. Removed "Registrations" option from Year.
# ─────────────────────────────────────────────────────────────────
st.sidebar.markdown("## 🔎 Dashboard Filters")
st.sidebar.markdown("---")

# YEAR — only actual year values, no "Registrations"
years_raw = sorted(raw_with_reg["Year"].dropna().unique())
# Filter out any non-numeric-looking year values like "Registrations"
years = [y for y in years_raw if str(y).strip().isdigit() or (len(str(y)) == 4 and str(y)[:4].isdigit())]
if not years:
    years = years_raw  # fallback if none match
sel_years = st.sidebar.multiselect("📅 Year", years, default=years)

filt = raw_with_reg[raw_with_reg["Year"].isin(sel_years)]

# ── Year-aware registrations: sum only selected years from mb_long ──
branch_reg = (
    mb_long[mb_long["Year"].isin(sel_years)]
    .groupby("Branch_Name")["Registrations"].sum()
    .reset_index()
)

# BRANCH filter (replaces State + City)
branches_all = sorted(filt["Branch_Name"].dropna().unique())
sel_branch = st.sidebar.selectbox("🏢 Branch", ["All"] + list(branches_all))
if sel_branch != "All":
    filt = filt[filt["Branch_Name"] == sel_branch]

# COUNSELLOR filter
counsellors = sorted(filt["Counsellor_ID"].dropna().unique())
sel_counsellors = st.sidebar.multiselect("👤 Counsellor", counsellors, default=counsellors)
filt = filt[filt["Counsellor_ID"].isin(sel_counsellors)]

st.sidebar.markdown("---")
top_n = st.sidebar.selectbox("🏆 Top N Counsellors", [5, 10, 20, 50], index=0)

rank_by_labels = {
    "Sessions Conducted": "Sessions_Conducted",
    "No of Student Participated": "Students_Attended",
    "Timing of Lecture / Duration": "Avg_Duration_Min",
}

rank_by_label = st.sidebar.selectbox(
    "📊 Rank Counsellors By",
    list(rank_by_labels.keys())
)

rank_by = rank_by_labels[rank_by_label]

# ── Cross-filter clear button always visible in sidebar ──
st.sidebar.markdown("---")
_qp_check = st.query_params
_cid_active    = _qp_check.get("cid", None)
_branch_active = _qp_check.get("branch", None)
if _cid_active or _branch_active:
    if _cid_active:
        st.sidebar.warning(f"📌 Chart filter: Counsellor **{_cid_active}**")
    if _branch_active:
        st.sidebar.warning(f"📌 Chart filter: Branch **{_branch_active}**")
    if st.sidebar.button("❌ Clear Chart Filter", use_container_width=True):
        st.query_params.clear()
        st.rerun()
else:
    st.sidebar.info("💡 Click any chart bar to cross-filter the dashboard.")


# ─────────────────────────────────────────────────────────────────
# COUNSELLOR SUMMARY
# ─────────────────────────────────────────────────────────────────
summary = filt.groupby("Counsellor_ID").agg(
    Sessions_Conducted   = ("Counsellor_ID", "count"),
    Students_Attended    = ("Students_Attended", "sum"),
    Estimated_Grant      = ("Estimated_Grant", "sum"),
    Avg_Participation_Rt = ("Participation_Rate", "mean"),
    Efficiency_Score     = ("Efficiency_Score", "mean"),
    Avg_Duration_Min     = ("Lecture_Duration_Minutes", "mean"),
    Branch_Name          = ("Branch_Name", lambda x: x.mode()[0] if len(x) > 0 else "Unknown"),
    State                = ("State", lambda x: x.mode()[0] if len(x) > 0 else "Unknown"),
).reset_index()

summary = summary.merge(branch_reg, on="Branch_Name", how="left")
summary["Registrations"] = summary["Registrations"].fillna(0)

summary["Cost_Per_Student"] = np.where(
    summary["Students_Attended"] > 0,
    summary["Estimated_Grant"] / summary["Students_Attended"], 0
)
summary["Participation_Rate"] = np.where(
    summary["Registrations"] > 0,
    (summary["Students_Attended"] / summary["Registrations"]) * 100, 0
).clip(0, 100)

def categorize(row, df):
    eff = row["Efficiency_Score"]
    p80 = df["Efficiency_Score"].quantile(0.80)
    p50 = df["Efficiency_Score"].quantile(0.50)
    p25 = df["Efficiency_Score"].quantile(0.25)
    if eff >= p80:   return "⭐ Best"
    elif eff >= p50: return "✅ Good"
    elif eff >= p25: return "🔶 Can Do Better"
    else:            return "⚠️ Needs Attention"

summary["Performance"] = summary.apply(lambda r: categorize(r, summary), axis=1)
summary = summary.sort_values(by=rank_by, ascending=False)

# FIX: Ensure Counsellor_ID is integer (no .0)
summary["Counsellor_ID"] = summary["Counsellor_ID"].astype(int)


# ─────────────────────────────────────────────────────────────────
# BRANCH SUMMARY
# ─────────────────────────────────────────────────────────────────
branch_sess = filt.groupby("Branch_Name").agg(
    Sessions          = ("Counsellor_ID", "count"),
    Students_Attended = ("Students_Attended", "sum"),
    Avg_Efficiency    = ("Efficiency_Score", "mean"),
    Estimated_Grant   = ("Estimated_Grant", "sum"),
).reset_index()

branch_combined = branch_sess.merge(branch_reg, on="Branch_Name", how="left")
branch_combined["Registrations"] = branch_combined["Registrations"].fillna(0)
branch_combined["Participation_Rate"] = np.where(
    branch_combined["Registrations"] > 0,
    (branch_combined["Students_Attended"] / branch_combined["Registrations"]) * 100, 0
).clip(0, 100)


# ─────────────────────────────────────────────────────────────────
# INTERACTIVE FILTER via click (Power BI style)
# ─────────────────────────────────────────────────────────────────
# Streamlit query params used to simulate click-filter
qp = st.query_params
clicked_cid   = qp.get("cid", None)
clicked_bname = qp.get("branch", None)

# Apply click-level cross-filter to views
filt_summary = summary.copy()
filt_branch  = branch_combined.copy()

if clicked_bname and clicked_bname != "__all__":
    filt_summary = filt_summary[filt_summary["Branch_Name"] == clicked_bname]

if clicked_cid and clicked_cid != "__all__":
    try:
        cid_int = int(clicked_cid)
        filt_summary = filt_summary[filt_summary["Counsellor_ID"] == cid_int]
        # filter branch chart to that counsellor's branch
        cid_branch = summary[summary["Counsellor_ID"] == cid_int]["Branch_Name"].values
        if len(cid_branch) > 0:
            filt_branch = filt_branch[filt_branch["Branch_Name"].isin(cid_branch)]
    except ValueError:
        pass

# Active chart filter banner on main page (shown when cross-filter is on)
if (clicked_cid and clicked_cid != "__all__") or (clicked_bname and clicked_bname != "__all__"):
    label = []
    if clicked_cid and clicked_cid != "__all__":
        label.append(f"Counsellor {clicked_cid}")
    if clicked_bname and clicked_bname != "__all__":
        label.append(f"Branch: {clicked_bname}")
    st.info(f"📌 Chart filter active: **{' | '.join(label)}** — use the **❌ Clear Chart Filter** button in the sidebar to reset.", icon="🔍")


# ─────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────
st.markdown("""
<div class="dash-title">
    <h1>📊 ICAI Counsellor Performance Dashboard</h1>
    <p>Owner View · Sessions · Participation · Registrations · Efficiency · Rewards</p>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────
# KPI ROW
# FIX: "Branch Registrations" → "Students Registered"
# ─────────────────────────────────────────────────────────────────
total_sessions    = int(filt.shape[0])
total_counsellors = filt["Counsellor_ID"].nunique()
total_attended    = int(filt["Students_Attended"].sum())
# FIX: Use mb_long filtered by selected years — branch_combined only contains
# branches that had sessions, missing branches with registrations but no sessions.
total_reg         = int(mb_long[mb_long["Year"].isin(sel_years)]["Registrations"].sum())
total_reward      = int(filt["Estimated_Grant"].sum())
avg_efficiency    = round(summary["Efficiency_Score"].mean(), 1) if len(summary) else 0
overall_part_rate = round((total_attended / total_reg * 100) if total_reg > 0 else 0, 1)

c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
c1.metric("📋 Total Sessions",       f"{total_sessions:,}")
c2.metric("👤 Active Counsellors",   f"{total_counsellors:,}")
c3.metric("🎓 Students Attended",    f"{total_attended:,}")
c4.metric("📝 Students Registered",  f"{total_reg:,}")   # FIX: was "Branch Registrations"
c5.metric("📈 Participation Rate",   f"{overall_part_rate}%")
c6.metric("💰 Total Rewards (₹)",    f"₹{total_reward:,}")
c7.metric("⚡ Avg Efficiency",        f"{avg_efficiency}")

st.markdown("<br>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────
# OWNER INSIGHTS BANNER
# ─────────────────────────────────────────────────────────────────
if len(summary) > 0:
    best      = summary.iloc[0]
    worst     = summary[summary["Performance"] == "⚠️ Needs Attention"]
    high_cost = summary.sort_values("Cost_Per_Student", ascending=False).iloc[0]

    col_ins1, col_ins2, col_ins3 = st.columns(3)
    with col_ins1:
        st.markdown(f"""
        <div class="insight-box">
            <b>🏆 Top Performer</b><br>
            <span style="font-size:20px; font-weight:700; color:#1d4ed8">{best['Counsellor_ID']}</span><br>
            Sessions: <b>{int(best['Sessions_Conducted'])}</b> &nbsp;|&nbsp;
            Attended: <b>{int(best['Students_Attended'])}</b><br>
            Efficiency: <b>{round(best['Efficiency_Score'],1)}</b>
        </div>
        """, unsafe_allow_html=True)
    with col_ins2:
        st.markdown(f"""
        <div class="insight-box">
            <b>⚠️ Needs Attention</b><br>
            <span style="font-size:20px; font-weight:700; color:#dc2626">{len(worst)} Counsellors</span><br>
            Low efficiency & participation.<br>
            Review & support recommended.
        </div>
        """, unsafe_allow_html=True)
    with col_ins3:
        st.markdown(f"""
        <div class="insight-box">
            <b>💸 Highest Cost Per Student</b><br>
            <span style="font-size:20px; font-weight:700; color:#d97706">{high_cost['Counsellor_ID']}</span><br>
            ₹{round(high_cost['Cost_Per_Student'],0):,.0f} per student<br>
            ROI may need review.
        </div>
        """, unsafe_allow_html=True)

st.markdown("---")


# ─────────────────────────────────────────────────────────────────
# SECTION 1 — TOP COUNSELLORS — GROUPED BAR CHART
# FIX: Always show Sessions + Participants + Metric side-by-side per counsellor
# ─────────────────────────────────────────────────────────────────
st.markdown(
    f'<div class="section-header">🏆 Top {top_n} Counsellors — Ranked by {rank_by_label}</div>',
    unsafe_allow_html=True
)
st.caption(f"Bars show Sessions Conducted, Students Attended, and {rank_by_label} per counsellor. Click a bar to cross-filter.")

top_df = filt_summary.head(top_n).copy()
top_df["Counsellor_ID"] = top_df["Counsellor_ID"].astype(str)

# Build a grouped bar with 3 metrics always visible
COLOR_SESSIONS  = "#3b82f6"   # blue
COLOR_ATTENDED  = "#10b981"   # green
COLOR_METRIC    = "#f59e0b"   # amber  (rank_by metric)

fig_top = go.Figure()

fig_top.add_trace(go.Bar(
    name="Sessions Conducted",
    x=top_df["Counsellor_ID"],
    y=top_df["Sessions_Conducted"],
    marker_color=COLOR_SESSIONS,
    text=top_df["Sessions_Conducted"],
    textposition="outside",
    hovertemplate="<b>Counsellor %{x}</b><br>Sessions: %{y}<extra></extra>",
))

fig_top.add_trace(go.Bar(
    name="Students Attended",
    x=top_df["Counsellor_ID"],
    y=top_df["Students_Attended"],
    marker_color=COLOR_ATTENDED,
    text=top_df["Students_Attended"],
    textposition="outside",
    hovertemplate="<b>Counsellor %{x}</b><br>Students Attended: %{y}<extra></extra>",
))

# Third bar: the chosen rank_by metric (only if different from the above two)
if rank_by not in ("Sessions_Conducted", "Students_Attended"):
    fig_top.add_trace(go.Bar(
        name=rank_by_label,
        x=top_df["Counsellor_ID"],
        y=top_df[rank_by],
        marker_color=COLOR_METRIC,
        text=top_df[rank_by].round(1),
        textposition="outside",
        hovertemplate=f"<b>Counsellor %{{x}}</b><br>{rank_by_label}: %{{y:.1f}}<extra></extra>",
    ))

fig_top.update_layout(
    barmode="group",
    height=460,
    showlegend=True,
    legend=dict(orientation="h", y=1.08),
    xaxis_title="Counsellor ID",
    yaxis_title="Count / Value",
    plot_bgcolor="white",
    paper_bgcolor="white",
    xaxis=dict(tickangle=-30, type="category"),
)

# Plotly click event via on_select (Streamlit ≥1.33)
event_top = st.plotly_chart(fig_top, use_container_width=True, on_select="rerun", key="top_chart")
if event_top and event_top.selection and event_top.selection.get("points"):
    clicked_x = event_top.selection["points"][0].get("x")
    if clicked_x:
        st.query_params["cid"] = str(clicked_x)
        st.rerun()

st.markdown("---")


# ─────────────────────────────────────────────────────────────────
# SECTION 2 — COUNSELLOR TIME EFFICIENCY ANALYSIS
# ─────────────────────────────────────────────────────────────────
# SECTION — COUNSELLOR TIME EFFICIENCY ANALYSIS
# ─────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="section-header">⏱️ Counsellor Time Efficiency Analysis</div>',
    unsafe_allow_html=True
)

st.caption(
    "Compare counsellors achieving high attendance in shorter lecture duration versus counsellors taking longer sessions with lower attendance."
)

eff_df = filt_summary.copy()

# Remove invalid rows
eff_df = eff_df[
    (eff_df["Avg_Duration_Min"] > 0) &
    (eff_df["Students_Attended"] > 0)
].copy()

# Time efficiency score
eff_df["Time_Efficiency"] = (
    eff_df["Students_Attended"] /
    eff_df["Avg_Duration_Min"]
)

# Efficient counsellors
best_eff = (
    eff_df.sort_values("Time_Efficiency", ascending=False)
    .head(10)
    .sort_values("Students_Attended")
)

# Inefficient counsellors
worst_eff = (
    eff_df.sort_values("Time_Efficiency", ascending=True)
    .head(10)
    .sort_values("Students_Attended")
)

# ───────────────── INSIGHT CARDS ─────────────────
good_count = len(best_eff)
bad_count = len(worst_eff)

c1, c2 = st.columns(2)

with c1:
    st.markdown(f"""
    <div style="
        background:#dcfce7;
        padding:18px;
        border-radius:12px;
        border-left:6px solid #22c55e;
        margin-bottom:10px;
    ">
        <h4 style="color:#166534; margin:0;">
            ✅ Efficient Counsellors
        </h4>
        <p style="margin-top:8px; color:#166534;">
            {good_count} counsellors achieved higher attendance
            with lower lecture duration.
        </p>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown(f"""
    <div style="
        background:#fee2e2;
        padding:18px;
        border-radius:12px;
        border-left:6px solid #ef4444;
        margin-bottom:10px;
    ">
        <h4 style="color:#991b1b; margin:0;">
            ⚠️ Needs Attention
        </h4>
        <p style="margin-top:8px; color:#991b1b;">
            {bad_count} counsellors are taking longer sessions
            with lower attendance.
        </p>
    </div>
    """, unsafe_allow_html=True)

# ───────────────── SIDE BY SIDE CHARTS ─────────────────
col1, col2 = st.columns(2)

# ───────────────── LEFT CHART ─────────────────
with col1:

    st.markdown("### ✅ High Attendance with Lower Lecture Duration")

    fig_best = go.Figure()

    fig_best.add_trace(go.Bar(
        x=best_eff["Students_Attended"],
        y=best_eff["Counsellor_ID"].astype(str).tolist(),
        orientation="h",
        marker_color="#22c55e",
        width=0.85,
        text=[
            f"{int(a)} Students | {int(t)} Min"
            for a, t in zip(
                best_eff["Students_Attended"],
                best_eff["Avg_Duration_Min"]
            )
        ],
        textposition="outside",
        hovertemplate=(
            "<b>Counsellor %{y}</b><br>"
            "Students Attended: %{x}<br>"
            "Lecture Duration: %{text}<extra></extra>"
        )
    ))

    fig_best.update_layout(
        height=650,
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis_title="Students Attended",
        yaxis_title="Counsellor ID",
        bargap=0.15,
        yaxis=dict(type="category"),
        margin=dict(l=40, r=180, t=40, b=40),
    )

    fig_best.update_traces(
        cliponaxis=False
    )

    st.plotly_chart(fig_best, use_container_width=True)

# ───────────────── RIGHT CHART ─────────────────
with col2:

    st.markdown("### ⚠️ Longer Sessions with Lower Attendance")

    fig_worst = go.Figure()

    fig_worst.add_trace(go.Bar(
        x=worst_eff["Students_Attended"],
        y=worst_eff["Counsellor_ID"].astype(str).tolist(),
        orientation="h",
        marker_color="#ef4444",
        width=0.85,
        text=[
            f"{int(a)} Students | {int(t)} Min"
            for a, t in zip(
                worst_eff["Students_Attended"],
                worst_eff["Avg_Duration_Min"]
            )
        ],
        textposition="outside",
        hovertemplate=(
            "<b>Counsellor %{y}</b><br>"
            "Students Attended: %{x}<br>"
            "Lecture Duration: %{text}<extra></extra>"
        )
    ))

    fig_worst.update_layout(
        height=650,
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis_title="Students Attended",
        yaxis_title="Counsellor ID",
        bargap=0.15,
        yaxis=dict(type="category"),
        margin=dict(l=40, r=180, t=40, b=40),
    )

    fig_worst.update_traces(
        cliponaxis=False
    )

    st.plotly_chart(fig_worst, use_container_width=True)

st.markdown("---")



# ─────────────────────────────────────────────────────────────────
# SECTION 3 — BRANCH EFFICIENCY — HORIZONTAL BAR (replaces scatter)
# ─────────────────────────────────────────────────────────────────
# SECTION — COUNSELLOR FREQUENT SESSION MONTH ANALYSIS
# ─────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="section-header">📅 Counsellor Frequent Session Month Analysis</div>',
    unsafe_allow_html=True
)

st.caption(
    "Identify counsellors conducting unusually high number of sessions within the same month."
)

# Prepare data
freq_df = filt.copy()

# Convert event date
freq_df["Event_Date"] = pd.to_datetime(
    freq_df["Event_Date"],
    errors="coerce"
)

# Extract month
freq_df["Frequent_Month"] = (
    freq_df["Event_Date"]
    .dt.strftime("%b-%Y")
)

# Session count per counsellor per month
monthly_sessions = (
    freq_df.groupby(
        ["Counsellor_ID", "Frequent_Month"]
    )
    .size()
    .reset_index(name="Session_Count")
)

# Keep highest frequency months
monthly_sessions = (
    monthly_sessions
    .sort_values("Session_Count", ascending=False)
    .head(25)
)

# Add performance category
monthly_sessions["Frequency_Level"] = np.where(
    monthly_sessions["Session_Count"] >= monthly_sessions["Session_Count"].quantile(0.75),
    "⚠️ Highly Concentrated",
    "✅ Balanced"
)

# Rename columns
monthly_sessions.rename(columns={
    "Counsellor_ID": "Counsellor ID",
    "Frequent_Month": "Frequent Month",
    "Session_Count": "No. of Sessions"
}, inplace=True)

# Reset index for Sr No
monthly_sessions.reset_index(drop=True, inplace=True)
monthly_sessions.index += 1

# Style function
def highlight_frequency(val):
    if "Highly" in str(val):
        return "background-color:#fee2e2; color:#991b1b; font-weight:600"
    return "background-color:#dcfce7; color:#166534; font-weight:600"

styled_matrix = (
    monthly_sessions.style
    .map(
        highlight_frequency,
        subset=["Frequency_Level"]
    )
)

st.dataframe(
    styled_matrix,
    use_container_width=True,
    height=600
)

# ───────────────── INSIGHT CARDS ─────────────────
peak_row = monthly_sessions.iloc[0]

c1, c2 = st.columns(2)

with c1:
    st.markdown(f"""
    <div style="
        background:#fee2e2;
        padding:18px;
        border-radius:12px;
        border-left:6px solid #ef4444;
    ">
        <h4 style="margin:0; color:#991b1b;">
            ⚠️ Highest Session Concentration
        </h4>
        <p style="margin-top:8px; color:#991b1b;">
            Counsellor ID <b>{peak_row['Counsellor ID']}</b>
            conducted <b>{peak_row['No. of Sessions']}</b>
            sessions in <b>{peak_row['Frequent Month']}</b>.
        </p>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown("""
    <div style="
        background:#dbeafe;
        padding:18px;
        border-radius:12px;
        border-left:6px solid #3b82f6;
    ">
        <h4 style="margin:0; color:#1e40af;">
            📊 Operational Insight
        </h4>
        <p style="margin-top:8px; color:#1e40af;">
            Repeatedly high session counts in a single month
            may indicate target rushing or uneven workload distribution.
        </p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("---")


# ─────────────────────────────────────────────────────────────────
# SECTION 4 — ROI ALERT: GROUPED BAR (Reward vs Participation)
# FIX: Replaces old single-metric bar; shows both Reward & Participation side-by-side
# ─────────────────────────────────────────────────────────────────
st.markdown('<div class="section-header">⚠️ Reward vs Participation — ROI Alert</div>', unsafe_allow_html=True)
st.caption(
    "Counsellors with high reward but low participation are flagged. "
    "Blue = Reward (₹), Orange = Students Attended. Large gap = ROI concern."
)

poor_roi = filt_summary.sort_values(
    by=["Estimated_Grant", "Students_Attended"],
    ascending=[False, True]
).head(top_n).copy()
poor_roi["Counsellor_ID"] = poor_roi["Counsellor_ID"].astype(str)

fig_roi = go.Figure()
fig_roi.add_trace(go.Bar(
    name="Reward (₹)",
    x=poor_roi["Counsellor_ID"],
    y=poor_roi["Estimated_Grant"],
    marker_color="#3b82f6",
    text=["₹{:,}".format(int(v)) for v in poor_roi["Estimated_Grant"]],
    textposition="outside",
    textfont=dict(color="#1e40af", size=12, family="Arial Black"),
    hovertemplate="<b>Counsellor %{x}</b><br>Reward: ₹%{y:,}<extra></extra>",
    yaxis="y1",
))
fig_roi.add_trace(go.Bar(
    name="Students Attended",
    x=poor_roi["Counsellor_ID"],
    y=poor_roi["Students_Attended"],
    marker_color="#f97316",
    text=poor_roi["Students_Attended"],
    textposition="outside",
    textfont=dict(color="#c2410c", size=12, family="Arial Black"),
    hovertemplate="<b>Counsellor %{x}</b><br>Students Attended: %{y:,}<extra></extra>",
    yaxis="y2",
))

fig_roi.update_layout(
    barmode="group",
    height=460,
    xaxis=dict(title="Counsellor ID", tickangle=-30, type="category"),
    yaxis=dict(
        title=dict(text="Reward (₹)", font=dict(color="#3b82f6")),
        tickfont=dict(color="#3b82f6"),
    ),
    yaxis2=dict(
        title=dict(text="Students Attended", font=dict(color="#f97316")),
        tickfont=dict(color="#f97316"),
        overlaying="y",
        side="right",
    ),
    legend=dict(orientation="h", y=1.08),
    plot_bgcolor="white",
    paper_bgcolor="white",
)

event_roi = st.plotly_chart(fig_roi, use_container_width=True, on_select="rerun", key="roi_chart")
if event_roi and event_roi.selection and event_roi.selection.get("points"):
    clicked_x2 = event_roi.selection["points"][0].get("x")
    if clicked_x2:
        st.query_params["cid"] = str(clicked_x2)
        st.rerun()

st.markdown("---")



# ─────────────────────────────────────────────────────────────────
# SECTION 6 — EFFICIENCY DISTRIBUTION
# ─────────────────────────────────────────────────────────────────
# SECTION — MOST EFFICIENT BRANCHES
# ─────────────────────────────────────────────────────────────────
st.markdown(
    '<div class="section-header">🏆 Most Efficient Branches</div>',
    unsafe_allow_html=True
)

st.caption(
    "Branches achieving strong student attendance with better participation and optimized lecture duration."
)

# Create branch duration data
branch_duration = (
    filt.groupby("Branch_Name")["Lecture_Duration_Minutes"]
    .mean()
    .reset_index()
    .rename(columns={
        "Lecture_Duration_Minutes": "Avg_Duration_Min"
    })
)

# Merge with branch summary
branch_eff = filt_branch.merge(
    branch_duration,
    on="Branch_Name",
    how="left"
)

# Remove invalid rows
branch_eff = branch_eff[
    (branch_eff["Students_Attended"] > 0) &
    (branch_eff["Participation_Rate"] > 0) &
    (branch_eff["Avg_Duration_Min"] > 0)
].copy()

# Efficiency score
branch_eff["Efficiency_Score"] = (
    (
        branch_eff["Students_Attended"] *
        branch_eff["Participation_Rate"]
    ) / branch_eff["Avg_Duration_Min"]
)

# Top branches
top_branches = (
    branch_eff.sort_values(
        "Efficiency_Score",
        ascending=False
    )
    .head(6)
)

# Cards layout
cols = st.columns(3)

for idx, (_, row) in enumerate(top_branches.iterrows()):

    with cols[idx % 3]:

        st.markdown(f"""
        <div style="
            background:white;
            padding:20px;
            border-radius:16px;
            box-shadow:0 2px 10px rgba(0,0,0,0.08);
            border-left:6px solid #22c55e;
            margin-bottom:20px;
        ">

        <h4 style="
            margin:0;
            color:#166534;
            font-size:20px;
        ">
            🏆 {row['Branch_Name']}
        </h4>

        <hr style="margin:12px 0;">

        <p style="margin:6px 0; font-size:15px;">
            👥 <b>Students Attended:</b>
            {int(row['Students_Attended']):,}
        </p>

        <p style="margin:6px 0; font-size:15px;">
            📈 <b>Participation:</b>
            {row['Participation_Rate']:.1f}%
        </p>

        <p style="margin:6px 0; font-size:15px;">
            ⏱️ <b>Avg Duration:</b>
            {row['Avg_Duration_Min']:.0f} Min
        </p>

        <p style="margin:6px 0; font-size:15px;">
            🎯 <b>Efficiency Score:</b>
            {row['Efficiency_Score']:.1f}
        </p>

        </div>
        """, unsafe_allow_html=True)

st.markdown("---")
# ─────────────────────────────────────────────────────────────────
# SECTION 7 — COUNSELLOR PERFORMANCE TABLE
# ─────────────────────────────────────────────────────────────────
st.markdown('<div class="section-header">📋 Counsellor Performance Table</div>', unsafe_allow_html=True)

display_df = filt_summary.copy()

# FIX: Counsellor_ID as int
display_df["Counsellor_ID"] = display_df["Counsellor_ID"].astype(int)

# Avg Max Lecture Duration
max_duration = (
    filt.groupby("Counsellor_ID")["Lecture_Duration_Minutes"]
    .max()
    .reset_index()
    .rename(columns={"Lecture_Duration_Minutes": "Max_Duration_Min"})
)

display_df = display_df.merge(
    max_duration,
    on="Counsellor_ID",
    how="left"
)

# Keep only required columns
display_df = display_df[[
    "Counsellor_ID",
    "Branch_Name",
    "Performance",
    "Sessions_Conducted",
    "Students_Attended",
    "Participation_Rate",
    "Estimated_Grant",
    "Cost_Per_Student",
    "Avg_Duration_Min",
    "Max_Duration_Min"
]]

# Add Sr No
display_df.insert(0, "Sr No", range(1, len(display_df) + 1))

# Round values
display_df["Participation_Rate"] = display_df["Participation_Rate"].round(1)
display_df["Avg_Duration_Min"] = display_df["Avg_Duration_Min"].round(0)
display_df["Max_Duration_Min"] = display_df["Max_Duration_Min"].round(0)

display_df["Estimated_Grant"] = display_df["Estimated_Grant"].round(0).astype(int)
display_df["Cost_Per_Student"] = display_df["Cost_Per_Student"].round(0).astype(int)

# Rename columns
display_df.rename(columns={
    "Counsellor_ID": "Counsellor ID",
    "Branch_Name": "Branch",
    "Sessions_Conducted": "Sessions Taken",
    "Students_Attended": "Attended / Participated Student No",
    "Participation_Rate": "Participants %",
    "Estimated_Grant": "Reward (₹)",
    "Cost_Per_Student": "Cost Per Student (₹)",
    "Avg_Duration_Min": "Avg Lec Timing (Min)",
    "Max_Duration_Min": "Avg Lec Timing (Max)",
}, inplace=True)

# Performance color
def color_performance(val):
    if "Best" in str(val):
        return "background-color:#dcfce7; color:#166534; font-weight:600"
    elif "Good" in str(val):
        return "background-color:#dbeafe; color:#1e40af; font-weight:600"
    elif "Can Do" in str(val):
        return "background-color:#fef9c3; color:#854d0e; font-weight:600"
    elif "Needs" in str(val):
        return "background-color:#fee2e2; color:#991b1b; font-weight:600"
    return ""

styled = display_df.style.map(
    color_performance,
    subset=["Performance"]
).format({
    "Participants %": "{:.1f}%",
    "Reward (₹)": "₹{:,}",
    "Cost Per Student (₹)": "₹{:,}",
    "Avg Lec Timing (Min)": "{:.0f}",
    "Avg Lec Timing (Max)": "{:.0f}",
})

st.dataframe(styled, use_container_width=True, height=500)
# ─────────────────────────────────────────────────────────────────
# DOWNLOAD
# ─────────────────────────────────────────────────────────────────
csv = display_df.to_csv(index=False).encode("utf-8")
st.download_button(
    "⬇️ Download Counsellor Report (CSV)",
    csv, "Counsellor_Performance_Report.csv", "text/csv",
    use_container_width=True,
)

st.markdown("---")
st.markdown(
    "<center style='color:#94a3b8; font-size:12px;'>ICAI Counsellor Dashboard · "
    "Sessions & Attendance from Input File · Registrations from Mapped Branches · "
    "Built for owner-level decision making</center>",
    unsafe_allow_html=True,
)