from __future__ import annotations

import json
import html
import tempfile
import sys
from pathlib import Path

APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components

from io_utils import save_uploads, list_xlsx_sheets, preview_file, preview_moda_core
from dq_matching import (
    profile_file,
    pairwise_matching,
    moda_file_overlap,
    unmatched_sites,
    scoped_source_site_matching,
)
from pipeline import run_analysis, available_periods
from analytics import rows_to_df, overview_metrics, cfm_journey, cfm_woreda_table, protection_summary
from outputs import detailed_tracker_xlsx, aggregated_tracker_xlsx, management_report_docx, dq_report_docx, bundle_zip
from mapping import build_site_points, apply_map_filters, apply_photo_filters, load_admin_assets, leaflet_html, ACTIVITY_LABELS

# -----------------------------------------------------------------------------
# PAGE / BRANDING
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="J-MAAP | Jijiga Monitoring, Analytics & Action Platform",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

WFP_BLUE = "#007DBC"
WFP_DARK = "#005B8E"
TEXT_DARK = "#16324F"
BG = "#F5F8FB"
CARD_BORDER = "#DCE6EF"

st.markdown(
    f"""
<style>
:root {{ --wfp-blue:{WFP_BLUE}; --wfp-dark:{WFP_DARK}; --text:{TEXT_DARK}; --bg:{BG}; --border:{CARD_BORDER}; }}
html, body, [class*="css"] {{ font-family: "Open Sans", "Segoe UI", sans-serif; color: var(--text); }}
.stApp {{ background: var(--bg); }}
.block-container {{ padding-top: 0.35rem; padding-bottom: 2rem; max-width: 100%; }}
[data-testid="stSidebar"] {{ background: #F1F6FA; border-right: 1px solid #D7E3EC; }}
[data-testid="stSidebar"] .block-container {{ padding-top: .65rem; }}
[data-testid="stSidebarNav"] {{ display:none; }}
[data-testid="stMetric"] {{ background:white; border:1px solid var(--border); padding:14px 16px; border-radius:12px; box-shadow:0 1px 2px rgba(23,49,79,.05); }}
[data-testid="stMetricLabel"] {{ color:#60758A; font-size:.82rem; }}
[data-testid="stMetricValue"] {{ color:#0B5790; font-size:1.65rem; font-weight:700; }}
[data-testid="stDataFrame"] {{ background:white; border:1px solid var(--border); border-radius:12px; overflow:hidden; }}
.stButton > button, .stDownloadButton > button {{ border-radius:9px; font-weight:600; }}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {{ background:var(--wfp-blue); border-color:var(--wfp-blue); }}
.wfp-topbar {{
  margin: -.35rem -1rem 1rem -1rem; padding: 15px 28px; color:white;
  background: linear-gradient(90deg, #006AA6 0%, #007DBC 56%, #0067A0 100%);
  display:flex; align-items:center; justify-content:space-between; gap:24px; min-height:74px;
  box-shadow:0 2px 5px rgba(0,66,105,.18);
}}
.wfp-brand {{ display:flex; align-items:center; gap:14px; min-width:255px; }}
.wfp-mark {{ width:48px; height:48px; border:2px solid rgba(255,255,255,.9); border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:20px; font-weight:800; }}
.wfp-brand-title {{ font-weight:700; font-size:17px; line-height:1.1; }}
.wfp-brand-sub {{ font-size:11px; opacity:.9; margin-top:3px; }}
.wfp-center {{ text-align:center; flex:1; }}
.wfp-center h1 {{ margin:0; color:white; font-size:27px; font-weight:750; letter-spacing:.1px; }}
.wfp-center p {{ margin:4px 0 0 0; font-size:12px; opacity:.95; }}
.wfp-period {{ min-width:200px; text-align:right; font-size:12px; }}
.page-title {{ font-size:29px; color:#123D67; font-weight:750; margin: 4px 0 2px 0; }}
.page-subtitle {{ color:#60758A; margin-bottom:14px; font-size:.94rem; }}
.section-title {{ color:#123D67; font-weight:700; font-size:1.08rem; margin:10px 0 8px 0; }}
.card {{ background:white; border:1px solid var(--border); border-radius:12px; padding:16px 18px; box-shadow:0 1px 2px rgba(23,49,79,.04); }}
.card-title {{ color:#123D67; font-size:.98rem; font-weight:700; margin-bottom:4px; }}
.muted {{ color:#6B7F92; font-size:.85rem; }}
.status-ok {{background:#ECFDF3;border:1px solid #ABEFC6;padding:9px 12px;border-radius:9px;color:#176B45;}}
.status-warn {{background:#FFFAEB;border:1px solid #FEDF89;padding:9px 12px;border-radius:9px;color:#8A5A00;}}
.status-info {{background:#EFF8FF;border:1px solid #B2DDFF;padding:9px 12px;border-radius:9px;color:#175C8D;}}
.step-card {{background:white;border:1px solid var(--border);border-radius:11px;padding:12px 14px;margin-bottom:8px;}}
.step-num {{display:inline-block;width:25px;height:25px;border-radius:50%;background:#EAF4FA;color:#006AA6;font-weight:700;text-align:center;line-height:25px;margin-right:8px;}}
.sidebar-brand {{ background:#006EA9; color:white; padding:14px 12px; border-radius:10px; margin-bottom:12px; }}
.sidebar-brand strong {{font-size:15px;}}
.sidebar-brand small {{display:block; margin-top:4px; opacity:.9;}}
hr {{ border-color:#DDE7EF; }}
</style>
""",
    unsafe_allow_html=True,
)


def branded_header(reporting_month: str) -> None:
    month_label = reporting_month
    try:
        month_label = pd.Period(reporting_month, freq="M").strftime("%B %Y")
    except Exception:
        pass
    st.markdown(
        f"""
<div class="wfp-topbar">
  <div class="wfp-brand">
    <div class="wfp-mark">WFP</div>
    <div><div class="wfp-brand-title">World Food Programme</div><div class="wfp-brand-sub">SAVING LIVES · CHANGING LIVES</div></div>
  </div>
  <div class="wfp-center"><h1>J-MAAP – Jijiga Monitoring, Analytics & Action Platform</h1><p>From Data to Action | Accountable Programmes | Stronger Communities</p></div>
  <div class="wfp-period">Reporting Month<br><strong>{month_label}</strong></div>
</div>
""",
        unsafe_allow_html=True,
    )


def page_heading(title: str, subtitle: str) -> None:
    st.markdown(f'<div class="page-title">{title}</div><div class="page-subtitle">{subtitle}</div>', unsafe_allow_html=True)


def fmt_pct(x):
    return f"{x:.1%}" if pd.notna(x) else ""


@st.cache_data(show_spinner=False)
def cached_raw_preview(path: str, sheet: str | None, source_type: str, sub_office: str | None = None):
    if source_type == "MoDa":
        return preview_moda_core(path, 1000, sub_office=sub_office)
    return preview_file(path, sheet, 1000)


@st.cache_data(show_spinner=False)
def cached_available_periods(moda_paths: tuple[str, ...]):
    return pd.DataFrame(available_periods(list(moda_paths)))


# -----------------------------------------------------------------------------
# STATE / SIDEBAR
# -----------------------------------------------------------------------------
if "workdir" not in st.session_state:
    st.session_state.workdir = tempfile.mkdtemp(prefix="jijiga_monitoring_")
if "analysis" not in st.session_state:
    st.session_state.analysis = None
if "upload_signature" not in st.session_state:
    st.session_state.upload_signature = None
if "generated_outputs" not in st.session_state:
    st.session_state.generated_outputs = None

# Build the global Sub-office filter from the currently uploaded MoDa source(s).
# Before any upload, show the three Somali Region operational offices as sensible defaults.
persisted_profiles = st.session_state.get("profiles", [])
persisted_moda_paths = tuple(p.get("path") for p in persisted_profiles if p.get("source_type") == "MoDa" and p.get("path"))
detected_suboffices = []
if persisted_moda_paths:
    try:
        _avail_sidebar = cached_available_periods(persisted_moda_paths)
        if not _avail_sidebar.empty:
            detected_suboffices = sorted(_avail_sidebar["Sub-office"].dropna().astype(str).str.strip().loc[lambda x: x.ne("")].unique().tolist())
    except Exception:
        detected_suboffices = []

if not detected_suboffices:
    detected_suboffices = ["Jijiga", "Gode", "Dollo Addo"]
else:
    # Keep Somali Region offices at the top while retaining all offices detected in an all-Ethiopia MoDa export.
    preferred = ["Jijiga", "Gode", "Dollo Addo"]
    ordered = [x for x in preferred if x in detected_suboffices]
    ordered.extend([x for x in detected_suboffices if x not in ordered])
    detected_suboffices = ordered

_previous_so = st.session_state.get("sub_office_filter", "Jijiga")
_default_so = _previous_so if _previous_so in detected_suboffices else ("Jijiga" if "Jijiga" in detected_suboffices else detected_suboffices[0])

with st.sidebar:
    st.markdown('<div class="sidebar-brand"><strong>Jijiga AO</strong><small>Monthly Monitoring Analytics</small></div>', unsafe_allow_html=True)
    reporting_month = st.text_input("Reporting month", value="2026-08", help="YYYY-MM; reporting period is derived from monitoring date.")
    sub_office = st.selectbox(
        "Sub-office",
        detected_suboffices,
        index=detected_suboffices.index(_default_so),
        key="sub_office_filter",
        help="Options are detected from the uploaded MoDa file(s). Jijiga, Gode and Dollo Addo are prioritised for Somali Region use.",
    )
    if persisted_moda_paths:
        st.caption(f"{len(detected_suboffices)} sub-office(s) detected in uploaded MoDa data")
    st.divider()
    nav = st.radio(
        "Navigation",
        [
            "Home",
            "1. Data Upload",
            "2. Data Quality & Matching",
            "3. Raw Data Explorer",
            "4. Monitoring Overview",
            "5. Visited Sites Map",
            "6. AAP / CFM",
            "7. Protection, Safety & Dignity",
            "8. Activity Analysis",
            "9. Findings & Actions",
            "10. Generate Outputs",
        ],
        index=0,
        label_visibility="collapsed",
    )
    st.divider()
    st.caption("Validation gates")
    st.markdown("1. System signal  \\n2. Field validation  \\n3. Internal WFP agreement  \\n4. Action assignment  \\n5. Verification & closure")
    st.divider()
    st.caption("Version 1.6 · Mapping-enabled deployment package")

# Changing the selected reporting cohort invalidates any previously calculated run.
current_params = (reporting_month.strip(), sub_office.strip())
if st.session_state.get("analysis") is not None and st.session_state.get("analysis_params") != current_params:
    st.session_state.analysis = None
    st.session_state.generated_outputs = None

branded_header(reporting_month)

# -----------------------------------------------------------------------------
# UPLOAD / PROFILING (persistent control at top)
# -----------------------------------------------------------------------------
uploads = st.file_uploader(
    "Upload monthly raw files",
    type=["xlsx", "xlsm", "csv"],
    accept_multiple_files=True,
    help="Upload one or more MoDa exports. RBMF, FRN, logistics/handover and previous action trackers are optional and will be profiled separately.",
    label_visibility="collapsed" if nav != "1. Data Upload" else "visible",
)

# Restore persisted upload/profile state on every Streamlit rerun.
# Page navigation and widget changes rerun the whole script; the file uploader
# can still contain files while the upload signature is unchanged.
profiles = st.session_state.get("profiles", [])
paths = st.session_state.get("paths", [])

if uploads:
    signature = tuple((f.name, getattr(f, "size", len(f.getbuffer()))) for f in uploads)
    if signature != st.session_state.upload_signature:
        st.session_state.workdir = tempfile.mkdtemp(prefix="jijiga_monitoring_")
        paths = save_uploads(uploads, st.session_state.workdir)
        profiles = []
        with st.spinner("Profiling uploaded files and building matching diagnostics..."):
            for p in paths:
                try:
                    profiles.append(profile_file(p))
                except Exception as e:
                    st.error(f"Could not profile {Path(p).name}: {e}")
            st.session_state.profiles = profiles
            st.session_state.paths = paths
            st.session_state.pairs = pairwise_matching(profiles)
            st.session_state.uuid_mat, st.session_state.site_mat = moda_file_overlap(profiles)
            st.session_state.unmatched = unmatched_sites(profiles)
        st.session_state.analysis = None
        st.session_state.analysis_params = None
        st.session_state.generated_outputs = None
        st.session_state.upload_signature = signature
        # Refresh once so the sidebar Sub-office dropdown immediately reflects the uploaded MoDa file(s).
        st.rerun()

    # Explicitly reload persisted values even when the signature did not change.
    profiles = st.session_state.get("profiles", [])
    paths = st.session_state.get("paths", [])

# -----------------------------------------------------------------------------
# HOME
# -----------------------------------------------------------------------------
if nav == "Home":
    page_heading("Monitoring Analytics Control Centre", "A transparent monthly workflow from raw submissions to validated management action.")
    if profiles:
        moda_count = sum(p["source_type"] == "MoDa" for p in profiles)
        a = st.session_state.analysis
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Files loaded", len(profiles))
        c2.metric("MoDa exports", moda_count)
        c3.metric("Sub-office", sub_office)
        c4.metric("Analysis status", "Ready" if a else "Analysis pending")
        c5.metric("Reporting month", reporting_month)
        if a:
            dq = a["dq"]
            st.markdown('<div class="status-ok">Monthly analysis is available. Review the DQ exceptions and field-validation status before generating final outputs.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-info">Files are loaded. Continue to Data Quality & Matching before running the monthly analysis.</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="status-info">Upload the monthly raw files to begin. MoDa exports drive the monitoring analysis; operational files are matched diagnostically.</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title">Monthly workflow</div>', unsafe_allow_html=True)
    cols = st.columns(4)
    steps = [
        ("1", "Upload", "Load raw MoDa and optional operational sources."),
        ("2", "Validate", "Inspect schema, duplicates, date logic and source matching."),
        ("3", "Analyse", "Calculate indicators, thematic signals and site-level findings."),
        ("4", "Act", "Validate findings, aggregate actions and generate management outputs."),
    ]
    for col, (n, t, d) in zip(cols, steps):
        with col:
            st.markdown(f'<div class="card"><span class="step-num">{n}</span><b>{t}</b><div class="muted" style="margin-top:8px">{d}</div></div>', unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# DATA UPLOAD
# -----------------------------------------------------------------------------
elif nav == "1. Data Upload":
    page_heading("1. Data Upload", "Load all monthly source files and confirm that the expected inputs have been recognised.")
    if not profiles:
        st.info("Upload one or more raw files using the uploader above.")
    else:
        inv = pd.DataFrame([{k: p.get(k) for k in ["file", "source_type", "sheet", "preview_rows", "columns", "site_column", "uuid_unique_preview"]} for p in profiles])
        moda_count = sum(p["source_type"] == "MoDa" for p in profiles)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Files uploaded", len(profiles))
        c2.metric("MoDa exports", moda_count)
        c3.metric("Other sources", len(profiles) - moda_count)
        c4.metric("Reporting month", reporting_month)
        st.dataframe(inv, use_container_width=True, hide_index=True)
        if moda_count:
            moda_paths_upload = tuple(p["path"] for p in profiles if p["source_type"] == "MoDa")
            try:
                _av = cached_available_periods(moda_paths_upload)
                if not _av.empty:
                    _so = sorted(_av["Sub-office"].dropna().astype(str).unique().tolist())
                    st.caption("Detected sub-offices: " + ", ".join(_so))
            except Exception:
                pass
            st.markdown('<div class="status-ok">MoDa source detected. Continue to Data Quality & Matching before running analysis.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="status-warn">No MoDa source detected. DQ profiling is available, but findings generation requires MoDa data.</div>', unsafe_allow_html=True)
        with st.expander("Column / schema inventory"):
            for p in profiles:
                st.markdown(f"**{p['file']} — {p['source_type']}**")
                st.write(p.get("column_names", []))

# -----------------------------------------------------------------------------
# DQ & MATCHING
# -----------------------------------------------------------------------------
elif nav == "2. Data Quality & Matching":
    page_heading("2. Data Quality & Matching", "Check source integrity, overlap and match quality before allowing the monthly analysis to proceed.")
    if not profiles:
        st.warning("Upload source files first.")
    else:
        moda_paths = [p["path"] for p in profiles if p["source_type"] == "MoDa"]
        availability = cached_available_periods(tuple(moda_paths)) if moda_paths else pd.DataFrame()
        eligible = 0
        available_months = []
        if not availability.empty:
            so_avail = availability[availability["Sub-office"].astype(str).str.casefold() == sub_office.strip().casefold()].copy()
            available_months = so_avail["Reporting month"].astype(str).tolist()
            match = so_avail[so_avail["Reporting month"] == reporting_month.strip()]
            eligible = int(match["Records"].sum()) if not match.empty else 0

            st.markdown('<div class="section-title">Analysis cohort availability</div>', unsafe_allow_html=True)
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Selected sub-office", sub_office)
            c2.metric("Selected reporting month", reporting_month)
            c3.metric("Eligible MoDa records", eligible)
            c4.metric("Detected months", len(available_months))
            st.dataframe(so_avail, use_container_width=True, hide_index=True)

            if eligible == 0:
                if available_months:
                    labels = ", ".join(available_months)
                    st.markdown(f'<div class="status-warn"><b>No eligible {sub_office} records were found for {reporting_month}.</b> Available actual-date periods in the uploaded MoDa file(s): {labels}. Change the Reporting month in the sidebar before running the analysis.</div>', unsafe_allow_html=True)
                else:
                    offices = ", ".join(sorted(availability["Sub-office"].astype(str).unique().tolist())[:20])
                    st.markdown(f'<div class="status-warn"><b>No records were found for sub-office “{sub_office}”.</b> Detected sub-offices include: {offices}.</div>', unsafe_allow_html=True)
            else:
                row = match.iloc[0]
                st.markdown(f'<div class="status-ok">Cohort is valid: <b>{eligible:,}</b> {sub_office} records dated {row["First date"]} to {row["Last date"]} are eligible for this run.</div>', unsafe_allow_html=True)
                try:
                    per = pd.Period(reporting_month.strip(), freq="M")
                    month_start = per.start_time.date().isoformat()
                    month_end = per.end_time.date().isoformat()
                    if str(row["First date"]) > month_start or str(row["Last date"]) < month_end:
                        st.markdown(f'<div class="status-warn"><b>Possible partial-month source:</b> the selected cohort covers {row["First date"]} to {row["Last date"]}, while the calendar month is {month_start} to {month_end}. If this is a completed reporting month, upload the additional MoDa export/form version covering the missing dates. J-MAAP will combine and UUID-deduplicate multiple exports.</div>', unsafe_allow_html=True)
                except Exception:
                    pass

        pairs = st.session_state.get("pairs", pd.DataFrame())
        if not pairs.empty:
            st.markdown('<div class="section-title">Cross-file reconciliation</div>', unsafe_allow_html=True)
            st.dataframe(pairs, use_container_width=True, hide_index=True)
        uuid_mat, site_mat = st.session_state.get("uuid_mat", pd.DataFrame()), st.session_state.get("site_mat", pd.DataFrame())
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="section-title">MoDa UUID overlap</div>', unsafe_allow_html=True)
            st.dataframe(uuid_mat, use_container_width=True) if not uuid_mat.empty else st.caption("No comparable MoDa UUID matrix available.")
        with c2:
            st.markdown('<div class="section-title">Normalized site overlap</div>', unsafe_allow_html=True)
            st.dataframe(site_mat, use_container_width=True) if not site_mat.empty else st.caption("No comparable site matrix available.")

        um = st.session_state.get("unmatched", pd.DataFrame())
        if not um.empty:
            st.markdown('<div class="section-title">Optional-source site matching against MoDa</div>', unsafe_allow_html=True)
            view = um.copy()
            if "Match %" in view.columns:
                view["Match %"] = view["Match %"].map(fmt_pct)
            st.dataframe(view, use_container_width=True, hide_index=True)
        st.caption("Exact site-name matching is intentionally conservative. Production reconciliation should use the canonical Site Crosswalk for spelling variants and corporate IDs.")

        can_run = bool(moda_paths) and (availability.empty or eligible > 0)
        if not can_run and moda_paths:
            st.button("Run controlled monthly analysis", type="primary", use_container_width=True, disabled=True, help="Select a reporting month/sub-office with eligible records first.")
        elif moda_paths and st.button("Run controlled monthly analysis", type="primary", use_container_width=True):
            with st.spinner("Running controlled analysis and DQ rules..."):
                st.session_state.analysis = run_analysis(moda_paths, reporting_month.strip(), sub_office.strip())
                st.session_state.analysis_params = (reporting_month.strip(), sub_office.strip())
                st.session_state.generated_outputs = None
            st.success("Analysis completed. Review DQ exceptions below before relying on findings.")

        if st.session_state.analysis:
            dq = st.session_state.analysis["dq"]
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Accepted submissions", dq["rows"])
            c2.metric("Unique UUIDs", dq["uuid_unique"])
            c3.metric("Duplicates removed", dq["duplicate_uuid_rows"])
            c4.metric("DQ issues", len(dq["issues"]))
            st.markdown('<div class="section-title">Automated DQ exceptions</div>', unsafe_allow_html=True)
            st.dataframe(pd.DataFrame(dq["issues"]), use_container_width=True, hide_index=True)
            scoped = scoped_source_site_matching(profiles, st.session_state.analysis["rows"])
            if not scoped.empty:
                scoped_view = scoped.copy()
                scoped_view["Match %"] = scoped_view["Match %"].map(fmt_pct)
                st.markdown('<div class="section-title">Matching to selected month / sub-office cohort</div>', unsafe_allow_html=True)
                st.dataframe(scoped_view, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# RAW EXPLORER
# -----------------------------------------------------------------------------
elif nav == "3. Raw Data Explorer":
    page_heading("3. Raw Data Explorer", "Explore uploaded records before analysis. Filters and previews are designed for transparency and spot-checking.")
    if not profiles:
        st.warning("Upload source files first.")
    else:
        a = st.session_state.analysis
        if a:
            rows = a["rows"]
            m = overview_metrics(rows)
            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("Total submissions", m["submissions"])
            c2.metric("Sites monitored", m["sites"])
            c3.metric("Woredas", m["woredas"])
            c4.metric("Activities", m["activities"])
            df_a = rows_to_df(rows)
            c5.metric("WFP submissions", int((df_a.get("provider", pd.Series(dtype=str)).astype(str).str.upper() == "WFP").sum()))
            c6.metric("TPM submissions", int((df_a.get("provider", pd.Series(dtype=str)).astype(str).str.upper() == "TPM").sum()))

        chosen = st.selectbox("Source file", [p["file"] for p in profiles])
        p = next(x for x in profiles if x["file"] == chosen)
        sheets = list_xlsx_sheets(p["path"])
        sheet = st.selectbox("Worksheet", sheets, index=(sheets.index("data") if "data" in sheets else 0)) if sheets else None
        dfraw = cached_raw_preview(p["path"], sheet, p["source_type"], sub_office=sub_office).copy()

        q = st.text_input("Search displayed columns", value="", placeholder="Type a value, site, woreda, activity or UUID fragment...")
        if q and not dfraw.empty:
            mask = dfraw.astype(str).apply(lambda s: s.str.contains(q, case=False, na=False)).any(axis=1)
            dfraw = dfraw[mask]

        if p["source_type"] == "MoDa":
            st.caption(f"Preview filtered to Sub-office: {sub_office}. For performance, the preview exposes core raw fields only; the complete workbook remains the source used by the analysis engine.")

        if not dfraw.empty:
            # Lightweight visual preview from displayed core fields when recognizable.
            date_col = next((c for c in dfraw.columns if str(c).lower() in {"monitoring_date", "date", "start"}), None)
            activity_col = next((c for c in dfraw.columns if "activity" in str(c).lower()), None)
            woreda_col = next((c for c in dfraw.columns if "woreda" in str(c).lower()), None)
            chart_cols = st.columns(3)
            if date_col:
                tmp = pd.to_datetime(dfraw[date_col], errors="coerce").dt.date.value_counts().sort_index().reset_index()
                tmp.columns = ["Date", "Records"]
                with chart_cols[0]: st.plotly_chart(px.bar(tmp, x="Date", y="Records", title="Records by date"), use_container_width=True, config={"displayModeBar":False})
            if activity_col:
                tmp = dfraw[activity_col].astype(str).value_counts().head(10).reset_index(); tmp.columns=["Activity","Records"]
                with chart_cols[1]: st.plotly_chart(px.pie(tmp, values="Records", names="Activity", hole=.5, title="Activity composition"), use_container_width=True, config={"displayModeBar":False})
            if woreda_col:
                tmp = dfraw[woreda_col].astype(str).value_counts().head(10).reset_index(); tmp.columns=["Woreda","Records"]
                with chart_cols[2]: st.plotly_chart(px.bar(tmp.sort_values("Records"), x="Records", y="Woreda", orientation="h", title="Top woredas"), use_container_width=True, config={"displayModeBar":False})

        st.markdown('<div class="section-title">Raw data preview</div>', unsafe_allow_html=True)
        st.dataframe(dfraw, use_container_width=True, height=540, hide_index=True)
        st.download_button("Download current preview as CSV", dfraw.to_csv(index=False).encode("utf-8-sig"), file_name=f"{Path(chosen).stem}_preview.csv", mime="text/csv")

# -----------------------------------------------------------------------------
# MONITORING OVERVIEW
# -----------------------------------------------------------------------------
elif nav == "4. Monitoring Overview":
    page_heading("4. Monitoring Overview", "Summarise the volume, geography and provider composition of the accepted monthly monitoring cohort.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis from Data Quality & Matching.")
    else:
        rows = a["rows"]; df = rows_to_df(rows); m = overview_metrics(rows)
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Submissions", m["submissions"]); c2.metric("Sites", m["sites"]); c3.metric("Woredas", m["woredas"]); c4.metric("Activities", m["activities"]); c5.metric("Providers", m["providers"])
        c1, c2 = st.columns(2)
        with c1:
            act = df["activity"].value_counts().reset_index(); act.columns=["Activity","Submissions"]
            st.plotly_chart(px.bar(act.sort_values("Submissions"), x="Submissions", y="Activity", orientation="h", title="Submissions by activity"), use_container_width=True)
        with c2:
            pr = df["provider"].value_counts().reset_index(); pr.columns=["Provider","Submissions"]
            st.plotly_chart(px.pie(pr, values="Submissions", names="Provider", hole=.55, title="WFP / TPM evidence composition"), use_container_width=True)
        if "monitoring_date" in df.columns:
            day = df.dropna(subset=["monitoring_date"]).groupby(df["monitoring_date"].dt.date).size().reset_index(name="Submissions"); day.columns=["Date","Submissions"]
            st.plotly_chart(px.line(day, x="Date", y="Submissions", markers=True, title="Monitoring submissions by date"), use_container_width=True)
        woreda = df["woreda"].replace("", pd.NA).value_counts().head(20).reset_index(); woreda.columns=["Woreda","Submissions"]
        st.plotly_chart(px.bar(woreda.sort_values("Submissions"), x="Submissions", y="Woreda", orientation="h", title="Submissions by woreda — top 20"), use_container_width=True)

# -----------------------------------------------------------------------------
# VISITED SITES MAP
# -----------------------------------------------------------------------------
elif nav == "5. Visited Sites Map":
    page_heading("5. Visited Sites Map", "Explore the monthly monitoring footprint and its photo evidence using street, satellite imagery, or the Somali Region administrative hollow map.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first. The map uses the accepted monthly cohort, its MoDa GPS coordinates and available attachment links.")
    else:
        rows = a["rows"]
        photos = a.get("photos", [])
        if not rows:
            st.info("No accepted monitoring records are available for the selected sub-office and reporting month.")
        else:
            activity_options = sorted({ACTIVITY_LABELS.get(str(r.get("activity") or "").strip(), str(r.get("activity") or "").strip()) for r in rows if str(r.get("activity") or "").strip()})
            zone_options = sorted({str(r.get("zone") or "").strip() for r in rows if str(r.get("zone") or "").strip()})
            provider_options = sorted({str(r.get("provider") or "").strip() for r in rows if str(r.get("provider") or "").strip()})

            f1, f2, f3, f4 = st.columns([1.1, 1, 1, 1])
            with f1:
                selected_activities = st.multiselect("Activity", activity_options, default=[])
            with f2:
                selected_zones = st.multiselect("Zone", zone_options, default=[])
            candidate_rows = apply_map_filters(rows, activities=selected_activities, zones=selected_zones)
            woreda_options = sorted({str(r.get("woreda") or "").strip() for r in candidate_rows if str(r.get("woreda") or "").strip()})
            with f3:
                selected_woredas = st.multiselect("Woreda", woreda_options, default=[])
            with f4:
                selected_providers = st.multiselect("Provider", provider_options, default=[])

            filtered_rows = apply_map_filters(rows, activities=selected_activities, zones=selected_zones, woredas=selected_woredas, providers=selected_providers)
            filtered_photos = apply_photo_filters(photos, activities=selected_activities, zones=selected_zones, woredas=selected_woredas, providers=selected_providers)
            points = build_site_points(filtered_rows, a.get("detail", []), filtered_photos)
            valid_submission_count = int(points["submissions"].sum()) if not points.empty else 0
            total_filtered = len(filtered_rows)
            missing_gps = max(total_filtered - valid_submission_count, 0)
            photo_site_count = len({(p.get("site"), p.get("woreda"), p.get("zone")) for p in filtered_photos if p.get("site")})

            c1, c2, c3, c4, c5, c6 = st.columns(6)
            c1.metric("Mapped sites", len(points))
            c2.metric("Monitoring events", int(points["events"].sum()) if not points.empty else 0)
            c3.metric("Mapped submissions", valid_submission_count)
            c4.metric("Woredas visited", int(points["woreda"].replace("", pd.NA).nunique()) if not points.empty else 0)
            c5.metric("Photo evidence", len(filtered_photos), help=f"Attachment links detected across {photo_site_count} mapped/monitored sites.")
            c6.metric("Records without valid GPS", missing_gps)

            ctl1, ctl2, ctl3, ctl4 = st.columns([1.15, 1.15, 1, 1])
            with ctl1:
                basemap = st.radio("Map view", ["Street", "Satellite imagery", "Administrative hollow"], horizontal=True)
            with ctl2:
                display_mode = st.selectbox("Point display", ["Activity", "Visit frequency", "Provider", "Finding severity"])
            with ctl3:
                show_woreda_labels = st.checkbox("Show woreda labels", value=False, help="Zone labels are shown by default. Woreda labels can be crowded at regional scale.")
            with ctl4:
                show_photo_thumbnails = st.checkbox("Photo thumbnails in popup", value=True, help="When accessible, show up to four recent MoDa photos when a site is clicked.")

            if points.empty:
                st.warning("No valid MoDa GPS coordinates were found for the current filters. Review the raw data coordinates or broaden the filters.")
            else:
                assets = load_admin_assets(APP_DIR)
                html_map = leaflet_html(points, assets, initial_basemap=basemap, display_mode=display_mode, show_woreda_labels=show_woreda_labels, show_photos=show_photo_thumbnails)
                components.html(html_map, height=740, scrolling=False)
                st.caption("Street map: © OpenStreetMap contributors. Satellite imagery: © Esri and contributors. Administrative boundaries: Somali Region 2024 shapefile supplied for J-MAAP. 📷 indicates photo evidence. Images remain hosted by MoDa and are shown only when the current browser can access the MoDa link; J-MAAP does not copy the image files.")

                st.markdown('<div class="section-title">Mapped site register</div>', unsafe_allow_html=True)
                table = points[["site","zone","woreda","activities","providers","events","submissions","latest_date","photo_count","photo_types","finding_signals","severity","lat","lon"]].copy()
                table.columns = ["Site","Zone","Woreda","Activities","Providers","Monitoring events","Submissions","Latest monitoring","Photo evidence","Photo types","Finding signals","Highest severity","Latitude","Longitude"]
                st.dataframe(table, use_container_width=True, height=390, hide_index=True)
                st.download_button("Download mapped-site register (.csv)", table.to_csv(index=False).encode("utf-8-sig"), file_name=f"{reporting_month}_{sub_office.replace(' ', '_')}_Visited_Sites_Map_Register.csv", mime="text/csv")

                # -------------------------------------------------------------
                # PHOTO EVIDENCE GALLERY
                # -------------------------------------------------------------
                st.markdown('<div class="section-title">Monitoring photo evidence</div>', unsafe_allow_html=True)
                if not filtered_photos:
                    st.info("No MoDa photo attachment links were detected for the current filters.")
                else:
                    site_keys = sorted({(str(p.get("site") or ""), str(p.get("woreda") or ""), str(p.get("zone") or "")) for p in filtered_photos if p.get("site")})
                    label_map = {f"{site} — {woreda or 'Woreda not recorded'}{(' · ' + zone) if zone else ''}": (site, woreda, zone) for site, woreda, zone in site_keys}
                    selected_label = st.selectbox("Review photos for site", list(label_map.keys()), key="map_photo_site")
                    skey = label_map[selected_label]
                    gallery = [p for p in filtered_photos if (str(p.get("site") or ""), str(p.get("woreda") or ""), str(p.get("zone") or "")) == skey]
                    gallery = sorted(gallery, key=lambda x: (x.get("date", ""), x.get("photo_type", "")), reverse=True)
                    st.caption(f"{len(gallery)} photo link(s) for {skey[0]}. Click any image or 'Open full image' to open the original MoDa attachment. Access may require your existing MoDa/WFP authentication.")

                    max_gallery = min(len(gallery), 12)
                    for i in range(0, max_gallery, 3):
                        cols = st.columns(3)
                        for j, pht in enumerate(gallery[i:i+3]):
                            with cols[j]:
                                url = str(pht.get("url") or "")
                                safe_url = html.escape(url, quote=True)
                                ptype = html.escape(str(pht.get("photo_type") or "Monitoring photo"))
                                pdate = html.escape(str(pht.get("date") or ""))
                                provider = html.escape(str(pht.get("provider") or ""))
                                activity = html.escape(str(pht.get("activity") or ""))
                                st.markdown(
                                    f"""
<div style="background:white;border:1px solid #DCE6EF;border-radius:10px;padding:8px;margin-bottom:8px;min-height:245px">
  <a href="{safe_url}" target="_blank" rel="noopener noreferrer">
    <img src="{safe_url}" loading="lazy" referrerpolicy="no-referrer" style="width:100%;height:170px;object-fit:cover;border-radius:7px;background:#EEF3F6" onerror="this.style.display='none';document.getElementById('fallback-{i}-{j}').style.display='block';">
  </a>
  <div id="fallback-{i}-{j}" style="display:none;padding:55px 8px;text-align:center;background:#F1F6FA;border-radius:7px;color:#60758A">Preview unavailable — use the link below</div>
  <div style="font-weight:700;color:#123D67;margin-top:6px">{ptype}</div>
  <div style="font-size:12px;color:#60758A">{pdate}{' · ' + activity if activity else ''}{' · ' + provider if provider else ''}</div>
  <a href="{safe_url}" target="_blank" rel="noopener noreferrer" style="font-size:12px;color:#006AA6;font-weight:700">Open full image ↗</a>
</div>
""",
                                    unsafe_allow_html=True,
                                )
                    if len(gallery) > max_gallery:
                        st.info(f"Showing the 12 most recent photos for this site. The photo register below contains all {len(gallery)} links.")

                    preg = pd.DataFrame(gallery)
                    show_cols = [c for c in ["date","photo_type","activity","provider","event_key","question","source_sheet","source_file","uuid","url"] if c in preg.columns]
                    preg = preg[show_cols].copy()
                    preg = preg.rename(columns={"date":"Monitoring date","photo_type":"Photo type","activity":"Activity","provider":"Provider","event_key":"Monitoring Event ID","question":"Source question","source_sheet":"Sheet","source_file":"Source file","uuid":"Submission UUID","url":"Image link"})
                    st.dataframe(
                        preg,
                        use_container_width=True,
                        hide_index=True,
                        height=300,
                        column_config={"Image link": st.column_config.LinkColumn("Image link", display_text="Open image")},
                    )
                    st.download_button("Download photo-evidence register (.csv)", preg.to_csv(index=False).encode("utf-8-sig"), file_name=f"{reporting_month}_{sub_office.replace(' ', '_')}_Photo_Evidence_Register.csv", mime="text/csv")

# -----------------------------------------------------------------------------
# AAP / CFM
# -----------------------------------------------------------------------------
elif nav == "6. AAP / CFM":
    page_heading("6. AAP / CFM", "Track the community feedback journey from awareness and access through usage, response and satisfaction.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        rows = a["rows"]
        act_label = st.selectbox("Activity", ["Activity 1 (Relief response)", "Activity 2 (Nutrition assistance)", "Activity 3 (Refugee operations)", "Activity 6 (resilience)"], key="cfmact")
        res = cfm_journey(rows, activity=act_label)
        fdf = pd.DataFrame(res["stages"])
        if res["base"]:
            fdf["Percent of base"] = fdf["pct_base"] * 100
            c1, c2 = st.columns([1.15, 1])
            with c1:
                st.plotly_chart(px.funnel(fdf, x="count", y="stage", title=f"CFM journey — base n={res['base']}"), use_container_width=True)
            with c2:
                stage_cards = "".join([f'<div class="step-card"><b>{r.stage}</b><span style="float:right;color:#006EA9;font-weight:700">{r["Percent of base"]:.0f}%</span><div class="muted">{int(r["count"])} records</div></div>' for _, r in fdf.iterrows()])
                st.markdown(stage_cards, unsafe_allow_html=True)
            st.caption("Response and Satisfaction are only shown where supported by the activity module; Response uses a strict nested rule to avoid skip-logic inflation.")

        if act_label == "Activity 1 (Relief response)":
            frames = []
            for mod in ["Food", "Cash"]:
                rr = cfm_journey(rows, activity=act_label, modality=mod)
                for x in rr["stages"]:
                    frames.append({"Modality": mod, "Stage": x["stage"], "Percent of base": 100 * x["pct_base"] if x["pct_base"] is not None else None, "Base": rr["base"]})
            mdf = pd.DataFrame(frames)
            st.plotly_chart(px.line(mdf, x="Stage", y="Percent of base", color="Modality", markers=True, title="Relief CFM journey — Food vs Cash"), use_container_width=True)
            st.caption("Interpret modality differences descriptively: the August cash sample is geographically concentrated and should not be treated as a causal modality effect.")

        heat = cfm_woreda_table(rows, act_label)
        if not heat.empty:
            stages = [x for x in ["Awareness", "Access", "Usage", "Response", "Satisfaction"] if x in heat.columns]
            long = heat.melt(id_vars=["Woreda", "N"], value_vars=stages, var_name="Stage", value_name="Rate").dropna()
            if not long.empty:
                pivot = long.pivot(index="Woreda", columns="Stage", values="Rate") * 100
                st.plotly_chart(px.imshow(pivot, text_auto=".0f", aspect="auto", color_continuous_scale="Blues", title="Woreda CFM stage heatmap (%)"), use_container_width=True)
                st.dataframe(heat, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# PROTECTION
# -----------------------------------------------------------------------------
elif nav == "7. Protection, Safety & Dignity":
    page_heading("7. Protection, Safety & Dignity", "Review protection, inclusion, integrity and dignity signals while separating prevalence from seriousness.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        pdf = protection_summary(a["indicators"])
        if not pdf.empty:
            view = pdf.copy(); view["Issue rate %"] = view["Issue rate"].map(lambda x: 100 * x if pd.notna(x) else None)
            st.plotly_chart(px.bar(view.sort_values("Issue rate %"), x="Issue rate %", y="Finding", color="Activity", orientation="h", hover_data=["Severity", "Applicable", "Affected sites"], title="Protection / integrity assurance signals"), use_container_width=True)
            st.dataframe(view[["Activity", "Theme", "Finding", "Issues", "Applicable", "Issue rate %", "Severity", "Affected sites"]], use_container_width=True, hide_index=True)
            st.info("Frequency and seriousness are interpreted separately. Rare payment, misconduct or stock-control signals remain high-priority verification items even when prevalence is low.")

# -----------------------------------------------------------------------------
# ACTIVITY ANALYSIS
# -----------------------------------------------------------------------------
elif nav == "8. Activity Analysis":
    page_heading("8. Activity Analysis", "Review configured thematic findings by programme activity before moving to management follow-up.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        idf = pd.DataFrame(a["indicators"])
        if idf.empty:
            st.info("No configured indicators available.")
        else:
            activity_col = "Activity" if "Activity" in idf.columns else ("activity" if "activity" in idf.columns else None)
            if activity_col:
                acts = sorted(idf[activity_col].dropna().astype(str).unique().tolist())
                act = st.selectbox("Activity", acts)
                view = idf[idf[activity_col].astype(str) == act].copy()
            else:
                view = idf.copy()
            rate_col = next((c for c in ["Issue_Rate", "Issue rate", "issue_rate"] if c in view.columns), None)
            if rate_col:
                view["Issue rate %"] = pd.to_numeric(view[rate_col], errors="coerce") * 100
                finding_col = next((c for c in ["Finding", "finding", "Indicator", "indicator"] if c in view.columns), view.columns[0])
                st.plotly_chart(px.bar(view.sort_values("Issue rate %"), x="Issue rate %", y=finding_col, orientation="h", title="Configured issue rates"), use_container_width=True)
            st.dataframe(view, use_container_width=True, hide_index=True)

# -----------------------------------------------------------------------------
# FINDINGS & ACTIONS
# -----------------------------------------------------------------------------
elif nav == "9. Findings & Actions":
    page_heading("9. Findings & Actions", "Move from provisional site-level signals to aggregated, owned and verifiable management follow-up.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        ddf = pd.DataFrame(a["detail"]); adf = pd.DataFrame(a["agg"])
        c1, c2, c3 = st.columns(3)
        c1.metric("Detailed finding records", len(ddf)); c2.metric("Aggregated actions", len(adf)); c3.metric("Pending field validation", int((ddf.get("Validation_Status", pd.Series(dtype=str)) == "Pending field validation").sum()))
        st.markdown('<div class="section-title">Aggregated management actions</div>', unsafe_allow_html=True)
        st.dataframe(adf, use_container_width=True, height=330, hide_index=True)
        st.markdown('<div class="section-title">Detailed findings / validation register</div>', unsafe_allow_html=True)
        if not ddf.empty:
            c1, c2 = st.columns(2)
            with c1:
                act_opts = ["All"] + sorted(ddf["Activity"].dropna().unique().tolist()); act = st.selectbox("Activity", act_opts)
            with c2:
                sev_opts = ["All"] + sorted(ddf["Severity"].dropna().unique().tolist()); sev = st.selectbox("Severity", sev_opts)
            filtered = ddf.copy()
            if act != "All": filtered = filtered[filtered["Activity"] == act]
            if sev != "All": filtered = filtered[filtered["Severity"] == sev]
            st.dataframe(filtered, use_container_width=True, height=520, hide_index=True)

# -----------------------------------------------------------------------------
# OUTPUTS
# -----------------------------------------------------------------------------
elif nav == "10. Generate Outputs":
    page_heading("10. Generate Outputs", "Generate the monthly trackers, reports and machine-readable datasets from the validated analytical cohort.")
    a = st.session_state.analysis
    if not a:
        st.warning("Run the controlled monthly analysis first.")
    else:
        if st.button("Prepare downloadable outputs", type="primary", use_container_width=True):
            with st.spinner("Building trackers and reports..."):
                detail_xlsx = detailed_tracker_xlsx(a["detail"])
                agg_xlsx = aggregated_tracker_xlsx(a["agg"])
                mgmt_docx = management_report_docx(reporting_month, a["dq"], a["indicators"], a["agg"], a["detail"], sub_office=sub_office)
                dq_docx = dq_report_docx(reporting_month, a["dq"], a["indicators"], sub_office=sub_office)
                std_csv = rows_to_df(a["rows"]).drop(columns=["date_obj"], errors="ignore").to_csv(index=False).encode("utf-8-sig")
                ind_csv = pd.DataFrame(a["indicators"]).to_csv(index=False).encode("utf-8-sig")
                so_slug = sub_office.replace(" ", "_")
                bundle = bundle_zip({
                    f"{reporting_month}_{so_slug}_Detailed_Findings_Tracker.xlsx": detail_xlsx,
                    f"{reporting_month}_{so_slug}_Aggregated_Action_Tracker.xlsx": agg_xlsx,
                    f"{reporting_month}_{so_slug}_Management_Monitoring_Report.docx": mgmt_docx,
                    f"{reporting_month}_{so_slug}_Data_Quality_Report.docx": dq_docx,
                    f"{reporting_month}_{so_slug}_Standardized_Submissions.csv": std_csv,
                    f"{reporting_month}_{so_slug}_Indicator_Summary.csv": ind_csv,
                    f"{reporting_month}_{so_slug}_DQ_Summary.json": json.dumps(a["dq"], default=str, indent=2).encode("utf-8"),
                })
                st.session_state.generated_outputs = {"month": reporting_month, "sub_office": sub_office, "detail": detail_xlsx, "agg": agg_xlsx, "mgmt": mgmt_docx, "dq": dq_docx, "bundle": bundle}
        out = st.session_state.generated_outputs
        if out and out.get("month") == reporting_month and out.get("sub_office") == sub_office:
            so_slug = sub_office.replace(" ", "_")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("Detailed findings tracker (.xlsx)", out["detail"], f"{reporting_month}_{so_slug}_Detailed_Findings_Tracker.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button("Management monitoring report (.docx)", out["mgmt"], f"{reporting_month}_{so_slug}_Management_Monitoring_Report.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            with c2:
                st.download_button("Aggregated action tracker (.xlsx)", out["agg"], f"{reporting_month}_{so_slug}_Aggregated_Action_Tracker.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.download_button("Data quality report (.docx)", out["dq"], f"{reporting_month}_{so_slug}_Data_Quality_Report.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            st.download_button("Download complete monthly output package (.zip)", out["bundle"], f"{reporting_month}_{so_slug}_Monitoring_Output_Package.zip", "application/zip", type="primary", use_container_width=True)
            st.caption("The generated trackers preserve validation and management-agreement fields so automated signals can become agreed actions without losing the evidence trail.")
        else:
            st.markdown('<div class="status-info">Outputs are generated on demand so dashboard interaction remains responsive.</div>', unsafe_allow_html=True)
