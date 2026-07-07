"""cs-reports-v2 — a simpler, focused CS report.

Upload one XL-Connector .xlsx export (Case / Case Comment / Email Message / Case History),
see seven metrics in two tabs, and download a styled Excel copy. All pandas logic lives in
metrics.py; the Excel builder in excel_export.py. Nothing is persisted.
"""

import io

import pandas as pd
import streamlit as st

import metrics as m
import excel_export as xl

st.set_page_config(page_title="CS Reports v2", layout="wide")
st.title("CS Reports v2")


def _fmt(v):
    return "N/A" if v is None or (isinstance(v, float) and pd.isna(v)) else v


uploaded = st.sidebar.file_uploader(
    "Upload CS export workbook (.xlsx)", type=["xlsx"],
    help="XL-Connector export with Case / Case Comment / Email Message / Case History sheets.",
)
if uploaded is None:
    st.info(
        "Upload the XL-Connector export (.xlsx with `Case`, `Case Comment`, "
        "`Email Message`, and `Case History` sheets) to begin."
    )
    st.stop()


@st.cache_data(show_spinner="Loading and standardizing data...")
def _load(file_bytes: bytes) -> dict:
    return m.load_and_standardize(io.BytesIO(file_bytes))


@st.cache_data(show_spinner="Computing metrics...")
def _compute(file_bytes: bytes) -> dict:
    return m.compute_all(_load(file_bytes))


base = _load(uploaded.getvalue())
for w in base["warnings"]:
    st.warning(w)

results = _compute(uploaded.getvalue())

# ── Download button ──────────────────────────────────────────────────────────
st.sidebar.download_button(
    "⬇ Download Excel report",
    data=xl.build_workbook(results).getvalue(),
    file_name=f"cs_report_{pd.Timestamp.now().strftime('%Y-%m-%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

tab_timing, tab_flags = st.tabs(["Update Timing", "Flags & Backlog"])

# ── Tab 1: Update timing (metrics 1-4) ───────────────────────────────────────
with tab_timing:
    dsu = results["days_since_update"]
    gaps = results["update_gaps"]
    ri = results["resp_installer"]
    re = results["resp_incoming"]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Avg Days Since Update", _fmt(dsu["overall_avg"]),
              help="Open cases only. Today − latest of (Last Masternaut Update, Last Status Change, newest internal comment/email).")
    c2.metric("Avg Gap Between Updates", _fmt(gaps["overall_avg"]),
              help="Open + closed. Average days between consecutive internal updates on a case (60s-collapsed).")
    c3.metric("Avg Resp — Installer Comment", _fmt(ri["overall_avg"]),
              help=f"Open + closed. Days from a Partner User comment to the next internal update. {ri['unanswered']} unanswered.")
    c4.metric("Avg Resp — Incoming Email", _fmt(re["overall_avg"]),
              help=f"Open + closed. Days from an incoming email to the next internal update. {re['unanswered']} unanswered.")

    st.subheader("All Timing Metrics — By CS Owner")
    st.caption(
        "One row per CS Owner across all four metrics (average days). Blank = no data for that "
        "owner/metric (e.g. an owner with only closed cases has no Days Since Update, which is open-only)."
    )
    st.dataframe(results["timing_by_owner"], use_container_width=True, hide_index=True)

    st.subheader("Avg Days Since Last Update — By CS Owner (Open Cases)")
    st.caption(f"{dsu['total_open']} open cases. Snapshot metric — open only.")
    st.dataframe(dsu["per_owner"], use_container_width=True, hide_index=True)
    if not dsu["per_owner"].empty:
        st.bar_chart(dsu["per_owner"].set_index("CS Owner")["Avg Days"])

    st.subheader("Avg Gap Between Internal Updates — By CS Owner (Open + Closed)")
    st.dataframe(gaps["per_owner"], use_container_width=True, hide_index=True)

    st.subheader("Avg Response Time to Installer Comment — By CS Owner (Open + Closed)")
    st.caption(f"{ri['unanswered']} installer comment(s) had no subsequent internal update (excluded from average).")
    st.dataframe(ri["per_owner"], use_container_width=True, hide_index=True)

    st.subheader("Avg Response Time to Incoming Email — By CS Owner (Open + Closed)")
    st.caption(f"{re['unanswered']} incoming email(s) had no subsequent internal update (excluded from average).")
    st.dataframe(re["per_owner"], use_container_width=True, hide_index=True)

# ── Tab 2: Flags & backlog (metrics 5-7) ─────────────────────────────────────
with tab_flags:
    q = results["cs_queue_over_20"]
    na = results["no_account"]
    mc = results["missing_contact_owner"]

    c1, c2, c3 = st.columns(3)
    c1.metric("CS Queue > 20 Days (excl. L2)", _fmt(q["count"]))
    c2.metric("Cases > 1d Without Account", _fmt(na["count"]))
    c3.metric("Missing Contact / CS Owner", _fmt(mc["count"]))

    st.subheader(f"Open CS-Queue Cases Older Than 20 Days ({q['count']} cases)")
    st.caption(
        f"{q['removed_l2_count']} case(s) removed because they were escalated to "
        f"{m.L2_QUEUE_NAME} at some point. Queue matched on Owner == '{m.CS_QUEUE_NAME}' "
        "(verify this exact string against your export)."
    )
    st.dataframe(q["cases"], use_container_width=True, hide_index=True)

    st.subheader(f"Open Cases Older Than 1 Day Without an Account ({na['count']} cases)")
    st.dataframe(na["cases"], use_container_width=True, hide_index=True)

    st.subheader(f"Open Cases Older Than 1 Day Missing Contact / CS Owner ({mc['count']} cases)")
    st.dataframe(mc["cases"], use_container_width=True, hide_index=True)
