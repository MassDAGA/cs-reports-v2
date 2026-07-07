"""Metric computations for cs-reports-v2.

A deliberately small, focused rewrite of the cs-reports pipeline. Every function is
DataFrame-in, DataFrame/dict-out — no Streamlit calls live here, so app.py can cache the
load step and the numbers can't drift between the dashboard and the Excel export.

Seven metrics in two groups:
  Group A — update timing (metrics 1-4)
  Group B — backlog & data-quality flags (metrics 5-7)

Definitions ported from the validated cs-reports/data_pipeline.py.
"""

import pandas as pd

from utils import normalize_columns, require_columns, clean_name

COLLAPSE_WINDOW_SEC = 60  # same-transaction dedup window, see collapse_same_transaction()

# Exact Owner.Name values on the Case export when a case's owner is a Queue, not a person.
# Only L2 was user-confirmed. If the real export's Owner value for the base queue differs,
# metric 5 reads zero until this constant is corrected. Flagged in the README.
CS_QUEUE_NAME = "US Customer Support Queue"
L2_QUEUE_NAME = "US Level 2 Customer Support"

SHEET_NAMES = {
    "case": "Case",
    "comment": "Case Comment",
    "email": "Email Message",
    "history": "Case History",
}


# ── LOAD & STANDARDIZE ──────────────────────────────────────────────────────

def load_and_standardize(file) -> dict:
    """Read the 4 sheets, normalize columns, standardize, and classify origin.

    Returns {cases, comments, emails, history, warnings}.
    """
    raw = {}
    for key, sheet_name in SHEET_NAMES.items():
        df = pd.read_excel(file, sheet_name=sheet_name)
        raw[key] = normalize_columns(df, key)

    require_columns(raw["case"], [
        "Id", "CaseNumber", "CreatedDate", "Status", "Owner.Name",
        "Last_Masternaut_Update__c", "Date_Time_of_Last_Status_Change__c",
        "Triagedby__r.Name", "Account.Name", "ContactId", "Contact.Name",
    ], "case")
    require_columns(raw["comment"], [
        "ParentId", "CreatedDate", "CreatedBy.Name", "CreatedBy.UserRole.Name",
    ], "comment")
    require_columns(raw["email"], [
        "ParentId", "CreatedDate", "Incoming", "CreatedBy.Name", "CreatedBy.UserRole.Name",
    ], "email")
    require_columns(raw["history"], [
        "CaseId", "Field", "OldValue", "NewValue", "CreatedDate",
        "CreatedBy.Name", "CreatedBy.UserRole.Name",
    ], "history")

    cases = _standardize_cases(raw["case"])
    comments, emails, history = _standardize_events(raw["comment"], raw["email"], raw["history"])
    warnings = _apply_origin(comments, emails, history)
    return {
        "cases": cases, "comments": comments, "emails": emails,
        "history": history, "warnings": warnings,
    }


def _standardize_cases(cases: pd.DataFrame) -> pd.DataFrame:
    cases = cases.rename(columns={
        "Id": "CaseId", "Owner.Name": "Owner",
        "Last_Masternaut_Update__c": "LastMasternautUpdate",
        "Date_Time_of_Last_Status_Change__c": "LastStatusChange",
        "Triagedby__r.Name": "CSOwner",
        "Account.Name": "AccountName", "Contact.Name": "ContactName",
    })
    cases["IsClosed"] = cases["Status"].isin(["Closed", "Cancelled"])
    cases["CreatedDate"] = pd.to_datetime(cases["CreatedDate"], utc=True, errors="coerce")
    for col in ["LastMasternautUpdate", "LastStatusChange"]:
        cases[col] = pd.to_datetime(cases[col], utc=True, errors="coerce")
    return cases


def _standardize_events(comments, emails, history):
    comments = comments.rename(columns={
        "ParentId": "CaseId", "CreatedBy.Name": "Agent", "CreatedBy.UserRole.Name": "Role",
    })
    comments["Type"] = "Comment"

    emails = emails.rename(columns={
        "ParentId": "CaseId", "CreatedBy.Name": "Agent", "CreatedBy.UserRole.Name": "Role",
    })
    emails["Type"] = "Email"
    # XL-Connector may export Is Incoming as text/0-1 rather than bool — coerce.
    if "Incoming" in emails.columns and emails["Incoming"].dtype != bool:
        emails["Incoming"] = (
            emails["Incoming"].astype(str).str.strip().str.lower().isin(["true", "1", "1.0", "yes"])
        )

    history = history.rename(columns={"CreatedBy.Name": "Agent", "CreatedBy.UserRole.Name": "Role"})
    history["Type"] = "Status Change"

    for df in (comments, emails, history):
        df["CreatedDate"] = pd.to_datetime(df["CreatedDate"], utc=True, errors="coerce")

    comments["Agent"] = comments["Agent"].apply(clean_name)
    emails["Agent"] = emails["Agent"].apply(clean_name)
    history["Agent"] = history["Agent"].apply(clean_name)
    return comments, emails, history


def _classify_origin(role) -> str:
    """Internal: role starts with 'NA '. Installer: role contains 'Partner User'.
    Everything else (incl. inbound customer emails, authored by a 'CEO'-role integration
    account) -> Customer/Other."""
    role = str(role) if pd.notna(role) else ""
    if role.startswith("NA "):
        return "Internal"
    if "Partner User" in role:
        return "Installer"
    return "Customer/Other"


def _apply_origin(comments, emails, history) -> list:
    warnings = []
    comments["Origin"] = comments["Role"].apply(_classify_origin)
    emails["Origin"] = emails["Role"].apply(_classify_origin)
    history["Origin"] = history["Role"].apply(_classify_origin)
    if "Incoming" in emails.columns:
        bad = emails[(emails["Incoming"] == True) & (emails["Origin"] != "Customer/Other")]  # noqa: E712
        if len(bad):
            warnings.append(
                f"{len(bad)} Incoming=True emails classified as non-Customer/Other. "
                "Check role assumptions."
            )
    return warnings


# ── INTERNAL UPDATE STREAM (shared by metrics 2, 3, 4) ──────────────────────

def build_internal_updates(cases, comments, emails, history) -> pd.DataFrame:
    """One row per collapsed internal update (comment/email/status-change by an NA role).
    Same-transaction events within 60s collapse to the earliest timestamp. Carries the
    case's CSOwner for per-owner attribution."""
    events = pd.concat([
        comments[comments["Origin"] == "Internal"][["CaseId", "CreatedDate"]],
        emails[emails["Origin"] == "Internal"][["CaseId", "CreatedDate"]],
        history[history["Origin"] == "Internal"][["CaseId", "CreatedDate"]],
    ], ignore_index=True).dropna(subset=["CreatedDate"])

    if events.empty:
        return events.assign(CSOwner=pd.Series(dtype="object"))

    events = events.sort_values(["CaseId", "CreatedDate"]).reset_index(drop=True)
    events["PrevDate"] = events.groupby("CaseId")["CreatedDate"].shift(1)
    new_touch = events["PrevDate"].isna() | (
        (events["CreatedDate"] - events["PrevDate"]).dt.total_seconds() > COLLAPSE_WINDOW_SEC
    )
    events["TouchGroup"] = new_touch.groupby(events["CaseId"]).cumsum()

    updates = events.groupby(["CaseId", "TouchGroup"]).agg(
        UpdateTime=("CreatedDate", "min"),
    ).reset_index().drop(columns=["TouchGroup"])

    owner_map = cases.set_index("CaseId")["CSOwner"]
    updates["CSOwner"] = updates["CaseId"].map(owner_map)
    return updates.sort_values(["CaseId", "UpdateTime"]).reset_index(drop=True)


def _grouped_and_per_owner(df: pd.DataFrame, value_col: str, count_label: str) -> dict:
    """Pooled overall mean + per-CS-Owner table (Cases/Avg/Max) for a value column."""
    if df.empty:
        empty = pd.DataFrame(columns=["CS Owner", count_label, "Avg Days", "Max Days"])
        return {"overall_avg": None, "per_owner": empty}
    owner = df.assign(_Owner=df["CSOwner"].fillna("Unassigned"))
    per_owner = (
        owner.groupby("_Owner")
        .agg(N=("CaseId", "count"), AvgDays=(value_col, "mean"), MaxDays=(value_col, "max"))
        .reset_index().sort_values("AvgDays", ascending=False)
    )
    per_owner["AvgDays"] = per_owner["AvgDays"].round(2)
    per_owner["MaxDays"] = per_owner["MaxDays"].round(2)
    per_owner.columns = ["CS Owner", count_label, "Avg Days", "Max Days"]
    return {"overall_avg": round(df[value_col].mean(), 2), "per_owner": per_owner}


# ── METRIC 1: Days since last update (OPEN cases) ───────────────────────────

def compute_days_since_update_open(cases, comments, emails) -> dict:
    cases_open = cases[~cases["IsClosed"]].copy()
    updates = pd.concat([
        comments[comments["Origin"] == "Internal"][["CaseId", "CreatedDate"]],
        emails[emails["Origin"] == "Internal"][["CaseId", "CreatedDate"]],
    ], ignore_index=True)
    updates = updates[updates["CaseId"].isin(cases_open["CaseId"])]

    latest = (
        updates.sort_values("CreatedDate", ascending=False)
        .groupby("CaseId").first().reset_index()
        .rename(columns={"CreatedDate": "LastActivityDate"})
    ) if len(updates) else pd.DataFrame(columns=["CaseId", "LastActivityDate"])

    merged = cases_open.merge(latest, on="CaseId", how="left")
    if "LastActivityDate" not in merged.columns:
        merged["LastActivityDate"] = pd.NaT
    merged["TrueLastUpdate"] = merged[
        ["LastMasternautUpdate", "LastStatusChange", "LastActivityDate"]
    ].max(axis=1)
    now = pd.Timestamp.now(tz="UTC")
    merged["DaysSinceUpdate"] = (now - merged["TrueLastUpdate"]).dt.days

    res = _grouped_and_per_owner(
        merged.rename(columns={"DaysSinceUpdate": "_v"}), "_v", "Open Cases"
    )
    res["total_open"] = len(merged)
    res["detail"] = merged
    return res


# ── METRIC 2: Gap between consecutive internal updates (OPEN + CLOSED) ───────

def compute_update_gaps(updates: pd.DataFrame) -> dict:
    """Elapsed days between consecutive internal updates on the same case, pooled."""
    if updates.empty:
        return _grouped_and_per_owner(updates.assign(_v=[]), "_v", "Gaps")
    df = updates.sort_values(["CaseId", "UpdateTime"]).reset_index(drop=True)
    df["PrevUpdate"] = df.groupby("CaseId")["UpdateTime"].shift(1)
    df["GapDays"] = (df["UpdateTime"] - df["PrevUpdate"]).dt.total_seconds() / 86400
    gaps = df.dropna(subset=["GapDays"]).copy()
    return _grouped_and_per_owner(gaps.rename(columns={"GapDays": "_v"}), "_v", "Gaps")


# ── METRICS 3 & 4: Response time from external input to next internal update ─

def compute_response_time(external: pd.DataFrame, updates: pd.DataFrame, cases: pd.DataFrame) -> dict:
    """For each external input, days until the next internal update on the same case.
    Unanswered inputs (no later update) are excluded from the average and counted.

    external: DataFrame with CaseId, CreatedDate (the inbound events).
    updates:  the collapsed internal-update stream (CaseId, UpdateTime).
    """
    ext = external.dropna(subset=["CreatedDate"])[["CaseId", "CreatedDate"]].copy()
    if ext.empty:
        res = _grouped_and_per_owner(ext.assign(_v=[]), "_v", "Responses")
        res["unanswered"] = 0
        return res

    # merge_asof requires each `on` key globally sorted (the `by` grouping is separate).
    upd = updates[["CaseId", "UpdateTime"]].sort_values("UpdateTime")
    ext = ext.rename(columns={"CreatedDate": "InputTime"}).sort_values("InputTime")

    # For each external event, find the earliest internal update strictly after it, same case.
    # direction='forward' + allow_exact_matches=False gives exactly that.
    paired = pd.merge_asof(
        ext, upd,
        left_on="InputTime", right_on="UpdateTime",
        by="CaseId", direction="forward", allow_exact_matches=False,
    )
    paired["RespDays"] = (paired["UpdateTime"] - paired["InputTime"]).dt.total_seconds() / 86400
    answered = paired.dropna(subset=["RespDays"]).copy()
    unanswered = int(paired["UpdateTime"].isna().sum())

    owner_map = cases.set_index("CaseId")["CSOwner"]
    answered["CSOwner"] = answered["CaseId"].map(owner_map)

    res = _grouped_and_per_owner(answered.rename(columns={"RespDays": "_v"}), "_v", "Responses")
    res["unanswered"] = unanswered
    return res


# ── METRIC 5: CS Queue cases > 20 days, minus ever-escalated-to-L2 ──────────

def compute_l2_routed_case_ids(history: pd.DataFrame, l2_queue_name: str = L2_QUEUE_NAME) -> set:
    """Case IDs ever routed to L2, per Case History Owner-change rows (matches Field on
    'owner' case-insensitively; checks both Old and New value)."""
    owner_changes = history[history["Field"].astype(str).str.contains("owner", case=False, na=False)]
    hit = owner_changes[
        owner_changes["OldValue"].astype(str).str.strip().eq(l2_queue_name)
        | owner_changes["NewValue"].astype(str).str.strip().eq(l2_queue_name)
    ]
    return set(hit["CaseId"].unique())


def compute_cs_queue_over_20(cases, history,
                             cs_queue_name: str = CS_QUEUE_NAME,
                             l2_queue_name: str = L2_QUEUE_NAME) -> dict:
    l2_ids = compute_l2_routed_case_ids(history, l2_queue_name)
    owner_stripped = cases["Owner"].astype(str).str.strip()
    now = pd.Timestamp.now(tz="UTC")
    age = (now - cases["CreatedDate"]).dt.days

    in_cs_queue_over_20 = ~cases["IsClosed"] & (owner_stripped == cs_queue_name) & (age > 20)
    kept = cases[in_cs_queue_over_20 & ~cases["CaseId"].isin(l2_ids)].copy()
    kept["CaseAgeDays"] = (now - kept["CreatedDate"]).dt.days
    removed = int((in_cs_queue_over_20 & cases["CaseId"].isin(l2_ids)).sum())

    out = kept[["CaseNumber", "CreatedDate", "CaseAgeDays", "Status", "CSOwner", "AccountName"]].copy()
    out["CreatedDate"] = out["CreatedDate"].dt.tz_localize(None)
    out["CSOwner"] = out["CSOwner"].fillna("Unassigned")
    out.columns = ["Case Number", "Created Date", "Age (Days)", "Status", "CS Owner", "Account"]
    out = out.sort_values("Age (Days)", ascending=False)
    return {"count": len(kept), "removed_l2_count": removed, "cases": out}


# ── METRICS 6 & 7: Data-quality flags (OPEN cases, age > 1 day) ─────────────

def _open_older_than_one_day(cases: pd.DataFrame) -> pd.DataFrame:
    now = pd.Timestamp.now(tz="UTC")
    age = (now - cases["CreatedDate"]).dt.days
    scoped = cases[~cases["IsClosed"] & (age > 1)].copy()
    scoped["CaseAgeDays"] = (now - scoped["CreatedDate"]).dt.days
    return scoped


def cases_no_account(cases: pd.DataFrame) -> dict:
    scoped = _open_older_than_one_day(cases)
    hit = scoped[scoped["AccountName"].isna()].copy()
    out = hit[["CaseNumber", "CreatedDate", "CaseAgeDays", "Status", "CSOwner", "ContactName"]].copy()
    out["CreatedDate"] = out["CreatedDate"].dt.tz_localize(None)
    out["CSOwner"] = out["CSOwner"].fillna("Unassigned")
    out.columns = ["Case Number", "Created Date", "Age (Days)", "Status", "CS Owner", "Contact"]
    out = out.sort_values("Age (Days)", ascending=False)
    return {"count": len(hit), "cases": out}


def cases_missing_contact_or_owner(cases: pd.DataFrame) -> dict:
    scoped = _open_older_than_one_day(cases)
    missing_contact = scoped["ContactName"].isna()
    missing_owner = scoped["CSOwner"].isna()
    hit = scoped[missing_contact | missing_owner].copy()
    hit["Missing Contact"] = hit["ContactName"].isna().map({True: "Yes", False: "No"})
    hit["Missing CS Owner"] = hit["CSOwner"].isna().map({True: "Yes", False: "No"})
    out = hit[["CaseNumber", "CreatedDate", "CaseAgeDays", "Status", "AccountName",
               "Missing Contact", "Missing CS Owner"]].copy()
    out["CreatedDate"] = out["CreatedDate"].dt.tz_localize(None)
    out.columns = ["Case Number", "Created Date", "Age (Days)", "Status", "Account",
                   "Missing Contact", "Missing CS Owner"]
    out = out.sort_values("Age (Days)", ascending=False)
    return {"count": len(hit), "cases": out}


# ── TOP-LEVEL ORCHESTRATION ─────────────────────────────────────────────────

def compute_all(base: dict) -> dict:
    """Run all seven metrics against a standardized base (from load_and_standardize)."""
    cases, comments, emails, history = base["cases"], base["comments"], base["emails"], base["history"]
    updates = build_internal_updates(cases, comments, emails, history)

    incoming_emails = emails[emails["Incoming"] == True] if "Incoming" in emails.columns else emails.iloc[0:0]  # noqa: E712
    installer_comments = comments[comments["Origin"] == "Installer"]

    return {
        "days_since_update": compute_days_since_update_open(cases, comments, emails),   # metric 1
        "update_gaps": compute_update_gaps(updates),                                    # metric 2
        "resp_installer": compute_response_time(installer_comments, updates, cases),    # metric 3
        "resp_incoming": compute_response_time(incoming_emails, updates, cases),        # metric 4
        "cs_queue_over_20": compute_cs_queue_over_20(cases, history),                   # metric 5
        "no_account": cases_no_account(cases),                                          # metric 6
        "missing_contact_owner": cases_missing_contact_or_owner(cases),                 # metric 7
        "total_cases_all": len(cases),
        "warnings": base.get("warnings", []),
    }
