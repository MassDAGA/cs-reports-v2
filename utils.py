"""Column-name normalization for the CS Update Cadence data pipeline.

XL-Connector doesn't preserve Salesforce dot-notation column names like
``Owner.Name`` — the real export uses ``ColonSpace``-separated relationship
paths (e.g. ``CreatedBy : Full Name``, ``Triagedby__r : Full Name``) and
plain-English labels (``Is Incoming``, ``Changed Field``). This module maps
whatever variant shows up back to a single canonical name so the rest of the
pipeline can assume a fixed schema. Confirmed against the real export headers
on 2026-07-06; earlier guessed variants kept as fallbacks.
"""

import pandas as pd

# canonical name -> list of variant spellings to accept, canonical form first.
SHEET_COLUMN_ALIASES: dict[str, dict[str, list[str]]] = {
    "case": {
        "Id": ["Id", "Case ID", "18 Digit Case ID"],
        "CaseNumber": ["CaseNumber", "Case Number"],
        "CreatedDate": ["CreatedDate", "Created Date"],
        "Status": ["Status"],
        "Owner.Name": ["Owner.Name", "Owner : Full Name", "Owner : Name", "Owner: Name", "Owner Name", "Owner_Name"],
        "Last_Masternaut_Update__c": [
            "Last_Masternaut_Update__c", "Last Masternaut Update", "Last_Masternaut_Update",
        ],
        "Date_Time_of_Last_Status_Change__c": [
            "Date_Time_of_Last_Status_Change__c",
            "Date/Time of Last Status Change",
            "Date_Time_of_Last_Status_Change",
        ],
        "Triagedby__r.Name": [
            "Triagedby__r.Name", "Triagedby__r : Full Name", "Triaged by: Name",
            "TriagedBy.Name", "Triaged By", "Triagedby__r_Name",
        ],
        "Account.Name": [
            "Account.Name", "Account : Account Name", "Account: Account Name", "Account Name", "Account_Name",
        ],
        "ContactId": ["ContactId", "Contact ID"],
        "Contact.Name": [
            "Contact.Name", "Contact : Full Name", "Contact: Name", "Contact Name", "Contact_Name",
        ],
    },
    "comment": {
        "ParentId": ["ParentId", "Parent ID", "Case ID"],
        "CreatedDate": ["CreatedDate", "Created Date"],
        "CreatedBy.Name": [
            "CreatedBy.Name", "CreatedBy : Full Name", "Created By: Full Name",
            "Created By Name", "CreatedBy_Name",
        ],
        "CreatedBy.UserRole.Name": [
            "CreatedBy.UserRole.Name", "CreatedBy : UserRole : Name", "Created By: Role Name",
            "Created By User Role", "CreatedBy_UserRole_Name",
        ],
    },
    "email": {
        # NOTE: the real Email Message export calls its parent-case column "Case ID",
        # not "Parent ID" — both accepted.
        "ParentId": ["ParentId", "Case ID", "Parent ID"],
        "CreatedDate": ["CreatedDate", "Created Date"],
        "Incoming": ["Incoming", "Is Incoming"],
        "CreatedBy.Name": [
            "CreatedBy.Name", "CreatedBy : Full Name", "Created By: Full Name",
            "Created By Name", "CreatedBy_Name",
        ],
        "CreatedBy.UserRole.Name": [
            "CreatedBy.UserRole.Name", "CreatedBy : UserRole : Name", "Created By: Role Name",
            "Created By User Role", "CreatedBy_UserRole_Name",
        ],
    },
    "history": {
        "CaseId": ["CaseId", "Case ID"],
        "Field": ["Field", "Changed Field"],
        "OldValue": ["OldValue", "Old Value"],
        "NewValue": ["NewValue", "New Value"],
        "CreatedDate": ["CreatedDate", "Created Date"],
        "CreatedBy.Name": [
            "CreatedBy.Name", "CreatedBy : Full Name", "Created By: Full Name",
            "Created By Name", "CreatedBy_Name",
        ],
        "CreatedBy.UserRole.Name": [
            "CreatedBy.UserRole.Name", "CreatedBy : UserRole : Name", "Created By: Role Name",
            "Created By User Role", "CreatedBy_UserRole_Name",
        ],
    },
}


def normalize_columns(df: pd.DataFrame, sheet_key: str) -> pd.DataFrame:
    """Rename whichever alias variant is present to the canonical column name.

    Column headers are matched after stripping leading/trailing whitespace,
    since exports sometimes carry stray spaces.
    """
    df = df.rename(columns=lambda c: c.strip() if isinstance(c, str) else c)
    aliases = SHEET_COLUMN_ALIASES[sheet_key]
    rename_map = {}
    for canonical, variants in aliases.items():
        if canonical in df.columns:
            continue
        for variant in variants:
            if variant in df.columns:
                rename_map[variant] = canonical
                break
    return df.rename(columns=rename_map)


def require_columns(df: pd.DataFrame, required: list[str], sheet_key: str) -> None:
    """Raise a clear error naming exactly which canonical columns are missing."""
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(
            f"Sheet '{sheet_key}' is missing required column(s) after normalization: {missing}. "
            f"Columns found: {list(df.columns)}"
        )


def clean_name(name):
    """Strip stray non-ASCII encoding artifacts (e.g. 'Phil Sawyer\\xca' -> 'Phil Sawyer')."""
    import re

    if pd.isna(name):
        return name
    return re.sub(r"[^\x20-\x7E]", "", str(name)).strip()
