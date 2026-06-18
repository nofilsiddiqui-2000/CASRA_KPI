"""Apply HazMat KPI logic after SNP exceptions (with live MM03 validation).

Flow:
    1. Read SNP Final + ZMNM.
    2. Filter ZMNM for HazMat indicator = HAZ  ->  the HAZ part list.
    3. For each HAZ part, open MM03 and read the Sales-Text long text for all
       three sales orgs (3000 / 4200 / 1000, plant 3099, dist channel 00).
    4. Per part: if EVERY sales org reads "HAZARDOUS MATERIAL" the part was
       correctly flagged  ->  Check_Hazards = 0. If ANY sales org is anything
       else (blank, "HAZMAT", extra text, ...) the part is an error  ->
       Check_Hazards = 1. One error per part, regardless of how many orgs are
       wrong.
    5. Append each validated HAZ part to Final Output with its Check_Hazards
       value (0 or 1) and recalculate Errors.

Parts that cannot be validated in MM03 (material not found, long-text control
missing, or a SAP error mid-lookup) are skipped, not appended, and recorded on
the "Skipped Parts" sheet of the debug workbook for manual follow-up. A total
SAP failure (no session / missing credentials) stops the step loudly instead.

Final Output is built from ZMNM rows that have SAPInt (often populated via ZMMR
q31). HAZ parts typically exist only on the ZMNM extract and are not present in
Final Output, so each validated HAZ ZMNM row is appended with Material Number,
Created on, Created, etc.

Pipeline order:
    apply_snp_exceptions  ->  generate_hazmat_kpi  ->  generate_kpi_metrics
"""

from __future__ import annotations

import os
import time
import traceback
from datetime import date
from pathlib import Path

import pandas as pd
import win32com.client

import access_db
from casra_common import (
    CHECK_COLUMNS,
    CHECK_HAZARDS_COL,
    PARTS_CREATED_COL,
    REPORT_DATE_COL,
    SAP_DIR,
    SHAREPOINT_SNP_FINAL_DIR,
    coerce_check_columns,
    ensure_output_dirs,
    find_col,
    hazmat_kpi_output,
    mirror_to_sharepoint,
    normalize_material_key,
    parse_date_range,
    read_config,
    read_excel_access,
    read_run_summary,
    snp_final_output,
    validate_file,
)

# --- Optional manual overrides (leave blank for a normal pipeline run) ---
ZMNM_FILE = r""
# When non-empty, validate exactly these material numbers in MM03 instead of the
# HAZ parts read from ZMNM. Useful for exercising the SAP path during testing
# (the validated rows still only append for materials that exist in ZMNM HAZ).
MATERIALS_OVERRIDE: list[str] = []

HAZMAT_VALUE = "HAZ"
HAZMAT_INDICATOR_COLUMNS = [
    "HazMat indicator",
    "HazMat Indicator",
    "Hazmat indicator",
    "Haz Mat indicator",
]
MATERIAL_NUMBER_COLUMNS = ["Material Number"]
LEGACY_ERROR_TYPE_COL = "Error Type"

# --- MM03 validation parameters (fixed for the CASRA HazMat check) ---
HAZARD_PASS_TEXT = "HAZARDOUS MATERIAL"   # exact text a correctly flagged part shows
SALES_ORGS = ["3000", "4200", "1000"]
PLANT = "3099"
DIST_CHANNEL = "00"

LONGTEXT_PATH = (
    "wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/"
    "subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell"
)

# --- SAP logon (same system/pattern as the ZMNM and ZMMR extracts) ---
SAP_SYSTEM = "PR2"
SAP_CLIENT = "320"
SAPSHCUT_PATH = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe"


# ----------------------------------------------------------------------------
# MM03 hazard-text validation (SAP GUI scripting)
# ----------------------------------------------------------------------------

def normalize_hazard_text(text) -> str:
    """Collapse whitespace/newlines and uppercase, for tolerant comparison."""
    if text is None:
        return ""
    cleaned = str(text).replace(" ", " ")
    return " ".join(cleaned.split()).strip().upper()


def hazard_passes(text) -> bool:
    """True when a single sales org shows the correct hazardous-material text."""
    return normalize_hazard_text(text) == HAZARD_PASS_TEXT


def hazard_check_value(org_texts: dict[str, str]) -> int:
    """Check_Hazards for one part: 0 only if every sales org is correctly flagged."""
    if not org_texts or len(org_texts) < len(SALES_ORGS):
        return 1
    return 0 if all(hazard_passes(text) for text in org_texts.values()) else 1


def get_hazard_text(session) -> str:
    return session.findById(LONGTEXT_PATH).Text


def read_material_hazards(session, material: str) -> dict[str, str]:
    """Look up one material in MM03 and read the Sales Text for each sales org.

    Returns {sales_org: hazard_text}. Raises if any SAP navigation step fails;
    the caller treats that as a skipped part. The SAP click path is unchanged
    from the standalone hazmat_mm03 validation script.
    """
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = material
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = len(material)
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08").select()

    org_texts: dict[str, str] = {}
    for i, sales_org in enumerate(SALES_ORGS):
        if i == 0:
            # First sales org needs the full org-level popup: plant, sales org, dist channel
            session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = PLANT
            session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = sales_org
            session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = DIST_CHANNEL
            session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").setFocus()
            session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").caretPosition = len(DIST_CHANNEL)
        else:
            # Subsequent sales orgs just reopen the org-level popup via "Other org. levels"
            session.findById("wnd[0]/tbar[1]/btn[13]").press()
            session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = sales_org
            session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").setFocus()
            session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").caretPosition = len(sales_org)

        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Brief pause to let the screen fully refresh before reading the text box,
        # to avoid false-blank reads if the long-text control lags behind the rest.
        time.sleep(0.5)

        org_texts[sales_org] = get_hazard_text(session)

    # Back out to the initial material-entry screen for the next material
    session.findById("wnd[0]/tbar[0]/btn[3]").press()
    return org_texts


def connect_sap_session(config: dict[str, str]):
    """Launch SAP via sapshcut and return a scripting session, or raise loudly."""
    username = config.get("username", "")
    password = config.get("password", "")
    if not (username and password):
        raise RuntimeError(
            "Missing SAP GUI credentials in config.txt (username / password). "
            "Cannot run the MM03 HazMat validation."
        )

    os.system(
        f'''"{SAPSHCUT_PATH}" '''
        f'''-system={SAP_SYSTEM} -client={SAP_CLIENT} -user={username} -pw={password}'''
    )

    session = None
    connection = None
    application = None
    sap_gui_auto = None

    count = 0
    while not isinstance(session, win32com.client.CDispatch):
        time.sleep(1)
        count += 1
        if count > 60:
            raise RuntimeError("SAP GUI instance not found after 60s. Is SAP logon working?")
        try:
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            if not isinstance(sap_gui_auto, win32com.client.CDispatch):
                continue
            application = sap_gui_auto.GetScriptingEngine
            if not isinstance(application, win32com.client.CDispatch):
                sap_gui_auto = None
                continue
            connection = application.Children(0)
            if not isinstance(connection, win32com.client.CDispatch):
                application = None
                sap_gui_auto = None
                continue
            session = connection.Children(0)
            if not isinstance(session, win32com.client.CDispatch):
                connection = None
                application = None
                sap_gui_auto = None
                continue
        except Exception:
            continue

    return session


def reset_to_mm03(session) -> None:
    """Best-effort recovery to the MM03 initial screen after a failed lookup."""
    for _ in range(3):
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM03"
            session.findById("wnd[0]").sendVKey(0)
            return
        except Exception:
            # A modal popup may be blocking wnd[0]; try to dismiss it and retry.
            try:
                session.findById("wnd[1]").sendVKey(0)
            except Exception:
                try:
                    session.findById("wnd[1]").close()
                except Exception:
                    pass


def close_sap_session(session) -> None:
    try:
        session.findById("wnd[0]").close()
        try:
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except Exception:
            pass
    except Exception:
        pass


def run_mm03_validation(
    material_keys: list[str],
) -> tuple[dict[str, dict[str, str]], list[tuple[str, str]]]:
    """Validate each material in MM03.

    Returns:
        results_by_material: {material_key: {sales_org: hazard_text}}
        skipped:             [(material_key, reason)]  parts that could not be read
    """
    results_by_material: dict[str, dict[str, str]] = {}
    skipped: list[tuple[str, str]] = []

    config = read_config()
    session = connect_sap_session(config)
    try:
        session.findById("wnd[0]").maximize()
        session.findById("wnd[0]/tbar[0]/okcd").text = "MM03"
        session.findById("wnd[0]").sendVKey(0)

        for key in material_keys:
            try:
                results_by_material[key] = read_material_hazards(session, key)
            except Exception as exc:
                traceback.print_exc()
                skipped.append((key, f"{type(exc).__name__}: {exc}"))
                reset_to_mm03(session)
    finally:
        close_sap_session(session)

    return results_by_material, skipped


# ----------------------------------------------------------------------------
# HAZ part selection + appending to Final Output
# ----------------------------------------------------------------------------

def filter_hazmat_parts(df: pd.DataFrame, hazmat_col: str) -> pd.DataFrame:
    hazmat = df[hazmat_col].fillna("").astype("string").str.strip().str.upper()
    material_col = find_col(df, MATERIAL_NUMBER_COLUMNS, "Material Number")
    return df.loc[hazmat.eq(HAZMAT_VALUE) & df[material_col].notna()].copy()


def hazmat_material_keys(hazmat_df: pd.DataFrame, material_col: str) -> list[str]:
    """Ordered, de-duplicated material keys for the HAZ parts to validate."""
    keys: list[str] = []
    seen: set[str] = set()
    for value in hazmat_df[material_col]:
        key = normalize_material_key(value)
        if not key or key in seen:
            continue
        seen.add(key)
        keys.append(key)
    return keys


def zmnm_field_value(zmnm_row: pd.Series, source_names: list[str]):
    for name in source_names:
        if name in zmnm_row.index and pd.notna(zmnm_row[name]):
            text = str(zmnm_row[name]).strip()
            if text:
                return zmnm_row[name]
    return pd.NA


def build_haz_row(
    zmnm_row: pd.Series,
    output_columns: list[str],
    report_date,
    query31_columns: dict[str, list[str]],
    check_hazards: int,
) -> dict:
    row: dict = {}
    for col in output_columns:
        if col == REPORT_DATE_COL:
            row[col] = report_date
        elif col == CHECK_HAZARDS_COL:
            row[col] = check_hazards
        elif col in CHECK_COLUMNS:
            row[col] = 0
        elif col == "Errors":
            row[col] = check_hazards
        elif col in query31_columns:
            row[col] = zmnm_field_value(zmnm_row, query31_columns[col])
        elif col in zmnm_row.index and pd.notna(zmnm_row[col]):
            row[col] = zmnm_row[col]
        else:
            row[col] = pd.NA
    row[CHECK_HAZARDS_COL] = check_hazards
    row["Errors"] = check_hazards
    return row


def materials_with_hazards_flag(final_df: pd.DataFrame, material_col: str) -> set[str]:
    keys = final_df[material_col].map(normalize_material_key)
    haz_mask = pd.to_numeric(final_df[CHECK_HAZARDS_COL], errors="coerce").fillna(0).astype(int).eq(1)
    return {key for key, flagged in zip(keys, haz_mask) if key and flagged}


def apply_hazmat_to_final(
    final_df: pd.DataFrame,
    hazmat_df: pd.DataFrame,
    check_by_material: dict[str, int],
    skipped_keys: set[str],
    query31_columns: dict[str, list[str]],
) -> tuple[pd.DataFrame, int, int]:
    """Append validated ZMNM HAZ rows with their MM03-derived Check_Hazards value.

    A part is appended once, with Check_Hazards = 0 (correctly flagged in all
    sales orgs) or 1 (wrong in at least one). Parts that were skipped in MM03,
    or have no validation result, are not appended.
    """
    final_df = final_df.copy()
    final_df = final_df.drop(columns=[LEGACY_ERROR_TYPE_COL], errors="ignore")
    final_df = coerce_check_columns(final_df, CHECK_COLUMNS)

    material_col = find_col(final_df, MATERIAL_NUMBER_COLUMNS, "Material Number")
    haz_material_col = find_col(hazmat_df, MATERIAL_NUMBER_COLUMNS, "Material Number")

    already_appended = materials_with_hazards_flag(final_df, material_col)
    report_date = final_df[REPORT_DATE_COL].iloc[0] if REPORT_DATE_COL in final_df.columns else date.today()
    output_columns = list(final_df.columns)

    new_rows = []
    for _, zmnm_row in hazmat_df.iterrows():
        key = normalize_material_key(zmnm_row[haz_material_col])
        if not key or key in already_appended or key in skipped_keys:
            continue
        if key not in check_by_material:
            continue
        new_rows.append(
            build_haz_row(zmnm_row, output_columns, report_date, query31_columns, check_by_material[key])
        )
        already_appended.add(key)

    appended_rows = len(new_rows)
    if new_rows:
        final_df = pd.concat([final_df, pd.DataFrame(new_rows, columns=output_columns)], ignore_index=True)

    final_df["Errors"] = final_df[CHECK_COLUMNS].sum(axis=1)
    flagged = int(final_df[CHECK_HAZARDS_COL].sum())
    return final_df, flagged, appended_rows


# ----------------------------------------------------------------------------
# Workbook I/O
# ----------------------------------------------------------------------------

def read_workbook_sheets(path: Path) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, pd.DataFrame]]:
    xl = pd.ExcelFile(path)
    final_df = pd.read_excel(path, sheet_name="Final Output", dtype=object)
    run_summary = read_run_summary(path)
    other_sheets = {
        name: pd.read_excel(path, sheet_name=name, dtype=object)
        for name in xl.sheet_names
        if name not in {"Final Output", "Run Summary"}
    }
    return final_df, run_summary, other_sheets


def write_workbook(
    path: Path,
    final_df: pd.DataFrame,
    run_summary: pd.DataFrame,
    other_sheets: dict[str, pd.DataFrame],
) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Final Output", index=False)
        run_summary.to_excel(writer, sheet_name="Run Summary", index=False)
        for sheet_name, sheet_df in other_sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)


def resolve_zmnm_path(date_from: str) -> Path:
    if ZMNM_FILE.strip():
        return validate_file(Path(ZMNM_FILE.strip().strip('"').strip("'")), "ZMNM")
    return validate_file(SAP_DIR / "ZMNM" / f"ZMNM_{date_from}.xlsx", "ZMNM")


def build_mm03_results_sheet(results_by_material: dict[str, dict[str, str]]) -> pd.DataFrame:
    rows = []
    for key, org_texts in results_by_material.items():
        for sales_org in SALES_ORGS:
            text = org_texts.get(sales_org, "")
            rows.append({
                "Material Number": key,
                "Sales Org": sales_org,
                "HazardText": text,
                "Correctly Flagged": "Yes" if hazard_passes(text) else "No",
            })
    return pd.DataFrame(rows, columns=["Material Number", "Sales Org", "HazardText", "Correctly Flagged"])


def build_debug_summary(
    zmnm_path: Path,
    final_path: Path,
    haz_rows_in_zmnm: int,
    materials_validated: int,
    check_hazards_flagged: int,
    appended_rows: int,
    skipped_count: int,
    parts_created: int,
) -> pd.DataFrame:
    haz_pct = check_hazards_flagged / parts_created if parts_created else 0.0
    return pd.DataFrame([{
        REPORT_DATE_COL: date.today(),
        "Source ZMNM File": zmnm_path.name,
        "Source Final File": final_path.name,
        PARTS_CREATED_COL: parts_created,
        "HAZ rows in ZMNM": haz_rows_in_zmnm,
        "HAZ parts validated (MM03)": materials_validated,
        "HAZ parts skipped (MM03)": skipped_count,
        "Check_Hazards errors": check_hazards_flagged,
        "HAZ rows appended": appended_rows,
        "Hazmat %": haz_pct,
    }])


def apply_hazmat_pass(final_path: Path, zmnm_path: Path) -> dict:
    zmnm_df = read_excel_access(zmnm_path)
    zmnm_df = access_db.ensure_description_column(zmnm_df)

    hazmat_col = find_col(zmnm_df, HAZMAT_INDICATOR_COLUMNS, "HazMat indicator")
    hazmat_df = filter_hazmat_parts(zmnm_df, hazmat_col)
    haz_rows_in_zmnm = len(hazmat_df)

    material_col = find_col(hazmat_df, MATERIAL_NUMBER_COLUMNS, "Material Number")
    material_keys = MATERIALS_OVERRIDE or hazmat_material_keys(hazmat_df, material_col)

    # Only open SAP if there is actually something to validate.
    if material_keys:
        results_by_material, skipped = run_mm03_validation(material_keys)
    else:
        results_by_material, skipped = {}, []

    check_by_material = {key: hazard_check_value(orgs) for key, orgs in results_by_material.items()}
    skipped_keys = {key for key, _reason in skipped}

    final_df, run_summary, other_sheets = read_workbook_sheets(final_path)
    if PARTS_CREATED_COL not in run_summary.columns:
        raise KeyError(f"Run Summary missing '{PARTS_CREATED_COL}'. Run access_db.py first.")
    parts_created = int(run_summary[PARTS_CREATED_COL].iloc[0])

    updated_final, check_hazards_flagged, appended_rows = apply_hazmat_to_final(
        final_df, hazmat_df, check_by_material, skipped_keys, access_db.QUERY31_COLUMNS
    )

    run_summary["Check_Hazards errors"] = check_hazards_flagged
    run_summary["HAZ rows appended"] = appended_rows
    run_summary["HAZ parts validated (MM03)"] = len(check_by_material)
    run_summary["HAZ parts skipped (MM03)"] = len(skipped)
    run_summary["Rows with Errors (post-HAZ)"] = int((updated_final["Errors"] > 0).sum())

    write_workbook(final_path, updated_final, run_summary, other_sheets)

    debug_file = hazmat_kpi_output(date.today().strftime("%Y%m%d"))
    mm03_results_df = build_mm03_results_sheet(results_by_material)
    skipped_df = pd.DataFrame(skipped, columns=["Material Number", "Reason"])
    debug_summary = build_debug_summary(
        zmnm_path, final_path, haz_rows_in_zmnm, len(check_by_material),
        check_hazards_flagged, appended_rows, len(skipped), parts_created,
    )
    with pd.ExcelWriter(debug_file, engine="openpyxl") as writer:
        hazmat_df.to_excel(writer, sheet_name="HAZ Parts", index=False)
        mm03_results_df.to_excel(writer, sheet_name="MM03 Results", index=False)
        skipped_df.to_excel(writer, sheet_name="Skipped Parts", index=False)
        debug_summary.to_excel(writer, sheet_name="Summary", index=False)

    haz_pct = check_hazards_flagged / parts_created if parts_created else 0.0
    return {
        "haz_rows_in_zmnm": haz_rows_in_zmnm,
        "materials_validated": len(check_by_material),
        "skipped_count": len(skipped),
        "check_hazards_flagged": check_hazards_flagged,
        "appended_rows": appended_rows,
        "parts_created": parts_created,
        "haz_pct": haz_pct,
        "final_path": final_path,
        "debug_file": debug_file,
    }


def main() -> None:
    date_from, date_to = parse_date_range("generate_hazmat_kpi")
    ensure_output_dirs()

    final_path = validate_file(snp_final_output(date_from, date_to), "SNP Final output")
    zmnm_path = resolve_zmnm_path(date_from)

    result = apply_hazmat_pass(final_path, zmnm_path)
    sharepoint_copy = mirror_to_sharepoint(final_path, SHAREPOINT_SNP_FINAL_DIR)

    print(f"ZMNM file:                {zmnm_path}")
    print(f"SNP Final file:           {result['final_path']}")
    print(f"Parts created:            {result['parts_created']}")
    print(f"HAZ rows in ZMNM:         {result['haz_rows_in_zmnm']}")
    print(f"HAZ parts validated:      {result['materials_validated']}")
    print(f"HAZ parts skipped:        {result['skipped_count']}")
    print(f"Check_Hazards errors:     {result['check_hazards_flagged']}")
    print(f"HAZ rows appended:        {result['appended_rows']}")
    print(f"Hazmat %:                 {result['haz_pct']:.4f}  ({result['haz_pct']:.2%})")
    print(f"\nDebug file:               {result['debug_file']}")
    print(f"Power BI copy:            {sharepoint_copy}")


if __name__ == "__main__":
    main()
