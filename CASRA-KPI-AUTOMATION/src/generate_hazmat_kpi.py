"""HazMat KPI step - runs between apply_snp_exceptions and generate_kpi_metrics.

Two modes, selected by HAZMAT_MODE below:

    "skip"  (default)  No SAP, no MM03. Every HAZ part read from ZMNM is
                       treated as correctly flagged, so Check_Hazards = 0 for
                       all of them and Hazmat % comes out 0.00%. HazMat errors
                       then have to be checked by hand in MM03 and patched into
                       CASRA_KPI_METRICS_MASTER.xlsx.

    "live"             Opens MM03 and validates each HAZ part for real.
                       NOT WORKING YET - the recorded navigation blocks in SAP.
                       The MM03 helpers are at the bottom of this file; once
                       they run cleanly, finish run_hazmat_live() and switch
                       HAZMAT_MODE to "live".

The business rule the live mode has to implement (unchanged from the spec in
CASRA KPI.md): for each ZMNM part with HazMat indicator = HAZ, read the Sales
Text long text in MM03 for all three sales orgs (3000 / 4200 / 1000, plant
3099, distribution channel 00). If every org reads "HAZARDOUS MATERIAL" the
part was flagged correctly -> Check_Hazards = 0. If any org shows anything else
-> Check_Hazards = 1. One error per part, however many orgs are wrong. Parts
that cannot be read at all are skipped, not appended, and listed on the
"Skipped Parts" sheet for manual follow-up.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd

import access_db
from casra_common import (
    CHECK_COLUMNS,
    CHECK_HAZARDS_COL,
    PARTS_CREATED_COL,
    REPORT_DATE_COL,
    SAP_DIR,
    coerce_check_columns,
    ensure_output_dirs,
    find_col,
    hazmat_kpi_output,
    normalize_material_key,
    parse_date_range,
    read_excel_access,
    read_run_summary,
    snp_final_output,
    validate_file,
)

# --- Mode ---------------------------------------------------------------
# "skip" assumes every HAZ part is correctly flagged; "live" validates in MM03.
# Flip to "live" only once run_hazmat_live() below actually works.
HAZMAT_MODE = "skip"

# --- Optional manual override (leave blank for a normal pipeline run) ---
ZMNM_FILE = r""

# Check_Hazards value assigned to every HAZ part when the MM03 pass is skipped.
ASSUMED_CHECK_HAZARDS = 0

# Written into the debug workbook so a skipped run can never be mistaken for a
# real MM03 validation later on. Value kept as-is so existing Run Summary
# readers do not change.
SKIP_NOTE = "DUMMY RUN - MM03 not opened"

HAZMAT_VALUE = "HAZ"
HAZMAT_INDICATOR_COLUMNS = [
    "HazMat indicator",
    "HazMat Indicator",
    "Hazmat indicator",
    "Haz Mat indicator",
]
MATERIAL_NUMBER_COLUMNS = ["Material Number"]
LEGACY_ERROR_TYPE_COL = "Error Type"

# Kept only so the debug workbook has the same shape as a live run.
SALES_ORGS = ["3000", "4200", "1000"]


# ----------------------------------------------------------------------------
# HAZ part selection + appending to Final Output
# ----------------------------------------------------------------------------

def filter_hazmat_parts(df: pd.DataFrame, hazmat_col: str) -> pd.DataFrame:
    hazmat = df[hazmat_col].fillna("").astype("string").str.strip().str.upper()
    material_col = find_col(df, MATERIAL_NUMBER_COLUMNS, "Material Number")
    return df.loc[hazmat.eq(HAZMAT_VALUE) & df[material_col].notna()].copy()


def hazmat_material_keys(hazmat_df: pd.DataFrame, material_col: str) -> list[str]:
    """Ordered, de-duplicated material keys for the HAZ parts."""
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
    """Append ZMNM HAZ rows with their Check_Hazards value (always 0 here)."""
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


def build_skipped_mode_mm03_sheet(material_keys: list[str]) -> pd.DataFrame:
    """Same shape as the live MM03 Results sheet, flagged as unvalidated."""
    rows = []
    for key in material_keys:
        for sales_org in SALES_ORGS:
            rows.append({
                "Material Number": key,
                "Sales Org": sales_org,
                "HazardText": SKIP_NOTE,
                "Correctly Flagged": "Not validated",
            })
    return pd.DataFrame(
        rows,
        columns=["Material Number", "Sales Org", "HazardText", "Correctly Flagged"],
    )


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
        "HazMat validation": SKIP_NOTE,
    }])


def run_hazmat_skip(final_path: Path, zmnm_path: Path) -> dict:
    zmnm_df = read_excel_access(zmnm_path)
    zmnm_df = access_db.ensure_description_column(zmnm_df)

    hazmat_col = find_col(zmnm_df, HAZMAT_INDICATOR_COLUMNS, "HazMat indicator")
    hazmat_df = filter_hazmat_parts(zmnm_df, hazmat_col)
    haz_rows_in_zmnm = len(hazmat_df)

    material_col = find_col(hazmat_df, MATERIAL_NUMBER_COLUMNS, "Material Number")
    material_keys = hazmat_material_keys(hazmat_df, material_col)

    # --- The only real difference from the live step ---
    # No SAP session, no MM03. Every HAZ part is assumed correctly flagged,
    # and nothing is skipped.
    check_by_material = {key: ASSUMED_CHECK_HAZARDS for key in material_keys}
    skipped: list[tuple[str, str]] = []
    skipped_keys: set[str] = set()

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
    run_summary["HazMat validation"] = SKIP_NOTE

    write_workbook(final_path, updated_final, run_summary, other_sheets)

    debug_file = hazmat_kpi_output(date.today().strftime("%Y%m%d"))
    mm03_results_df = build_skipped_mode_mm03_sheet(material_keys)
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


# ----------------------------------------------------------------------------
# Live MM03 validation - NOT WORKING YET
# ----------------------------------------------------------------------------
#
# Lifted from the old standalone hazmat_mm03.py. The navigation below currently
# blocks part-way through a run in SAP, which is why HAZMAT_MODE defaults to
# "skip" and HazMat errors are being checked by hand for now.
#
# To finish this off:
#   1. Get read_material_hazard_texts() running cleanly against a handful of
#      known HAZ parts (that is where the block happens).
#   2. Fill in run_hazmat_live() below - it only has to do what run_hazmat_skip
#      does, with check_by_material built from real MM03 reads instead of a
#      fixed ASSUMED_CHECK_HAZARDS, and unreadable parts collected into
#      `skipped` / `skipped_keys` rather than validated.
#   3. Set HAZMAT_MODE = "live".

MM03_LONGTEXT_PATH = (
    "wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/"
    "subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell"
)
MM03_PLANT = "3099"
MM03_DIST_CHANNEL = "00"

# A part is correctly flagged only when every sales org reads exactly this.
HAZARDOUS_TEXT = "HAZARDOUS MATERIAL"


def read_hazard_text(session) -> str:
    return session.findById(MM03_LONGTEXT_PATH).Text


def read_material_hazard_texts(session, material: str) -> list[tuple[str, str]]:
    """Read the MM03 Sales Text for one material across every sales org.

    Returns [(sales_org, hazard_text), ...]. Leaves the session back on the
    material-entry screen, ready for the next part.
    """
    import time

    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = material
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = len(material)
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08").select()

    results: list[tuple[str, str]] = []

    for index, sales_org in enumerate(SALES_ORGS):
        if index == 0:
            # First sales org needs the full org-level popup: plant, sales org,
            # distribution channel.
            session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = MM03_PLANT
            session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = sales_org
            session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = MM03_DIST_CHANNEL
            session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").setFocus()
            session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").caretPosition = len(MM03_DIST_CHANNEL)
        else:
            # Later sales orgs just reopen the popup via "Other org. levels".
            session.findById("wnd[0]/tbar[1]/btn[13]").press()
            session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = sales_org
            session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").setFocus()
            session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").caretPosition = len(sales_org)

        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        # Let the screen settle before reading the long-text control, so a lagging
        # control is not misread as blank.
        time.sleep(0.5)

        results.append((sales_org, read_hazard_text(session)))

    # Back out to the material-entry screen for the next material.
    session.findById("wnd[0]/tbar[0]/btn[3]").press()

    return results


def check_hazards_from_texts(hazard_texts: list[tuple[str, str]]) -> int:
    """Turn one part's MM03 reads into its Check_Hazards value.

    0 when every sales org reads HAZARDOUS MATERIAL, 1 otherwise. One error per
    part, no matter how many orgs are wrong.
    """
    if not hazard_texts:
        return 1
    for _sales_org, text in hazard_texts:
        if str(text or "").strip().upper() != HAZARDOUS_TEXT:
            return 1
    return 0


def run_hazmat_live(final_path: Path, zmnm_path: Path) -> dict:
    """Validate every HAZ part in MM03 and apply the result to Final Output."""
    raise NotImplementedError(
        "Live MM03 HazMat validation is not finished - the recorded navigation "
        "blocks mid-run in SAP. Keep HAZMAT_MODE = \"skip\" and verify HazMat "
        "errors by hand in MM03 until read_material_hazard_texts() runs cleanly. "
        "See the notes above this function."
    )


# ----------------------------------------------------------------------------
# Entry point
# ----------------------------------------------------------------------------

HAZMAT_RUNNERS = {
    "skip": run_hazmat_skip,
    "live": run_hazmat_live,
}


def main() -> None:
    if HAZMAT_MODE not in HAZMAT_RUNNERS:
        raise ValueError(
            f"HAZMAT_MODE is {HAZMAT_MODE!r}; expected one of {sorted(HAZMAT_RUNNERS)}."
        )

    date_from, date_to = parse_date_range(f"generate_hazmat_kpi ({HAZMAT_MODE})")
    ensure_output_dirs()

    final_path = validate_file(snp_final_output(date_from, date_to), "SNP Final output")
    zmnm_path = resolve_zmnm_path(date_from)

    if HAZMAT_MODE == "skip":
        print("*** HAZMAT MODE: SKIP - MM03 not opened, all HAZ parts assumed clean ***\n")

    result = HAZMAT_RUNNERS[HAZMAT_MODE](final_path, zmnm_path)

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

    if HAZMAT_MODE == "skip":
        print("\n*** Hazmat % is 0.00% because MM03 was not run. Verify HAZ parts manually. ***")


if __name__ == "__main__":
    main()
