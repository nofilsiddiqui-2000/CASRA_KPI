from pathlib import Path
import pandas as pd

from casra_dates import parse_date_range


ROOT_DIR = Path(r"C:\Users\B1020000\Documents\Nofil\Dashboards\CASRA MM Dashboard\CASRA-KPI-AUTOMATION") # path to the directory containing all the scripts and folders. This should be changed to match the user's environment.
# ROOT_DIR = Path(__file__).parent


SAP_DIR = ROOT_DIR / "SAP_Extracts"
LOOKUP_DIR = ROOT_DIR / "LookUp Tables"
OUTPUT_DIR = ROOT_DIR / "CASRA_KPI_OUTPUT"

ZMMR_DIR = SAP_DIR / "ZMMR2199M"
ZMNM_DIR = SAP_DIR / "ZMNM"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def validate_file(path: Path, label: str):
    if not path.exists():
        raise FileNotFoundError(f"{label} file not found: {path}")
    return path


date_from, date_to = parse_date_range("access-db")

ZMMR_FILE = validate_file(
    ZMMR_DIR / f"ZMMR2199M_{date_from}.xlsx",
    "ZMMR2199M"
)

ZMNM_FILE = validate_file(
    ZMNM_DIR / f"ZMNM_{date_from}.xlsx",
    "ZMNM"
)

RULE_SLOC_FILE = validate_file(LOOKUP_DIR / "RuleSloc.xlsx", "RuleSloc")
QMAT_MISSING_FILE = validate_file(LOOKUP_DIR / "QMatMissing.xlsx", "QMatMissing")
QMAT_RULES_FILE = validate_file(LOOKUP_DIR / "QMATRules.xlsx", "QMATRules")

OUTPUT_FILE = OUTPUT_DIR / f"CASRA_KPI_OUTPUT_{date_from}.xlsx"



CHECK_COLUMNS = [
    "Check_SNP",
    "Check_UofM",
    "Check_QMAT Extra",
    "Check_QMAT Missing",
    "Check_VType Extra",
    "Check_VType Missing",
    "Check_VType Error",
    "Check_SLoc Missing",
    "Check_SLoc_MRPInd",
    "Check_Batch",
    "Check_MOA",
    "Check_Missing_Model",
    "Check_Missing_MOA_Class",
    "Check_Class_Status",
    "Check_MRPArea",
]

FINAL_COLUMNS = [
    "SAPInt",
    "Material Number",
    "Created on",
    "Created",
    "Full Name",
    "Group",
    "MTyp",
    "Matl group",
    "ItCGr",
    "Plnt",
    "Check_Missing_Model",
    "Check_UofM",
    "Check_MOA",
    "Check_Missing_MOA_Class",
    "Check_Class_Status",
    "Check_SNP",
    "Check_Batch",
    "Check_SLoc Missing",
    "Check_SLoc_MRPInd",
    "Check_QMAT Missing",
    "Check_VType Missing",
    "Check_VType Extra",
    "Check_VType Error",
    "Check_QMAT Extra",
    "Check_MRPArea",
    "Errors",
]

QUERY31_COLUMNS = {
    "SAPInt": ["Material"],
    "Material Number": ["Material Number"],
    "Description": ["Description", "Material Description"],
    "Created on": ["Created on", "Created On"],
    "Created": ["Created"],
    "Full Name": ["Full Name", "Full Name of Person"],
    "Group": ["Group"],
    "MTyp": ["MTyp"],
    "Matl group": ["Matl group", "Matl Group", "Material Group"],
    "ItCGr": ["ItCGr"],
    "Plnt": ["Plnt", "Plant"],
}


def read_excel(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str)
    df.columns = df.columns.astype(str).str.strip()
    return clean_text_cols(df)

def clean_text_cols(df: pd.DataFrame) -> pd.DataFrame:
    # following all steps inside access for consistency
    for col in df.columns:
        df[col] = df[col].astype("string").replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    return df

def find_col(df: pd.DataFrame, possible_names: list[str]) -> str:
    lookup = {col.strip().lower(): col for col in df.columns}
    for name in possible_names:
        col = lookup.get(name.strip().lower())
        if col is not None:
            return col
    raise KeyError(f"Missing column. Tried: {possible_names}")


def norm_access(s: pd.Series) -> pd.Series:
    # Access text comparisons are case-insensitive, but imported values are not auto-trimmed.
    return s.fillna("").astype("string").str.lower()

def norm_material(s: pd.Series) -> pd.Series:
    return s.astype("string").str.strip()

def unique_materials(values) -> pd.Index:
    return pd.Index(pd.Series(values, dtype="string").str.strip().dropna().unique())

def flag_materials(ctx: dict, materials, check_col: str) -> None:
    zmnm = ctx["zmnm"]
    material_col = ctx["zmnm_material_col"]
    material_set = unique_materials(materials)

    if check_col not in zmnm.columns:
        zmnm[check_col] = 0

    zmnm[check_col] = pd.to_numeric(zmnm[check_col], errors="coerce").fillna(0).astype(int)
    mask = norm_material(zmnm[material_col]).isin(material_set)
    zmnm.loc[mask, check_col] = 1
    ctx["zmnm"] = zmnm


def flag_from_week1(ctx: dict, mask: pd.Series, check_col: str) -> None:
    week1 = ctx["week1"]
    flag_materials(ctx, week1.loc[mask, ctx["week1_material_col"]], check_col)


def concat_key(df: pd.DataFrame, columns: list[str]) -> pd.Series:
    key = pd.Series("", index=df.index, dtype="string")
    for col in columns:
        key = key + df[col].fillna("").astype("string")
    return key


def qmat_key(df: pd.DataFrame, mtyp_col: str, itcgr_col: str, plnt_col: str, actual_col: str) -> pd.Series:
    itcgr_part = norm_access(df[itcgr_col]).eq("erla").map({True: "ERLA", False: "BLANK"}).astype("string")
    return df[mtyp_col].fillna("").astype("string") + itcgr_part + df[plnt_col].fillna("").astype("string") + df[actual_col].fillna("").astype("string")


def load_context() -> dict:
    return {
        "zmnm": read_excel(ZMNM_FILE),
        "week1_raw": read_excel(ZMMR_FILE),
        "rule_sloc": read_excel(RULE_SLOC_FILE),
        "qmat_missing": read_excel(QMAT_MISSING_FILE),
        "qmat_rules": read_excel(QMAT_RULES_FILE),
    }


def set_pipeline_columns(ctx: dict) -> None:
    week1 = ctx["week1"]
    zmnm = ctx["zmnm"]

    ctx["field_col"] = find_col(week1, ["Field"])
    ctx["field_name_col"] = find_col(week1, ["Field Name"])
    ctx["report_col"] = find_col(week1, ["Report selection", "Report Selection"])
    ctx["actual_col"] = find_col(week1, ["Actuel Value", "Actuel Val"])
    ctx["mtyp_col"] = find_col(week1, ["MTyp"])
    ctx["itcgr_col"] = find_col(week1, ["ItCGr"])
    ctx["plnt_col"] = find_col(week1, ["Plnt", "Plant"])
    ctx["zmnm_material_col"] = find_col(zmnm, ["Material Number"])

    ctx["field"] = norm_access(week1[ctx["field_col"]])
    ctx["field_name"] = norm_access(week1[ctx["field_name_col"]])
    ctx["report"] = norm_access(week1[ctx["report_col"]])
    ctx["mtyp"] = norm_access(week1[ctx["mtyp_col"]])
    ctx["itcgr"] = norm_access(week1[ctx["itcgr_col"]])



###################### Access queries ######################


def q30_clean_week1(ctx: dict) -> None:
    week1 = ctx["week1_raw"]
    counter_col = find_col(week1, ["Counter"])
    material_col = find_col(week1, ["Material Number"])

    ctx["week1_material_col"] = material_col
    ctx["week1"] = week1[week1[counter_col].notna() & week1[material_col].notna()].copy()


def q31_insert_week1_into_zmnm(ctx: dict) -> None:
    week1 = ctx["week1"]
    zmnm = ctx["zmnm"]

    rows_to_insert = pd.DataFrame({
        target: week1[find_col(week1, source_options)]
        for target, source_options in QUERY31_COLUMNS.items()
    }).drop_duplicates()

    for col in rows_to_insert.columns:
        if col not in zmnm.columns:
            zmnm[col] = pd.NA

    ctx["zmnm"] = pd.concat(
        [zmnm, rows_to_insert.reindex(columns=zmnm.columns)],
        ignore_index=True,
    )

    set_pipeline_columns(ctx)


def q12_q32_missing_model(ctx: dict) -> None:
    flag_from_week1(ctx, ctx["field"].eq("zmm_moi_pn-admoi"), "Check_Missing_Model")


def q16_q33_uom(ctx: dict) -> None:
    flag_from_week1(ctx, ctx["field"].eq("mara-meins"), "Check_UofM")


def q13_q34_moa(ctx: dict) -> None:
    flag_from_week1(ctx, ctx["field"].eq("klah-versi"), "Check_MOA")


def q14_q35_missing_moa_class(ctx: dict) -> None:  
    flag_from_week1(ctx, ctx["field_name"].eq("moa class"), "Check_Missing_MOA_Class")


def q15_q36_class_status(ctx: dict) -> None:
    # Trailing space is intentional. This matches the Access query condition.
    flag_from_week1(ctx, ctx["field"].eq("klah-statu "), "Check_Class_Status")


def q18_q37_snp(ctx: dict) -> None:
    flag_from_week1(ctx, ctx["field"].eq("marc-sernp"), "Check_SNP")


def q17_q38_batch(ctx: dict) -> None:
    flag_from_week1(ctx, ctx["field"].isin(["mara-mhdrz", "mara-mhdhb", "marc-xchpf"]), "Check_Batch")


def q03_q04_q39_sloc_missing(ctx: dict) -> None:
    week1 = ctx["week1"]
    rule_key_col = find_col(ctx["rule_sloc"], ["Key"])
    rule_keys = ctx["rule_sloc"][rule_key_col].astype("string").dropna().unique()

    mask = ctx["report"].eq("plant sloc") & ctx["field_name"].eq("storage location")
    query03 = week1.loc[mask].copy()
    query03["Key"] = concat_key(query03, [ctx["mtyp_col"], ctx["itcgr_col"], ctx["plnt_col"], ctx["actual_col"]])

    flag_materials(
        ctx,
        query03.loc[query03["Key"].astype("string").isin(rule_keys), ctx["week1_material_col"]],
        "Check_SLoc Missing",
    )


def q05_q40_sloc_mrp_indicator(ctx: dict) -> None:
    flag_from_week1(ctx, ctx["report"].eq("plant sloc") & ctx["field"].eq("mard-diskz"), "Check_SLoc_MRPInd")


def q41_qmat_missing(ctx: dict) -> None:
    material_col = find_col(ctx["qmat_missing"], ["Material Number"])
    flag_materials(ctx, ctx["qmat_missing"][material_col], "Check_QMAT Missing")


def q10_q42_vtype_missing(ctx: dict) -> None:
    mtyp = ctx["mtyp"]
    itcgr = ctx["itcgr"]
    report = ctx["report"]
    field = ctx["field"]

    mask = (
        (report.eq("valuation type new") & field.eq("mvke-mtpos"))
        | (mtyp.eq("halb") & report.eq("valuation type no value") & field.eq("mvke-mtpos"))
        | (mtyp.eq("halb") & ~itcgr.eq("norm") & ~itcgr.eq("erla") & report.eq("valuation type core") & field.eq("mvke-mtpos"))
        | (mtyp.eq("halb") & ~itcgr.eq("norm") & ~itcgr.eq("erla") & report.eq("valuation type rotable") & field.eq("mvke-mtpos"))
    )

    flag_from_week1(ctx, mask, "Check_VType Missing")


def q11_q43_vtype_extra(ctx: dict) -> None:
    mtyp = ctx["mtyp"]
    itcgr = ctx["itcgr"]
    report = ctx["report"]
    field = ctx["field"]

    mask = (
        (~mtyp.eq("halb") & report.eq("valuation type no value") & field.eq("mvke-mtpos"))
        # Intentional: kept exactly like Access SQL, even though a value cannot be both norm and erla.
        | (mtyp.eq("halb") & itcgr.eq("norm") & itcgr.eq("erla") & report.eq("valuation type core") & field.eq("mvke-mtpos"))
        | (mtyp.eq("halb") & itcgr.eq("norm") & itcgr.eq("erla") & report.eq("valuation type rotable") & field.eq("mvke-mtpos"))
        | (report.str.startswith("valuation type") & field.eq("mara-mtart"))
    )

    flag_from_week1(ctx, mask, "Check_VType Extra")


def q09_q44_vtype_error(ctx: dict) -> None:
    week1 = ctx["week1"]
    mask = (
        ctx["report"].str.startswith("valuation type")
        & ctx["field"].isin(["mbew-bklas", "mbew-vprsv"])
        & week1[ctx["actual_col"]].notna()
    )

    flag_from_week1(ctx, mask, "Check_VType Error")


def q06_q081_q45_qmat_extra(ctx: dict) -> None:
    week1 = ctx["week1"]
    rules_key_col = find_col(ctx["qmat_rules"], ["Key"])
    rule_keys = ctx["qmat_rules"][rules_key_col].astype("string").dropna().unique()

    mask = ctx["report"].eq("plant data") & ctx["field"].eq("qmat-art")
    query06 = week1.loc[mask].copy()
    query06["Key"] = qmat_key(query06, ctx["mtyp_col"], ctx["itcgr_col"], ctx["plnt_col"], ctx["actual_col"])

    q081 = query06[query06[ctx["actual_col"]].notna() & ~query06["Key"].astype("string").isin(rule_keys)]
    flag_materials(ctx, q081[ctx["week1_material_col"]], "Check_QMAT Extra")


def q19_q46_mrp_area(ctx: dict) -> None:
    flag_from_week1(ctx, ctx["field"].eq("marc-diber"), "Check_MRPArea")


def q47_calculate_errors(ctx: dict) -> None:
    zmnm = ctx["zmnm"]

    for col in CHECK_COLUMNS:
        if col not in zmnm.columns:
            zmnm[col] = 0
        zmnm[col] = pd.to_numeric(zmnm[col], errors="coerce").fillna(0).astype(int)

    zmnm["Errors"] = zmnm[CHECK_COLUMNS].sum(axis=1)
    ctx["zmnm"] = zmnm


def prepare_final_output(ctx: dict) -> pd.DataFrame:
    zmnm = ctx["zmnm"]

    missing_cols = [col for col in FINAL_COLUMNS if col not in zmnm.columns]
    if missing_cols:
        raise KeyError(f"Missing final output columns: {missing_cols}")

    sapint_not_blank = zmnm["SAPInt"].fillna("").astype("string").str.strip().ne("")
    return zmnm.loc[sapint_not_blank, FINAL_COLUMNS].copy()


def run_pipeline(ctx: dict) -> None:
    # Data cleanup / insert
    q30_clean_week1(ctx)
    q31_insert_week1_into_zmnm(ctx)

    # Direct field / field-name checks
    q12_q32_missing_model(ctx)
    q16_q33_uom(ctx)
    q13_q34_moa(ctx)
    q14_q35_missing_moa_class(ctx)
    q15_q36_class_status(ctx)
    q18_q37_snp(ctx)
    q17_q38_batch(ctx)

    # SLoc checks
    q03_q04_q39_sloc_missing(ctx)
    q05_q40_sloc_mrp_indicator(ctx)

    # QMAT missing + valuation type checks
    q41_qmat_missing(ctx)
    q10_q42_vtype_missing(ctx)
    q11_q43_vtype_extra(ctx)
    q09_q44_vtype_error(ctx)

    # QMAT extra, MRP area, final total
    q06_q081_q45_qmat_extra(ctx)
    q19_q46_mrp_area(ctx)
    q47_calculate_errors(ctx)


def count_parts_created(ctx: dict) -> int:
    # Raw row count of populated Material Number cells in the ZMMR extract.
    # User-facing KPI: "X parts were created". Duplicates are kept on purpose
    # (one part may appear multiple times, once per check field in ZMMR).
    week1_raw = ctx["week1_raw"]
    material_col = find_col(week1_raw, ["Material Number"])
    return int(week1_raw[material_col].notna().sum())


def main() -> None:
    ctx = load_context()

    parts_created = count_parts_created(ctx)

    run_pipeline(ctx)

    final_output = prepare_final_output(ctx)
    rows_with_errors = int((final_output["Errors"] > 0).sum())

    run_summary = pd.DataFrame([{
        "Date From": date_from,
        "Date To": date_to,
        "Parts Created (ZMMR rows)": parts_created,
        "Rows in Output": len(final_output),
        "Rows with Errors (pre-SNP)": rows_with_errors,
    }])

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        final_output.to_excel(writer, sheet_name="Final Output", index=False)
        run_summary.to_excel(writer, sheet_name="Run Summary", index=False)

    print(f"Parts created (ZMMR rows): {parts_created}")
    print(f"Rows exported: {len(final_output)}")
    print(f"Rows with errors: {rows_with_errors}")
    print(f"Output created: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
