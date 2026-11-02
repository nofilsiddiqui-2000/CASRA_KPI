"""ZMMR2199M extract - step 2 of the pipeline.

Runs the ZMMR2199M report in SAP for the given date range, downloads the
pipe-delimited list, parses it in Python and writes it as .xlsx.

Output: SAP_Extracts/ZMMR2199M/ZMMR2199M_<date_from>.xlsx  (read by access_db.py)
"""

import os
import time

import win32com.client

from casra_common import (
    SAP_DIR,
    close_sap_session,
    open_sap_session,
    parse_date_range,
    read_config,
)

EXPORT_DIR = SAP_DIR / "ZMMR2199M"


def parse_zmmr2199m_txt(txt_path):
    header = None
    rows = []

    with open(txt_path, "r", encoding="utf-8", errors="ignore") as f:
        for raw_line in f:
            line = raw_line.rstrip("\r\n")
            stripped = line.strip()

            if not stripped:
                continue
            if set(stripped) <= {"-"}:
                continue
            if "|" not in line:
                continue

            parts = [p.strip() for p in line.split("|")[1:-1]]
            if not parts or not any(parts):
                continue

            norm = [p.replace(" ", "").lower() for p in parts]

            if header is None:
                if "counter" in norm and "materialnumber" in norm:
                    header = parts
                continue

            if len(parts) < len(header):
                parts += [""] * (len(header) - len(parts))
            elif len(parts) > len(header):
                parts = parts[:len(header)]

            if not parts[0].strip().isdigit():
                continue

            rows.append(parts)

    if header is None:
        raise Exception("Could not find header row in ZMMR2199M export.")
    if not rows:
        raise Exception("No data rows found in ZMMR2199M export.")

    # SAP exports this column as "Rep. sel."; access_db.py expects "Report selection".
    header = ["Report selection" if h == "Rep. sel." else h for h in header]

    return header, rows


def save_rows_to_xlsx(headers, rows, xlsx_path, sheet_name="ZMMR2199M"):
    excel = None
    workbook = None
    try:
        if os.path.exists(xlsx_path):
            os.remove(xlsx_path)

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Add()
        worksheet = workbook.Worksheets(1)
        worksheet.Name = sheet_name

        all_rows = [headers] + rows
        row_count = len(all_rows)
        col_count = len(headers)

        data_range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(row_count, col_count))
        data_range.NumberFormat = "@"
        data_range.Value = tuple(tuple(r) for r in all_rows)

        header_range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, col_count))
        header_range.Font.Bold = True

        worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(row_count, col_count)).AutoFilter()
        worksheet.Columns.AutoFit()

        workbook.SaveAs(xlsx_path, FileFormat=51)

    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()


def run_extract(session, date_from, date_to, export_dir, txt_file_name,
                txt_file_path, xlsx_file_path):
    """Drive ZMMR2199M and leave the finished .xlsx at xlsx_file_path."""
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMMR2199M"
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
    session.findById("wnd[1]/tbar[0]/btn[2]").press()

    session.findById("wnd[0]/usr/ctxtSO_ERSDA-LOW").text = date_from
    session.findById("wnd[0]/usr/ctxtSO_ERSDA-HIGH").text = date_to
    session.findById("wnd[0]").sendVKey(0)

    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = export_dir
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = txt_file_name
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    try:
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    except Exception:
        pass

    time.sleep(2)

    if not os.path.exists(txt_file_path) or os.path.getsize(txt_file_path) == 0:
        raise RuntimeError("Export TXT file was not created correctly.")

    headers, rows = parse_zmmr2199m_txt(txt_file_path)
    save_rows_to_xlsx(headers, rows, xlsx_file_path)

    if not os.path.exists(xlsx_file_path) or os.path.getsize(xlsx_file_path) == 0:
        raise RuntimeError("Excel file was not created correctly.")

    try:
        os.remove(txt_file_path)
    except Exception:
        pass


def main():
    date_from, date_to = parse_date_range("Main_SAP_ZMMR2199M_xl")

    config = read_config()
    username = config.get("username", "")
    password = config.get("password", "")

    if not username or not password:
        raise RuntimeError("Missing SAP GUI credentials. Set username/password in config.txt.")

    EXPORT_DIR.mkdir(parents=True, exist_ok=True)

    txt_file_name = f"ZMMR2199M_{date_from}.txt"
    txt_file_path = str(EXPORT_DIR / txt_file_name)
    xlsx_file_path = str(EXPORT_DIR / f"ZMMR2199M_{date_from}.xlsx")

    for stale in (txt_file_path, xlsx_file_path):
        if os.path.exists(stale):
            os.remove(stale)

    session = open_sap_session(username, password)
    if session is None:
        raise RuntimeError("SAP GUI instance not found.")

    try:
        run_extract(
            session, date_from, date_to, str(EXPORT_DIR),
            txt_file_name, txt_file_path, xlsx_file_path,
        )
    finally:
        close_sap_session(session)

    print(f"ZMMR2199M extract created: {xlsx_file_path}")
    print(f"Date range: {date_from} -> {date_to}")


if __name__ == "__main__":
    main()
