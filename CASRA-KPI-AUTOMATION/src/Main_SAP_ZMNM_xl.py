"""ZMNM extract - step 1 of the pipeline.

Runs the ZMNM report in SAP for the given date range, downloads it as a
tab-delimited .txt, converts it to .xlsx and deletes the .txt.

Output: SAP_Extracts/ZMNM/ZMNM_<date_from>.xlsx  (read by access_db.py)
"""

import os
import time

import win32com.client

from casra_common import (
    SAP_DIR,
    close_sap_session,
    config_list,
    open_sap_session,
    parse_date_range,
    read_config,
)

EXPORT_DIR = SAP_DIR / "ZMNM"

PLANT = "3099"

# The "Created by" multiple-selection table shows 8 rows at a time. Entries
# past the eighth are reached by scrolling, which is what enter_created_by
# below reproduces.
VISIBLE_SELECTION_ROWS = 8

SELECTION_TABLE = (
    "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE"
)
SELECTION_FIELD = SELECTION_TABLE + "/txtRSCSEL_255-SLOW_I[1,{row}]"


def convert_txt_to_xlsx(txt_path, xlsx_path):
    excel = None
    workbook = None
    try:
        if os.path.exists(xlsx_path):
            os.remove(xlsx_path)

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        excel.Workbooks.OpenText(
            Filename=txt_path,
            DataType=1,
            Tab=True,
            Semicolon=False,
            Comma=False,
            Space=False,
            Other=False
        )

        workbook = excel.ActiveWorkbook
        workbook.SaveAs(xlsx_path, FileFormat=51)  # .xlsx

    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()


def enter_created_by(session, user_ids):
    """Type the "Created by" user IDs into the multiple-selection screen.

    The IDs come from config.txt (zmnm_created_by). They are the KPI's scope
    filter, so the list must match what the business expects.
    """
    last_field = None
    last_value = ""

    for index, user_id in enumerate(user_ids):
        if index < VISIBLE_SELECTION_ROWS:
            row = index
        else:
            # Scroll so this entry lands on the row after the previous one.
            session.findById(SELECTION_TABLE).verticalScrollbar.position = index - 6
            row = 6

        last_field = SELECTION_FIELD.format(row=row)
        last_value = user_id
        session.findById(last_field).text = user_id

    if last_field is not None:
        session.findById(last_field).setFocus()
        session.findById(last_field).caretPosition = len(last_value)


def run_extract(session, date_from, date_to, created_by, txt_file_path, xlsx_file_path):
    """Drive ZMNM and leave the finished .xlsx at xlsx_file_path."""
    session.findById("wnd[0]").maximize()
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZMNM"
    session.findById("wnd[0]").sendVKey(0)

    # low date
    session.findById("wnd[0]/usr/ctxtSO_ERSDA-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtSO_ERSDA-LOW").caretPosition = 0
    session.findById("wnd[0]").sendVKey(4)
    session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").focusDate = date_from
    session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").firstVisibleDate = date_from
    session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = f"{date_from},{date_from}"

    # high date
    session.findById("wnd[0]/usr/ctxtSO_ERSDA-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtSO_ERSDA-HIGH").caretPosition = 0
    session.findById("wnd[0]").sendVKey(4)
    session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").focusDate = date_to
    session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = f"{date_to},{date_to}"

    # created by filter
    session.findById("wnd[0]/usr/txtSO_ERNAM-LOW").setFocus()
    session.findById("wnd[0]/usr/txtSO_ERNAM-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_SO_ERNAM_%_APP_%-VALU_PUSH").press()

    enter_created_by(session, created_by)
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # plant filter
    session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").text = PLANT
    session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").setFocus()
    session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").caretPosition = 4

    # execute
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    # export txt first
    session.findById("wnd[0]/usr/cntlMY_AREA_CONTROL/shellcont/shell").pressToolbarButton("Z_DOWNLOAD")
    session.findById("wnd[1]/usr/ctxtRLGRAP-FILENAME").text = txt_file_path
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    try:
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
    except Exception:
        pass

    time.sleep(2)

    if not os.path.exists(txt_file_path) or os.path.getsize(txt_file_path) == 0:
        raise RuntimeError("Export TXT file was not created correctly.")

    convert_txt_to_xlsx(txt_file_path, xlsx_file_path)

    if not os.path.exists(xlsx_file_path) or os.path.getsize(xlsx_file_path) == 0:
        raise RuntimeError("Excel file was not created correctly.")

    try:
        os.remove(txt_file_path)
    except Exception:
        pass


def main():
    date_from, date_to = parse_date_range("Main_SAP_ZMNM_xl")

    config = read_config()
    username = config.get("username", "")
    password = config.get("password", "")
    created_by = config_list(config, "zmnm_created_by")

    if not username or not password:
        raise RuntimeError("Missing SAP GUI credentials. Set username/password in config.txt.")
    if not created_by:
        raise RuntimeError(
            "No ZMNM 'Created by' users configured. Set zmnm_created_by in config.txt "
            "(comma-separated SAP user IDs)."
        )

    EXPORT_DIR.mkdir(parents=True, exist_ok=True)

    txt_file_path = str(EXPORT_DIR / f"ZMNM_{date_from}.txt")
    xlsx_file_path = str(EXPORT_DIR / f"ZMNM_{date_from}.xlsx")

    for stale in (txt_file_path, xlsx_file_path):
        if os.path.exists(stale):
            os.remove(stale)

    session = open_sap_session(username, password)
    if session is None:
        raise RuntimeError("SAP GUI instance not found.")

    try:
        run_extract(session, date_from, date_to, created_by, txt_file_path, xlsx_file_path)
    finally:
        close_sap_session(session)

    print(f"ZMNM extract created: {xlsx_file_path}")
    print(f"Date range: {date_from} -> {date_to}")
    print(f"Created by: {', '.join(created_by)}")


if __name__ == "__main__":
    main()
