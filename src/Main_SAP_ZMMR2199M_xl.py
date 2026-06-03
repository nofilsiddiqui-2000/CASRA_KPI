from datetime import datetime, timedelta
import os, time
import win32com.client

from casra_config import read_config
from casra_dates import parse_date_range


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

config = read_config()

asset_num = config.get("asset_num", "")
if asset_num == "P22208617":
    execution_line = "Executed on LAPTOP (MTL)"
else:
    execution_line = f"Executed Manually on {asset_num}"

automation_status = "On"

if automation_status == "On":

    status = {
        "Meta": {
            "Execution Line": execution_line,
            "Start Time": datetime.today().strftime("%H:%M")
        },
        "Reporting Period": datetime.today().strftime("%b %d, %Y"),
        "Error Count": 0,
        "ZMMR2199M": "null"
    }

    session = None
    connection = None
    application = None
    SapGuiAuto = None

    username = config.get("username", "")
    password = config.get("password", "")

    if username and password:
        # export_path = r"C:\Users\B1020000\Bombardier\Team CM Aftermarket - SAP Reports\ZMMR2199M" # disabling for testing.
        export_path = r"C:\Users\B1020000\Documents\Nofil\Dashboards\CASRA MM Dashboard\CASRA-KPI-AUTOMATION\SAP_Extracts\ZMMR2199M"   
        os.makedirs(export_path, exist_ok=True)

        date_from, date_to = parse_date_range("Main_SAP_ZMMR2199M_xl")

        txt_file_name = f"ZMMR2199M_{date_from}.txt"
        xlsx_file_name = f"ZMMR2199M_{date_from}.xlsx"

        txt_file_path = os.path.join(export_path, txt_file_name)
        xlsx_file_path = os.path.join(export_path, xlsx_file_name)

        if os.path.exists(txt_file_path):
            os.remove(txt_file_path)
        if os.path.exists(xlsx_file_path):
            os.remove(xlsx_file_path)

        os.system(f'''"C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe" -system=PR2 -client=320 -user={username} -pw={password}''')

        count = 0
        while not isinstance(session, win32com.client.CDispatch):
            time.sleep(1)
            count += 1
            if count > 60:
                status["ZMMR2199M"] = "⚠ Error: SAP GUI instance not found."
                status["Error Count"] += 1
                break
            try:
                SapGuiAuto = win32com.client.GetObject("SAPGUI")
                if not isinstance(SapGuiAuto, win32com.client.CDispatch):
                    continue
                application = SapGuiAuto.GetScriptingEngine
                if not isinstance(application, win32com.client.CDispatch):
                    SapGuiAuto = None
                    continue
                connection = application.Children(0)
                if not isinstance(connection, win32com.client.CDispatch):
                    application = None
                    SapGuiAuto = None
                    continue
                session = connection.Children(0)
                if not isinstance(session, win32com.client.CDispatch):
                    connection = None
                    application = None
                    SapGuiAuto = None
                    continue
            except Exception:
                continue

    else:
        status["ZMMR2199M"] = "⚠ Error: Missing SAP GUI credentials."
        status["Error Count"] += 1

    if status["ZMMR2199M"] == "null":
        try:
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
            session.findById("wnd[1]/usr/ctxtDY_PATH").text = export_path
            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = txt_file_name
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

            try:
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            except Exception:
                pass

            time.sleep(2)

            if not os.path.exists(txt_file_path) or os.path.getsize(txt_file_path) == 0:
                status["ZMMR2199M"] = "⚠ Error: Export TXT file was not created correctly."
                status["Error Count"] += 1
            else:
                # headers, rows = parse_zmmr2199m_txt(txt_file_path)
                # save_rows_to_xlsx(headers, rows, xlsx_file_path)

                headers, rows = parse_zmmr2199m_txt(txt_file_path)

                mtyp_index = [h.strip().replace(" ", "").lower() for h in headers].index("mtyp")
                rows = [row for row in rows if row[mtyp_index].strip().upper() == "HALB"]

                save_rows_to_xlsx(headers, rows, xlsx_file_path)

                if not os.path.exists(xlsx_file_path) or os.path.getsize(xlsx_file_path) == 0:
                    status["ZMMR2199M"] = "⚠ Error: Excel file was not created correctly."
                    status["Error Count"] += 1
                else:
                    try:
                        os.remove(txt_file_path)
                    except Exception:
                        pass
                    status["ZMMR2199M"] = "✓"

        except Exception as err:
            print(f"Unexpected {err=}, {type(err)=}")
            status["ZMMR2199M"] = "⚠ Error: VBScript execution failed."
            status["Error Count"] += 1

    if session is not None:
        try:
            session.findById("wnd[0]").close()
            try:
                session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            except Exception:
                pass
        except Exception:
            pass

    session = None
    connection = None
    application = None
    SapGuiAuto = None
