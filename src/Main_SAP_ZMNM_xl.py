from datetime import datetime, timedelta
import os, time
import win32com.client

from casra_config import read_config
from casra_dates import parse_date_range

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
        "ZMNM": "null"
    }

    session = None
    connection = None
    application = None
    SapGuiAuto = None

    username = config.get("username", "")
    password = config.get("password", "")

    if username and password:
        # export_path = r"C:\Users\B1020000\Bombardier\Team CM Aftermarket - SAP Reports\ZMNM"
        export_path = r"C:\Users\B1020000\Documents\Nofil\Dashboards\CASRA MM Dashboard\CASRA-KPI-AUTOMATION\SAP_Extracts\ZMNM"      # this needs to  be the path to save the extracts to, and the next script will need to be updated to pull from this path as well. I have it set to a folder in my documents for testing, but in production it should be a shared drive or SharePoint folder that the next script can access.
        os.makedirs(export_path, exist_ok=True)

        date_from, date_to = parse_date_range("Main_SAP_ZMNM_xl")

        txt_file_name = f"ZMNM_{date_from}.txt"
        xlsx_file_name = f"ZMNM_{date_from}.xlsx"

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
                status["ZMNM"] = "⚠ Error: SAP GUI instance not found."
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
        status["ZMNM"] = "⚠ Error: Missing SAP GUI credentials."
        status["Error Count"] += 1

    if status["ZMNM"] == "null":
        try:
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "ZMNM"
            session.findById("wnd[0]").sendVKey(0)

            # low date
            session.findById("wnd[0]/usr/ctxtSO_ERSDA-LOW").setFocus()
            session.findById("wnd[0]/usr/ctxtSO_ERSDA-LOW").caretPosition = 0
            session.findById("wnd[0]").sendVKey(4)
            session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").focusDate = date_from
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

            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").columns.elementAt(1).width = 12
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").text = "B0516399"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").text = "B0466709"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,2]").text = "B0536466"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,3]").text = "B0460021"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,4]").text = "B0534548"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,5]").text = "B0533127"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,6]").text = "B0530939"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,7]").text = "B0075972"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,7]").setFocus()
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,7]").caretPosition = 9
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.position = 7
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").text = "B0737251"
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").setFocus()
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").caretPosition = 9
            session.findById("wnd[1]").sendVKey(0)
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            session.findById("wnd[1]/tbar[0]/btn[8]").press()

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
                status["ZMNM"] = "⚠ Error: Export TXT file was not created correctly."
                status["Error Count"] += 1
            else:
                convert_txt_to_xlsx(txt_file_path, xlsx_file_path)

                if not os.path.exists(xlsx_file_path) or os.path.getsize(xlsx_file_path) == 0:
                    status["ZMNM"] = "⚠ Error: Excel file was not created correctly."
                    status["Error Count"] += 1
                else:
                    try:
                        os.remove(txt_file_path)
                    except Exception:
                        pass
                    status["ZMNM"] = "✓"

        except Exception as err:
            print(f"Unexpected {err=}, {type(err)=}")
            status["ZMNM"] = "⚠ Error: VBScript execution failed."
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
