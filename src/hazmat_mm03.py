from datetime import datetime
import os, time, traceback
import win32com.client

from casra_common import read_config

# --- Hardcoded for validation phase ---
# Once the pipeline is confirmed working, this list will be replaced with the
# real part list (likely read from a file or another source).
MATERIALS = ["501847", "N381-1557"]
SALES_ORGS = ["3000", "4200", "1000"]
PLANT = "3099"
DIST_CHANNEL = "00"

LONGTEXT_PATH = (
    "wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/"
    "subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell"
)


def get_hazard_text(session):
    return session.findById(LONGTEXT_PATH).Text


def extract_material(session, material, results):
    """Look up a single material in MM03 and read the Sales Text for each sales org."""
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = material
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = len(material)
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08").select()

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
        # to avoid false-blank reads if the long-text control lags behind the rest of the screen.
        time.sleep(0.5)

        hazard_text = get_hazard_text(session)
        results.append((material, sales_org, hazard_text))

    # Back out to the initial material-entry screen for the next material
    session.findById("wnd[0]/tbar[0]/btn[3]").press()


def main():
    status = {
        "Meta": {
            "Start Time": datetime.today().strftime("%H:%M")
        },
        "Reporting Period": datetime.today().strftime("%b %d, %Y"),
        "Error Count": 0,
        "MM03_HAZMAT": "null"
    }

    session = None
    connection = None
    application = None
    SapGuiAuto = None
    results = []

    config = read_config()
    username = config.get("username", "")
    password = config.get("password", "")

    if username and password:
        os.system(
            f'''"C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe" '''
            f'''-system=PR2 -client=320 -user={username} -pw={password}'''
        )

        count = 0
        while not isinstance(session, win32com.client.CDispatch):
            time.sleep(1)
            count += 1
            if count > 60:
                status["MM03_HAZMAT"] = "⚠ Error: SAP GUI instance not found."
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
        status["MM03_HAZMAT"] = "⚠ Error: Missing SAP GUI credentials."
        status["Error Count"] += 1

    if status["MM03_HAZMAT"] == "null":
        try:
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").text = "MM03"
            session.findById("wnd[0]").sendVKey(0)

            for material in MATERIALS:
                extract_material(session, material, results)

            # NOTE: update this path to a shared drive / SharePoint location before production use,
            # same as the export_path note in the ZMNM script.
            export_path = r"C:\Users\B1020000\Documents\Nofil\Dashboards\CASRA MM Dashboard\CASRA-KPI-AUTOMATION\SAP_Extracts\MM03_HAZMAT"
            os.makedirs(export_path, exist_ok=True)

            csv_file_name = f"hazmat_results_{datetime.today().strftime('%Y%m%d')}.csv"
            csv_file_path = os.path.join(export_path, csv_file_name)

            if os.path.exists(csv_file_path):
                os.remove(csv_file_path)

            with open(csv_file_path, "w", encoding="utf-8") as f:
                f.write("MaterialNumber,SalesOrg,HazardText\n")
                for material, sales_org, hazard_text in results:
                    f.write(f'{material},{sales_org},"{hazard_text}"\n')

            if not os.path.exists(csv_file_path) or os.path.getsize(csv_file_path) == 0:
                status["MM03_HAZMAT"] = "⚠ Error: CSV file was not created correctly."
                status["Error Count"] += 1
            else:
                status["MM03_HAZMAT"] = "✓"

        except Exception as err:
            traceback.print_exc()
            status["MM03_HAZMAT"] = "⚠ Error: Script execution failed."
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

    return status


if __name__ == "__main__":
    result_status = main()
    print(result_status)
