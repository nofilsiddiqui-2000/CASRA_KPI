import pandas as pd
df = pd.read_excel(r"C:\Users\B1020000\Documents\Nofil\Dashboards\CASRA MM Dashboard\CASRA-KPI-AUTOMATION\SAP_Extracts\ZMMR2199M\ZMMR2199M_20260803.xlsx", dtype=str)
print(df["Rep. sel."].dropna().unique())