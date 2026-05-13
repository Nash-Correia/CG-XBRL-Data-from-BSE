import os
import pandas as pd
import xml.etree.ElementTree as ET

# ─── CONFIG ────────────────────────────────────────────────────────────────
BASE_DIR    = r"d:\OneDrive - Institutional Investor Advisory Services India Limited\Desktop\Project\CG-XBRL-Data-from-BSE"
EXCEL_PATH  = os.path.join(BASE_DIR, "companies.xlsx")
YEARS       = [2025]
# ────────────────────────────────────────────────────────────────────────────

# fields to pull from the header of each XML  
header_fields = [
    "ScripCode", "Symbol", "MSEISymbol", "ISIN",
    "NameOfTheCompany", "DateOfStartOfFinancialYear",
    "DateOfEndOfFinancialYear", "DateOfEndOfReportingPeriod", "ReportingQuarter",
    "DateOfReport", "RiskManagementCommittee",
    "MarketCapitalisationAsPerImmediatePreviousFinancialYear",
    "WhetherTheListedEntityHasARegularChairperson",
    "WhetherChairpersonIsRelatedToMDOrCEO"
] 

# director-specific fields
director_fields = [
    "Title", "NameOftheDirector", "PermanentAccountNumberOfDirector",
    "DirectorIdentificationNumberOfDirector", "PositionOfDirectorInBoardOne",
    "PositionOfDirectorInBoardTwo", "PositionOfDirectorInBoardThree",
    "DateOfBirth", "WhetherTheDirectorIsDisqualified",
    "CurrentStatusDirector", "WhetherSpecialResolutionPassed",
    "DateOfAppointmentOfDirector", "DateOfReappointmentOfDirector",
    "TenureOfDirector", "NumberOfDirectorshipInListedEntitiesIncludingThisListedEntity",
    "NumberOfIndependentDirectorshipInListedEntitiesIncludingThisListedEntity",
    "NumberOfMembershipsInAuditOrStakeholderCommitteesIncludingThisListedEntity",
    "NumberOfPostOfChairpersonInAuditOrStakeholderCommitteeHeldInListedEntitiesIncludingThisListedEntity"
]

def local_name(tag: str) -> str:
    """Strip namespace, return local tag name."""
    return tag.split('}', 1)[-1] if '}' in tag else tag

def pick_first_existing_column(df: pd.DataFrame, candidates: list[str]) -> str:
    """Return the first matching column name from candidates (case-insensitive)."""
    lower_map = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in lower_map:
            return lower_map[key]
    raise KeyError(f"None of these columns were found in Excel: {candidates}. Available: {list(df.columns)}")

# load company list
df_companies = pd.read_excel(EXCEL_PATH, sheet_name="all")

# keep Company Name for output (optional but you asked mapping remains same)
name_col = pick_first_existing_column(df_companies, ["Name", "CompanyName", "Company Name"])

# NSE symbol column used for XML filename lookup
nse_col = pick_first_existing_column(df_companies, ["NSE", "NSE Symbol", "NSE_Symbol", "Symbol", "NSESymbol"])

# BSE code column
bse_col = pick_first_existing_column(df_companies, ["BSE", "BSE Code", "BSE_Code", "BSECode"])

for year in YEARS:
    XML_DIR     = os.path.join(BASE_DIR, f"XBRL_{year}")
    OUTPUT_PATH = os.path.join(BASE_DIR, f"Directors Data BSE_{year}.xlsx")

    print(f"\n── Processing year {year} ──")

    rows = []
    missing_rows = []

    for _, r in df_companies.iterrows():
        company_name = str(r.get(name_col, "")).strip()
        nse_symbol = str(r.get(nse_col, "")).strip().replace("&", "_")
        bse_code = str(r.get(bse_col, "")).strip()

        if not nse_symbol or nse_symbol.lower() == "nan":
            print(f"⚠ NSE symbol missing for '{company_name}', skipping.")
            continue

        # Your files are saved as NSE_SYMBOL.xml
        xml_path = os.path.join(XML_DIR, f"{nse_symbol}.xml")

        if not os.path.isfile(xml_path):
            # small fallback: try uppercase if files are stored as uppercase
            xml_path2 = os.path.join(XML_DIR, f"{nse_symbol.upper()}.xml")
            if os.path.isfile(xml_path2):
                xml_path = xml_path2
            else:
                print(f"⚠ XML not found for NSE '{nse_symbol}' (Company: {company_name}), skipping.")
                missing_rows.append({"Name": company_name, "BSE": bse_code, "NSE": nse_symbol})
                continue

        tree = ET.parse(xml_path)
        root = tree.getroot()

        # pull header data
        header_data = {"CompanyName": company_name, "NSESymbol": nse_symbol}
        for elem in root.iter():
            tag = local_name(elem.tag)
            if tag in header_fields:
                header_data[tag] = (elem.text or "").strip()

        # collect directors by normalized contextRef (handles both D_CompBODx & CompBODx)
        directors = {}
        for elem in root.iter():
            tag = local_name(elem.tag)
            ctx = elem.attrib.get("contextRef", "")
            if tag in director_fields and (ctx.startswith("D_CompBOD") or ctx.startswith("CompBOD")):
                # normalize key so D_CompBOD1 and CompBOD1 both map to 'CompBOD1'
                key = ctx[2:] if ctx.startswith("D_") else ctx
                directors.setdefault(key, {})[tag] = (elem.text or "").strip()

        # append a row per director
        for _, data in directors.items():
            row = dict(header_data)  # copy header fields
            # fill director fields (missing ones become blank)
            for fld in director_fields:
                row[fld] = data.get(fld, "")
            rows.append(row)

    # build DataFrame and write to Excel
    df_out = pd.DataFrame(rows) if rows else pd.DataFrame()
    df_missing = pd.DataFrame(missing_rows, columns=["Name", "BSE", "NSE"]) if missing_rows else pd.DataFrame(columns=["Name", "BSE", "NSE"])

    if not df_out.empty:
        # set column order: CompanyName + NSESymbol + header_fields + director_fields
        cols = ["CompanyName", "NSESymbol"] + header_fields + director_fields
        cols = [c for c in cols if c in df_out.columns]
        df_out = df_out.loc[:, cols]

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        if not df_out.empty:
            df_out.to_excel(writer, sheet_name="Director Data", index=False)
        else:
            pd.DataFrame().to_excel(writer, sheet_name="Director Data", index=False)
        df_missing.to_excel(writer, sheet_name="Missing XML", index=False)

    print(f"✅ Directors data written to: {OUTPUT_PATH}")
    if missing_rows:
        print(f"⚠ {len(missing_rows)} companies had missing XMLs — see 'Missing XML' sheet.")
