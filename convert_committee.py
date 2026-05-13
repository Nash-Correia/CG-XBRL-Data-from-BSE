import os
import pandas as pd
import xml.etree.ElementTree as ET

# ─── CONFIG ────────────────────────────────────────────────────────────────
EXCEL_PATH  = r"D:\OneDrive - Institutional Investor Advisory Services India Limited\Desktop\projects\board data from xbrl\companies.xlsx"
XML_DIR     = r"D:\OneDrive - Institutional Investor Advisory Services India Limited\Desktop\projects\board data from xbrl\XBRL Files"
OUTPUT_PATH = r"D:\OneDrive - Institutional Investor Advisory Services India Limited\Desktop\projects\board data from xbrl\Committee Composition.xlsx"
# ────────────────────────────────────────────────────────────────────────────

# Header fields to carry into every committee row
HEADER_FIELDS = [
    "ScripCode", "Symbol", "MSEISymbol", "ISIN",
    "NameOfTheCompany", "DateOfStartOfFinancialYear",
    "DateOfEndOfFinancialYear", "DateOfEndOfReportingPeriod",
    "ReportingQuarter", "DateOfReport",
]

# Per-member fields keyed by contextRef type
#   D_CompComit{N}  -> dimension-qualified (most fields)
#   CompComit{N}    -> plain (dates without dimensions)
D_FIELDS = [
    "NameOfCommittee",
    "DirectorIdentificationNumberOfDirector",
    "NameOfCommitteeMembers",
    "PositionOfDirectorInCommitteeOne",
    "PositionOfDirectorInCommitteeTwo",
]
PLAIN_FIELDS = [
    "DateOfAppointmentOfDirectorInCommittee",
    "DateOfCessationOfDirectorInCommittee",
]

# Map normalised committee name keywords -> Excel sheet name (max 31 chars)
COMMITTEE_SHEET_MAP = [
    ("audit",            "Audit Committee"),
    ("nomination",       "Nomination Remun Committee"),
    ("stakeholder",      "Stakeholders Rel Committee"),
    ("risk management",  "Risk Management Committee"),
    ("corporate social", "CSR Committee"),
]

ALL_SHEET_NAMES = [sheet for _, sheet in COMMITTEE_SHEET_MAP]


def local_name(tag: str) -> str:
    """Strip XML namespace, return local tag name."""
    return tag.split('}', 1)[-1] if '}' in tag else tag


def pick_column(df: pd.DataFrame, candidates: list) -> str:
    """Return the first matching column (case-insensitive)."""
    lower_map = {c.strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in lower_map:
            return lower_map[key]
    raise KeyError(f"None of {candidates} found in Excel. Available: {list(df.columns)}")


def committee_sheet(name: str) -> str:
    """Map a raw committee name from XML to an Excel sheet name."""
    lower = name.lower()
    for keyword, sheet in COMMITTEE_SHEET_MAP:
        if keyword in lower:
            return sheet
    return None   # skip committees not in the target list


def parse_xml(xml_path: str, company_name: str, nse_symbol: str) -> list:
    """
    Parse one XML file and return a list of committee-member row dicts.
    Each dict contains header metadata + committee member fields.
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # ── 1. Collect header data ────────────────────────────────────────────
    header_data = {"CompanyName": company_name, "NSESymbol": nse_symbol}
    for elem in root.iter():
        tag = local_name(elem.tag)
        if tag in HEADER_FIELDS:
            header_data.setdefault(tag, (elem.text or "").strip())

    # ── 2. Collect committee-member data keyed by index ───────────────────
    # We normalise D_CompComit7 and CompComit7 both to index "7"
    members: dict = {}   # index -> {"_d_fields": {...}, "_plain_fields": {...}}

    for elem in root.iter():
        tag = local_name(elem.tag)
        ctx = elem.attrib.get("contextRef", "")
        val = (elem.text or "").strip()

        if ctx.startswith("D_CompComit") and tag in D_FIELDS:
            idx = ctx[len("D_CompComit"):]
            members.setdefault(idx, {})
            members[idx][tag] = val

        elif ctx.startswith("CompComit") and tag in PLAIN_FIELDS:
            # plain CompComit{N}  (no D_ prefix)
            idx = ctx[len("CompComit"):]
            members.setdefault(idx, {})
            members[idx][tag] = val

    # ── 3. Build rows ─────────────────────────────────────────────────────
    rows = []
    for data in members.values():
        raw_name = data.get("NameOfCommittee", "")
        sheet = committee_sheet(raw_name)
        if sheet is None:
            continue   # not one of the 5 target committees

        row = dict(header_data)
        row["CommitteeSheet"] = sheet   # internal routing key
        for fld in D_FIELDS + PLAIN_FIELDS:
            default = "-" if fld == "DateOfCessationOfDirectorInCommittee" else ""
            row[fld] = data.get(fld, default) or default
        rows.append(row)

    return rows


# ─── Main ─────────────────────────────────────────────────────────────────
df_companies = pd.read_excel(EXCEL_PATH)

name_col = pick_column(df_companies, ["Name", "CompanyName", "Company Name"])
nse_col  = pick_column(df_companies, ["NSE", "NSE Symbol", "NSE_Symbol", "Symbol", "NSESymbol"])
bse_col  = pick_column(df_companies, ["BSE", "BSE Code", "BSE_Code", "BSECode"])

all_rows = []
missing_rows = []

for _, r in df_companies.iterrows():
    company_name = str(r.get(name_col, "")).strip()
    nse_symbol   = str(r.get(nse_col,  "")).strip().replace("&", "_")
    bse_code     = str(r.get(bse_col,  "")).strip()

    if not nse_symbol or nse_symbol.lower() == "nan":
        print(f"[WARN] NSE symbol missing for '{company_name}', skipping.")
        continue

    xml_path = os.path.join(XML_DIR, f"{nse_symbol}.xml")
    if not os.path.isfile(xml_path):
        xml_path2 = os.path.join(XML_DIR, f"{nse_symbol.upper()}.xml")
        if os.path.isfile(xml_path2):
            xml_path = xml_path2
        else:
            print(f"[WARN] XML not found for '{nse_symbol}' ({company_name}), skipping.")
            missing_rows.append({"Name": company_name, "BSE": bse_code, "NSE": nse_symbol})
            continue

    rows = parse_xml(xml_path, company_name, nse_symbol)
    all_rows.extend(rows)

# ─── Split into per-committee DataFrames ──────────────────────────────────
sheet_dfs = {sheet: [] for sheet in ALL_SHEET_NAMES}

for row in all_rows:
    sheet_name = row.pop("CommitteeSheet")
    sheet_dfs[sheet_name].append(row)

# Column order for output
base_cols = ["CompanyName", "NSESymbol"] + HEADER_FIELDS + D_FIELDS + PLAIN_FIELDS

# ─── Write Excel ──────────────────────────────────────────────────────────
df_missing = pd.DataFrame(missing_rows, columns=["Name", "BSE", "NSE"]) if missing_rows else pd.DataFrame(columns=["Name", "BSE", "NSE"])

with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
    for sheet_name in ALL_SHEET_NAMES:
        rows_for_sheet = sheet_dfs[sheet_name]
        if rows_for_sheet:
            df = pd.DataFrame(rows_for_sheet)
            cols = [c for c in base_cols if c in df.columns]
            df = df.loc[:, cols]
        else:
            df = pd.DataFrame()
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    df_missing.to_excel(writer, sheet_name="Missing XML", index=False)

print(f"\nDone. Committee composition written to: {OUTPUT_PATH}")
if missing_rows:
    print(f"[WARN] {len(missing_rows)} companies had missing XMLs -- see 'Missing XML' sheet.")
