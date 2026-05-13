import os
from pathlib import Path
import xml.etree.ElementTree as ET
import pandas as pd

EXCEL_PATH  = r"D:\OneDrive - Institutional Investor Advisory Services India Limited\Desktop\projects\board table database\companies.xlsx"
XML_DIR     = r"D:\OneDrive - Institutional Investor Advisory Services India Limited\Desktop\projects\board table database\XBRL Files"
OUTPUT_PATH = r"D:\OneDrive - Institutional Investor Advisory Services India Limited\Desktop\projects\board table database\Directors Data BSE.xlsx"

# -----------------------------
# Schemas
# -----------------------------
header_fields = [
    "ScripCode", "Symbol", "MSEISymbol", "ISIN",
    "NameOfTheCompany", "DateOfStartOfFinancialYear",
    "DateOfEndOfFinancialYear", "ReportingQuarter",
    "DateOfReport", "RiskManagementCommittee",
    "MarketCapitalisationAsPerImmediatePreviousFinancialYear",
    "WhetherTheListedEntityHasARegularChairperson",
    "WhetherChairpersonIsRelatedToMDOrCEO"
]

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

committee_fields = [
    "Committee",
    "DIN",
    "Member Name",
    "Role (Type)",
    "Position",
    "Appointment Date",
    "Cessation Date",
    "ContextRef",
]

# -----------------------------
# Helpers
# -----------------------------
def clean_symbol(sym: str) -> str:
    if sym is None:
        return ""
    sym = str(sym).strip()
    if sym.lower() in {"nan", "none", "not found", ""}:
        return ""
    return sym.upper()

def enforce_columns(df: pd.DataFrame, cols: list) -> pd.DataFrame:
    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols]

def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag

def normalize_director_ctx(ctx: str) -> str | None:
    """
    Normalize director context refs so that:
      - D_CompBODx stays D_CompBODx
      - CompBODx becomes D_CompBODx (so both halves merge)
    Return None for anything that isn't board-director context.
    """
    if not ctx:
        return None
    if ctx.startswith("D_CompBOD"):
        return ctx
    if ctx.startswith("CompBOD"):
        return "D_" + ctx
    return None

def filter_director_records(df: pd.DataFrame) -> pd.DataFrame:
    """
    Keep only rows that represent a real director.
    A 'real' director has at least one strong identity field.
    """
    df = df.copy()
    name = df.get("NameOftheDirector", pd.Series([""] * len(df))).astype(str).str.strip().replace({"nan": ""})
    din  = df.get("DirectorIdentificationNumberOfDirector", pd.Series([""] * len(df))).astype(str).str.strip().replace({"nan": ""})
    keep = (name.ne("")) | (din.ne(""))
    return df.loc[keep].reset_index(drop=True)

def filter_committee_records(df: pd.DataFrame) -> pd.DataFrame:
    """
    Keep only real committee member rows:
      - Member Name present
      - DIN present
      - AND at least one of: Committee / Role / Appointment Date / Cessation Date present
    (Position can be blank; that's fine.)
    """
    df = df.copy()
    member = df.get("Member Name", pd.Series([""] * len(df))).astype(str).str.strip().replace({"nan": ""})
    din    = df.get("DIN", pd.Series([""] * len(df))).astype(str).str.strip().replace({"nan": ""})

    identity_ok = member.ne("") & din.ne("")

    detail_cols = ["Committee", "Role (Type)", "Appointment Date", "Cessation Date"]
    details_ok = False
    for c in detail_cols:
        if c in df.columns:
            details_ok = details_ok | df[c].astype(str).str.strip().replace({"nan": ""}).ne("")

    keep = identity_ok & details_ok
    return df.loc[keep].reset_index(drop=True)

# -----------------------------
# Parsing
# -----------------------------
def parse_xbrl_facts(xml_path: str):
    """
    Parse an XBRL XML and return:
      - header_row: dict (single row)
      - directors_df: one row per director (D_CompBODx + CompBODx merged)
      - committees_df: one row per committee member (filtered for sanity)
    """
    tree = ET.parse(xml_path)
    root = tree.getroot()

    # Collect facts once
    facts = []
    for el in root.iter():
        if not el.text:
            continue
        txt = el.text.strip()
        if not txt:
            continue
        ln = local_name(el.tag)
        ctx = el.attrib.get("contextRef", "")
        facts.append((ln, ctx, txt))

    # 1) Header row (first occurrence wins)
    header_row = {f: "" for f in header_fields}
    for ln, ctx, txt in facts:
        if ln in header_row and header_row[ln] == "":
            header_row[ln] = txt

    # 2) Directors: merge CompBODx -> D_CompBODx
    director_by_ctx = {}
    for ln, ctx, txt in facts:
        if ln not in director_fields:
            continue
        key = normalize_director_ctx(ctx)
        if key is None:
            continue
        if key not in director_by_ctx:
            director_by_ctx[key] = {"ContextRef": key}
        director_by_ctx[key][ln] = txt

    directors_df = pd.DataFrame(list(director_by_ctx.values())) if director_by_ctx else pd.DataFrame()
    directors_df = enforce_columns(directors_df, ["ContextRef"] + director_fields)
    directors_df = filter_director_records(directors_df)

    # 3) Committees
    committee_tag_map = {
        "NameOfCommittee": "Committee",
        "DirectorIdentificationNumberOfDirector": "DIN",
        "NameOfCommitteeMembers": "Member Name",
        "PositionOfDirectorInCommitteeOne": "Role (Type)",
        "PositionOfDirectorInCommitteeTwo": "Position",
        "DateOfAppointmentOfDirectorInCommittee": "Appointment Date",
        "DateOfCessationOfDirectorInCommittee": "Cessation Date",
    }

    committee_by_ctx = {}
    for ln, ctx, txt in facts:
        if ln not in committee_tag_map:
            continue
        if not ctx:
            continue

        key = ctx
        # normalize CompComitXX -> D_CompComitXX
        if ctx.startswith("CompComit"):
            key = "D_" + ctx

        if key not in committee_by_ctx:
            committee_by_ctx[key] = {"ContextRef": key}

        committee_by_ctx[key][committee_tag_map[ln]] = txt

    committees_df = pd.DataFrame(list(committee_by_ctx.values())) if committee_by_ctx else pd.DataFrame()
    committees_df = enforce_columns(committees_df, committee_fields)
    committees_df = filter_committee_records(committees_df)

    return header_row, directors_df, committees_df

def split_by_committee(committees_df: pd.DataFrame):
    df = committees_df.copy()
    norm = df["Committee"].astype(str).str.strip().str.lower()

    def pick(keywords):
        mask = False
        for k in keywords:
            mask = mask | norm.str.contains(k, na=False)
        return df.loc[mask].reset_index(drop=True)

    ac  = pick(["audit committee"])
    nrc = pick(["nomination", "remuneration"])
    csr = pick(["corporate social responsibility", "csr"])
    rmc = pick(["risk management"])
    src = pick(["stakeholders", "relationship"])

    return ac, nrc, csr, rmc, src

# -----------------------------
# Pipeline
# -----------------------------
def run_pipeline():
    companies = pd.read_excel(EXCEL_PATH)
    if "NSE" not in companies.columns:
        raise KeyError(f"'NSE' column not found in {EXCEL_PATH}. Found: {list(companies.columns)}")

    nse_list = companies["NSE"].map(clean_symbol).dropna().unique().tolist()
    nse_list = [s for s in nse_list if s]

    header_rows = []
    all_directors = []
    all_committees = []
    missing = []

    for sym in nse_list:
        xml_path = str(Path(XML_DIR) / f"{sym}.xml")

        if not os.path.exists(xml_path):
            missing.append({"NSE": sym, "XML Path Tried": xml_path})
            continue

        header_row, directors_df, committees_df = parse_xbrl_facts(xml_path)

        header_row["Symbol"] = header_row.get("Symbol") or sym
        header_row["_XML_Path"] = xml_path
        header_rows.append(header_row)

        id_cols = {
            "Symbol": header_row.get("Symbol", sym),
            "ISIN": header_row.get("ISIN", ""),
            "ScripCode": header_row.get("ScripCode", ""),
            "NameOfTheCompany": header_row.get("NameOfTheCompany", ""),
            "_XML_Path": xml_path,
        }

        if len(directors_df) > 0:
            for k, v in id_cols.items():
                directors_df[k] = v
            all_directors.append(directors_df)

        if len(committees_df) > 0:
            for k, v in id_cols.items():
                committees_df[k] = v
            all_committees.append(committees_df)

    header_df = pd.DataFrame(header_rows)
    header_df = enforce_columns(header_df, header_fields + ["_XML_Path"]) if len(header_df) else pd.DataFrame(columns=header_fields + ["_XML_Path"])

    directors_all_df = pd.concat(all_directors, ignore_index=True) if all_directors else pd.DataFrame()
    committees_all_df = pd.concat(all_committees, ignore_index=True) if all_committees else pd.DataFrame()

    director_sheet_cols = ["Symbol", "ISIN", "ScripCode", "NameOfTheCompany", "_XML_Path", "ContextRef"] + director_fields
    directors_all_df = enforce_columns(directors_all_df, director_sheet_cols) if len(directors_all_df) else pd.DataFrame(columns=director_sheet_cols)

    committee_sheet_cols = ["Symbol", "ISIN", "ScripCode", "NameOfTheCompany", "_XML_Path"] + committee_fields
    committees_all_df = enforce_columns(committees_all_df, committee_sheet_cols) if len(committees_all_df) else pd.DataFrame(columns=committee_sheet_cols)

    ac, nrc, csr, rmc, src = split_by_committee(committees_all_df) if len(committees_all_df) else (
        pd.DataFrame(columns=committee_sheet_cols),
        pd.DataFrame(columns=committee_sheet_cols),
        pd.DataFrame(columns=committee_sheet_cols),
        pd.DataFrame(columns=committee_sheet_cols),
        pd.DataFrame(columns=committee_sheet_cols),
    )

    missing_df = pd.DataFrame(missing)

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        header_df.to_excel(writer, sheet_name="Header Data", index=False)
        directors_all_df.to_excel(writer, sheet_name="Board Data", index=False)

        ac.to_excel(writer,  sheet_name="AC",  index=False)
        nrc.to_excel(writer, sheet_name="NRC", index=False)
        csr.to_excel(writer, sheet_name="CSR", index=False)
        rmc.to_excel(writer, sheet_name="RMC", index=False)
        src.to_excel(writer, sheet_name="SRC", index=False)

        missing_df.to_excel(writer, sheet_name="XML Missing", index=False)

    print(f"✅ Done. Output written to: {OUTPUT_PATH}")

if __name__ == "__main__":
    run_pipeline()
