#!/usr/bin/env python3
"""
Generate test/sample-roadmap.xlsx matching the roadmap.html data structure,
without any third-party libraries (an .xlsx is just a zip of XML parts).

Dates are written as ISO strings (parsed by roadmap.html's parseExcelDate);
numbers as numeric cells; everything else as inline strings.
"""
import zipfile
from xml.sax.saxutils import escape

def col_letter(n):  # 0-based -> A, B, ...
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def cell_xml(col, row, value):
    ref = f"{col_letter(col)}{row}"
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}" t="n"><v>{value}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>'

def sheet_xml(rows):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>']
    for r, row in enumerate(rows, start=1):
        cells = "".join(cell_xml(c, r, v) for c, v in enumerate(row))
        out.append(f'<row r="{r}">{cells}</row>')
    out.append('</sheetData></worksheet>')
    return "".join(out)

# ---- Workbook data ----

loes_themes = [
    ["LOE ID", "Title"],
    ["L1", "Service Transformation"],
    ["L2", "Enabling Foundations"],
    [],
    ["LOE ID", "Theme ID", "Title", "Description"],
    ["L1", "T1", "Customer Service", "Improve the end-to-end customer experience"],
    ["L1", "T2", "Digital Channels", None],
    ["L2", "T3", "Data Platform", "Foundational data and analytics capability"],
    [],
]

activities = [
    ["Theme ID", "Activity ID", "UniqueAct ID", "Level", "Title", "Description", "Start Date", "End Date", "Resourcing"],
    ["T1", "CS",      "A-CS",    1, "Contact centre modernisation", "Replace legacy telephony", "2026-01-01", "2026-12-15", "Funded"],
    ["T1", "CS-IVR",  "A-CSIVR", 2, "New IVR rollout", "Self-service voice menus", "2026-02-01", "2026-06-30", "Funded"],
    ["T1", "CS-IVR-PILOT", "A-CSPILOT", 3, "IVR pilot in two regions", "", "2026-02-01", "2026-04-01", "At risk"],
    ["T1", "CS-CRM",  "A-CSCRM", 2, "CRM integration", "Link CRM to contact centre", "2026-05-01", "2026-11-30", "At risk"],
    ["T2", "DC",      "A-DC",    1, "Digital self-service portal", "", "2026-03-01", "2027-03-31", "Unfunded"],
    ["T2", "DC-AUTH", "A-DCAUTH", 2, "Single sign-on", "Federated identity", "2026-03-01", "2026-08-31", "Funded"],
    ["T3", "DP",      "A-DP",    1, "Enterprise data platform", "Cloud lakehouse", "2026-01-15", "2026-10-31", "Funded"],
    ["T3", "DP-ING",  "A-DPING", 2, "Source system ingestion", "", "2026-02-15", "2026-07-31", "Funded"],
]

milestones = [
    ["Activity ID", "UniqueMSt ID", "Milestone Title", "Description", "Date", "Owner", "Delivery Confidence", "Spare1"],
    ["A-CS",     "M-CS1",     "Vendor selected",        "Procurement complete", "2026-02-15", "A. Patel", "High",   "ignore"],
    ["A-CSIVR",  "M-IVR1",    "IVR live",               "",                     "2026-06-15", "J. Okoro", "Medium", "ignore"],
    ["A-CSPILOT","M-PILOT1",  "Pilot go/no-go",         "Decision gate",        "2026-03-20", "J. Okoro", "Low",    "ignore"],
    ["A-CSCRM",  "M-CRM1",    "CRM integrated",         "",                     "2026-11-15", "S. Khan",  "Medium", "ignore"],
    ["A-DC",     "M-DC1",     "Portal beta",            "Public beta release",  "2026-09-30", "R. Diaz",  "Low",    "ignore"],
    ["A-DCAUTH", "M-AUTH1",   "SSO enabled",            "",                     "2026-08-15", "R. Diaz",  "High",   "ignore"],
    ["A-DP",     "M-DP1",     "Platform GA",            "General availability", "2026-10-15", "L. Wong",  "High",   "ignore"],
    ["A-DPING",  "M-ING1",    "First 5 sources live",   "",                     "2026-07-15", "L. Wong",  "Medium", "ignore"],
]

benefits = [
    ["Milestone ID", "Benefit Title", "Category", "Beneficiary", "Impact"],
    ["M-IVR1", "Reduced call handling time", "Efficiency", "Contact centre agents", "High"],
    ["M-DP1",  "Single source of truth",     "Quality",    "Analysts",              "High"],
    ["M-AUTH1","Fewer password resets",      "Experience", "All staff",             "Medium"],
]

risks = [
    ["Milestone ID", "Risk Title", "Risk Owner", "RAG"],
    ["M-PILOT1", "Pilot regions lack capacity", "J. Okoro", "Red"],
    ["M-CRM1",   "Integration API unstable",    "S. Khan",  "Amber"],
    ["M-DC1",    "Funding not yet confirmed",   "R. Diaz",  "Amber"],
]

# Lookups table anchored at F1 (column index 5). A..E left blank.
PAD = [None] * 5
lookups = [
    PAD + ["Activities", "Milestones", "Benefits", "Risks"],
    PAD + ["Resourcing", "Delivery Confidence", "Category", "RAG"],
    PAD + ["green;amber;grey", "green;amber;red", "blue;purple;pink", "red;amber;green"],
    PAD + ["Funded", "High", "Efficiency", "Red"],
    PAD + ["At risk", "Medium", "Quality", "Amber"],
    PAD + ["Unfunded", "Low", "Experience", "Green"],
]

sheets = [
    ("LOEs & Themes", loes_themes),
    ("Activities", activities),
    ("Milestones", milestones),
    ("Benefits", benefits),
    ("Risks", risks),
    ("Lookups", lookups),
]

# ---- Package parts ----

content_types = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
for i in range(len(sheets)):
    content_types.append(f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
content_types.append('</Types>')

root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>')

wb_sheets = "".join(f'<sheet name="{escape(n)}" sheetId="{i+1}" r:id="rId{i+1}"/>' for i, (n, _) in enumerate(sheets))
workbook = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    f'<sheets>{wb_sheets}</sheets></workbook>')

wb_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
for i in range(len(sheets)):
    wb_rels.append(f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>')
wb_rels.append('</Relationships>')

with zipfile.ZipFile("test/sample-roadmap.xlsx", "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", "".join(content_types))
    z.writestr("_rels/.rels", root_rels)
    z.writestr("xl/workbook.xml", workbook)
    z.writestr("xl/_rels/workbook.xml.rels", "".join(wb_rels))
    for i, (_, rows) in enumerate(sheets):
        z.writestr(f"xl/worksheets/sheet{i+1}.xml", sheet_xml(rows))

print("wrote test/sample-roadmap.xlsx")
