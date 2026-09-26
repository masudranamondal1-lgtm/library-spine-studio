"""
Personal Library Manager & Digital Archive — Workbook Generator
Produces: Personal Library Manager & Digital Archive.xlsx
Requires: openpyxl
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────
# STYLE CONSTANTS
# ─────────────────────────────────────────────────────────────────
HEADER_BG   = "1F3864"
HEADER_FG   = "FFFFFF"
SECTION_BG  = "2E5C8A"
INPUT_BG    = "FFF9E6"   # pale cream — editable
GEN_BG      = "F2F2F2"   # grey — generated / formula
ERROR_BG    = "F8D7DA"
WARN_BG     = "FFF3CD"
OK_BG       = "D4EDDA"
BORDER_CLR  = "BFBFBF"

FONT = "Calibri"
BENGALI_FONT = "Nirmala UI"  # falls back gracefully

thin = Side(style="thin", color=BORDER_CLR)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

def hdr(cell):
    cell.font = Font(name=FONT, bold=True, color=HEADER_FG, size=11)
    cell.fill = PatternFill("solid", fgColor=HEADER_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER

def input_cell(cell):
    cell.fill = PatternFill("solid", fgColor=INPUT_BG)
    cell.protection = Protection(locked=False)
    cell.border = BORDER
    cell.alignment = Alignment(vertical="center", wrap_text=True)

def gen_cell(cell):
    cell.fill = PatternFill("solid", fgColor=GEN_BG)
    cell.protection = Protection(locked=True)
    cell.border = BORDER
    cell.alignment = Alignment(vertical="center", wrap_text=True)

def section(cell):
    cell.font = Font(name=FONT, bold=True, color="FFFFFF", size=12)
    cell.fill = PatternFill("solid", fgColor=SECTION_BG)
    cell.alignment = Alignment(horizontal="left", vertical="center")

# ─────────────────────────────────────────────────────────────────
# WORKBOOK
# ─────────────────────────────────────────────────────────────────
wb = Workbook()
wb.remove(wb.active)

ROWS_CAT = 500
ROWS_DIG = 500
ROWS_ARC = 200
ROWS_CIR = 500
ROWS_MEM = 100
ROWS_SPN = 500

# ═════════════════════════════════════════════════════════════════
# 1. SETTINGS  (created first so other sheets can reference it)
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Settings")
ws.column_dimensions["A"].width = 32
ws.column_dimensions["B"].width = 40
ws.column_dimensions["C"].width = 50

ws["A1"] = "SETTINGS — Personal Library Manager"
ws["A1"].font = Font(name=FONT, bold=True, size=14, color=HEADER_BG)
ws.merge_cells("A1:C1")

ws["A3"] = "Setting"; ws["B3"] = "Value"; ws["C3"] = "Purpose"
for c in ("A3","B3","C3"): hdr(ws[c])

settings_data = [
    ("Library Name",                "Personal Library",       "Display name of the library"),
    ("Collection Name",             "Main Collection",        "Name of the primary collection"),
    ("Default Language",            "English",                "Default language for new records"),
    ("Default Loan Period (days)",  14,                       "Circulation loan period in days"),
    ("Accession Prefix",            "ACC",                    "Prefix for accession numbers"),
    ("Author Code Length",          3,                        "Number of letters for auto author code"),
    ("Date Format",                 "YYYY-MM-DD",             "Display format for dates"),
    ("Default Resource Type",       "Book",                   "Default resource type"),
    ("Metadata Profile Version",    "2.0",                    "Dublin Core profile version"),
    ("Workbook Protection Password","library",                "Password to unprotect sheets"),
]
for i, (k, v, p) in enumerate(settings_data, start=4):
    ws.cell(i, 1, k).font = Font(name=FONT, bold=True)
    ws.cell(i, 2, v).fill = PatternFill("solid", fgColor=INPUT_BG)
    ws.cell(i, 2).protection = Protection(locked=False)
    ws.cell(i, 3, p).font = Font(name=FONT, italic=True, size=9)

# Named references
wb.defined_names.add(DefinedName("LibraryName",    attr_text="Settings!$B$4"))
wb.defined_names.add(DefinedName("CollectionName", attr_text="Settings!$B$5"))
wb.defined_names.add(DefinedName("DefaultLang",    attr_text="Settings!$B$6"))
wb.defined_names.add(DefinedName("LoanPeriod",     attr_text="Settings!$B$7"))
wb.defined_names.add(DefinedName("AccPrefix",      attr_text="Settings!$B$8"))
wb.defined_names.add(DefinedName("AuthorCodeLen",  attr_text="Settings!$B$9"))

# ═════════════════════════════════════════════════════════════════
# 2. LOOKUP_CODES
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Lookup_Codes")
vocab = {
    "A": ("Resource Type", [
        "Book","Magazine","Journal","Periodical","Newsletter","Annual","Special Issue",
        "Article","Image","Audio Recording","Video Recording","Dataset","Website",
        "Manuscript","Newspaper","Newspaper Clipping","Photograph","Poster","Map",
        "Pamphlet","Brochure","Personal Papers","Ephemera","Other"]),
    "B": ("Genre", [
        "Literature (LIT)","History (HIS)","Philosophy (PHI)","Science & Tech (SCI)",
        "Politics (POL)","Essays (ESS)","Biography / Memoir (BIO)",
        "Mystery / Thriller (THR)","Religion & Mythology (REL)","Art & Cinema (ART)","Other"]),
    "C": ("Language", ["English","Bengali","Hindi","Other"]),
    "D": ("Format", [
        "PDF","EPUB","MOBI","DJVU","TXT","DOCX","HTML","XML","TEI XML","Markdown","RTF",
        "JPG","JPEG","PNG","TIFF","WebP","MP3","WAV","FLAC","M4A",
        "MP4","MKV","MOV","WebM","ZIP","Other"]),
    "E": ("OCR Status", [
        "Not Applicable","Not Digitised","Scanned (No OCR)","OCR Pending",
        "OCR Complete","Searchable (OCR)","Image Only"]),
    "F": ("Condition", ["New","Excellent","Good","Fair","Poor","Damaged","Under Repair"]),
    "G": ("Circulation Status", ["Available","On Loan","Reference Only","Lost","Withdrawn"]),
    "H": ("Location Type", ["Shelf","Magazine Rack","Journal Rack","Archive","Cabinet","Box","Other"]),
    "I": ("Access Level", ["Public","Private","Restricted","Research Only","Permission Required"]),
    "J": ("Rights", ["Copyright","Public Domain","Licensed","Unknown","User-Owned","Institutional","Other"]),
    "K": ("Acquisition Type", ["Purchase","Gift","Donation","Exchange","Legacy","Unknown"]),
    "L": ("Frequency", ["Daily","Weekly","Fortnightly","Monthly","Bi-Monthly","Quarterly","Annual","Irregular"]),
    "M": ("Digitisation Status", ["Not Digitised","Digitisation Planned","Scanned","OCR Pending","OCR Complete","OCR Verified"]),
    "N": ("Physical/Digital", ["Physical","Digital","Both"]),
    "O": ("Membership Type", ["Student","Faculty","Public","Institutional","Other"]),
    "P": ("Member Status", ["Active","Inactive","Suspended"]),
    "Q": ("Reading/Usage Status", ["Read","Reading","Unread","N/A"]),
    "R": ("Preservation Status", ["Master","Derivative","Archived","At Risk"]),
}
for col, (title, values) in vocab.items():
    ws.column_dimensions[col].width = max(20, len(title) + 6)
    c = ws[f"{col}1"]; c.value = title; hdr(c)
    for i, v in enumerate(values, start=2):
        cell = ws[f"{col}{i}"]; cell.value = v
        cell.font = Font(name=FONT, size=10)

# Named ranges for validation
wb.defined_names.add(DefinedName("ResourceTypeList",   attr_text="Lookup_Codes!$A$2:$A$30"))
wb.defined_names.add(DefinedName("GenreList",          attr_text="Lookup_Codes!$B$2:$B$15"))
wb.defined_names.add(DefinedName("LanguageList",       attr_text="Lookup_Codes!$C$2:$C$10"))
wb.defined_names.add(DefinedName("FormatList",         attr_text="Lookup_Codes!$D$2:$D$30"))
wb.defined_names.add(DefinedName("OCRStatusList",      attr_text="Lookup_Codes!$E$2:$E$10"))
wb.defined_names.add(DefinedName("ConditionList",      attr_text="Lookup_Codes!$F$2:$F$10"))
wb.defined_names.add(DefinedName("CirculationList",    attr_text="Lookup_Codes!$G$2:$G$10"))
wb.defined_names.add(DefinedName("LocationTypeList",   attr_text="Lookup_Codes!$H$2:$H$10"))
wb.defined_names.add(DefinedName("AccessLevelList",    attr_text="Lookup_Codes!$I$2:$I$10"))
wb.defined_names.add(DefinedName("RightsList",         attr_text="Lookup_Codes!$J$2:$J$10"))
wb.defined_names.add(DefinedName("AcqTypeList",        attr_text="Lookup_Codes!$K$2:$K$10"))
wb.defined_names.add(DefinedName("FrequencyList",      attr_text="Lookup_Codes!$L$2:$L$12"))
wb.defined_names.add(DefinedName("DigitisationList",   attr_text="Lookup_Codes!$M$2:$M$10"))
wb.defined_names.add(DefinedName("PhysDigiList",       attr_text="Lookup_Codes!$N$2:$N$5"))
wb.defined_names.add(DefinedName("MembershipList",     attr_text="Lookup_Codes!$O$2:$O$10"))
wb.defined_names.add(DefinedName("MemberStatusList",   attr_text="Lookup_Codes!$P$2:$P$6"))
wb.defined_names.add(DefinedName("ReadingStatusList",  attr_text="Lookup_Codes!$Q$2:$Q$6"))
wb.defined_names.add(DefinedName("PreservationList",   attr_text="Lookup_Codes!$R$2:$R$6"))

# ═════════════════════════════════════════════════════════════════
# 3. CATALOG
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Catalog")
cat_cols = [
    ("Item ID", 8), ("Accession Number", 16), ("Call Number", 26),
    ("Title", 32), ("Subtitle", 22), ("Creator", 26),
    ("Author Code Override", 12), ("Author Code", 12),
    ("Editor", 20), ("Translator", 20),
    ("Type", 16), ("Genre", 18), ("Language", 12),
    ("Volume", 10), ("Edition", 12), ("Publisher", 22),
    ("Publication Place", 16), ("Publication Year", 8),
    ("ISBN", 16), ("ISSN", 14), ("Pages", 8), ("Series", 20),
    ("Location Type", 16), ("Location Identifier", 16),
    ("Condition", 12), ("Circulation Status", 14),
    ("Acquisition Date", 14), ("Acquisition Type", 14),
    ("Price", 10), ("Donor/Vendor", 20),
    ("Notes", 30), ("Rights", 16), ("Description", 30),
]
for i, (name, w) in enumerate(cat_cols, start=1):
    col = get_column_letter(i)
    ws.column_dimensions[col].width = w
    c = ws.cell(1, i, name); hdr(c)
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(cat_cols))}{ROWS_CAT+1}"

for r in range(2, ROWS_CAT + 2):
    # Item ID (generated)
    ws.cell(r, 1, f'=IF(ISBLANK(D{r}),"",ROW()-1)')
    gen_cell(ws.cell(r, 1))
    # Accession Number (input)
    input_cell(ws.cell(r, 2))
    # Call Number (generated)
    ws.cell(r, 3, (
        f'=IF(ISBLANK(D{r}),"",'
        f'MID(K{r},SEARCH("(",K{r})+1,SEARCH(")",K{r})-SEARCH("(",K{r})-1)&"-"&'
        f'MID(L{r},SEARCH("(",L{r})+1,SEARCH(")",L{r})-SEARCH("(",L{r})-1)&"-"&'
        f'H{r}&"-"&TEXT(A{r},"000")&'
        f'IF(ISBLANK(N{r}),"","-"&SUBSTITUTE(N{r}," ","")))'
    ))
    gen_cell(ws.cell(r, 3))
    # Title, Subtitle, Creator (input)
    for col in (4, 5, 6): input_cell(ws.cell(r, col))
    # Author Code Override (input)
    input_cell(ws.cell(r, 7))
    # Author Code (generated, surname-aware)
    ws.cell(r, 8, (
        f'=IF(ISBLANK(F{r}),"",'
        f'IF(G{r}<>"",UPPER(TRIM(G{r})),'
        f'IF(ISNUMBER(SEARCH(" ",TRIM(F{r}))),'
        f'UPPER(LEFT(TRIM(RIGHT(SUBSTITUTE(TRIM(F{r})," ",REPT(" ",100)),100)),Settings!$B$9)),'
        f'UPPER(LEFT(TRIM(F{r}),Settings!$B$9)))))'
    ))
    gen_cell(ws.cell(r, 8))
    # Editor, Translator (input)
    for col in (9, 10): input_cell(ws.cell(r, col))
    # Type, Genre, Language (input, dropdowns)
    for col in (11, 12, 13): input_cell(ws.cell(r, col))
    # Volume, Edition, Publisher, Pub Place, Pub Year (input)
    for col in (14, 15, 16, 17, 18): input_cell(ws.cell(r, col))
    # ISBN, ISSN, Pages, Series (input)
    for col in (19, 20, 21, 22): input_cell(ws.cell(r, col))
    # Location Type (input dropdown), Location Identifier (input)
    for col in (23, 24): input_cell(ws.cell(r, col))
    # Condition (input)
    input_cell(ws.cell(r, 25))
    # Circulation Status (generated from Circulation sheet)
    ws.cell(r, 26, (
        f'=IF(ISBLANK(D{r}),"",'
        f'IF(COUNTIFS(Circulation!$B$2:$B${ROWS_CIR},A{r},Circulation!$H$2:$H${ROWS_CIR},"Active")>0,"On Loan",'
        f'"Available"))'
    ))
    gen_cell(ws.cell(r, 26))
    # Acquisition Date, Type, Price, Donor, Notes, Rights, Description
    for col in range(27, 34): input_cell(ws.cell(r, col))

# Data validations
dv_type    = DataValidation(type="list", formula1="=ResourceTypeList", allow_blank=True)
dv_genre   = DataValidation(type="list", formula1="=GenreList", allow_blank=True)
dv_lang    = DataValidation(type="list", formula1="=LanguageList", allow_blank=True)
dv_loctype = DataValidation(type="list", formula1="=LocationTypeList", allow_blank=True)
dv_cond    = DataValidation(type="list", formula1="=ConditionList", allow_blank=True)
dv_acq     = DataValidation(type="list", formula1="=AcqTypeList", allow_blank=True)
dv_rights  = DataValidation(type="list", formula1="=RightsList", allow_blank=True)
for dv in (dv_type, dv_genre, dv_lang, dv_loctype, dv_cond, dv_acq, dv_rights):
    ws.add_data_validation(dv)
dv_type.add(f"K2:K{ROWS_CAT+1}")
dv_genre.add(f"L2:L{ROWS_CAT+1}")
dv_lang.add(f"M2:M{ROWS_CAT+1}")
dv_loctype.add(f"W2:W{ROWS_CAT+1}")
dv_cond.add(f"Y2:Y{ROWS_CAT+1}")
dv_acq.add(f"AB2:AB{ROWS_CAT+1}")
dv_rights.add(f"AF2:AF{ROWS_CAT+1}")

# Conditional formatting — missing title
ws.conditional_formatting.add(
    f"D2:D{ROWS_CAT+1}",
    FormulaRule(formula=[f'ISBLANK($D2)'],
                fill=PatternFill("solid", fgColor=ERROR_BG), stopIfTrue=False))
ws.conditional_formatting.add(
    f"F2:F{ROWS_CAT+1}",
    FormulaRule(formula=[f'AND($D2<>"",ISBLANK($F2))'],
                fill=PatternFill("solid", fgColor=WARN_BG), stopIfTrue=False))
ws.conditional_formatting.add(
    f"Z2:Z{ROWS_CAT+1}",
    FormulaRule(formula=['$Z2="On Loan"'],
                fill=PatternFill("solid", fgColor=WARN_BG), stopIfTrue=False))

# ── Sample data (rows 2–21 preserve original collection) ──
samples = [
    ("ACC-0001","Rabindra Rachanabali","Vol 1","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 1","","","","","","Shelf","A2","Good","","","","","","","Read: complete"),
    ("ACC-0002","Rabindra Rachanabali","Vol 2","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 2","","","","","","Shelf","A2","Good","","","","","","",""),
    ("ACC-0003","Rabindra Rachanabali","Vol 3","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 3","","","","","","Shelf","A2","Good","","","","","","",""),
    ("ACC-0004","Rabindra Rachanabali","Vol 4","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 4","","","","","","Shelf","A2","Good","","","","","","",""),
    ("ACC-0005","Rabindra Rachanabali","Vol 5","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 5","","","","","","Shelf","A2","Good","","","","","","",""),
    ("ACC-0006","Rabindra Rachanabali","Vol 6","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 6","","","","","","Shelf","A2","Good","","","","","","",""),
    ("ACC-0007","Rabindra Rachanabali","Vol 7","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 7","","","","","","Shelf","A2","Good","","","","","","",""),
    ("ACC-0008","Rabindra Rachanabali","Vol 8","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 8","","","","","","Shelf","A2","Good","","","","","","",""),
    ("ACC-0009","Rabindra Rachanabali","Vol 9","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 9","","","","","","Shelf","A1","Good","","","","","","",""),
    ("ACC-0010","Rabindra Rachanabali","Vol 10","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 10","","","","","","Shelf","A1","Good","","","","","","",""),
    ("ACC-0011","Rabindra Rachanabali","Vol 11","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 11","","","","","","Shelf","A1","Good","","","","","","",""),
    ("ACC-0012","Rabindra Rachanabali","Vol 12","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 12","","","","","","Shelf","A1","Good","","","","","","",""),
    ("ACC-0013","Rabindra Rachanabali","Vol 13","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 13","","","","","","Shelf","A1","Good","","","","","","",""),
    ("ACC-0014","Rabindra Rachanabali","Vol 14","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 14","","","","","","Shelf","A1","Good","","","","","","",""),
    ("ACC-0015","Rabindra Rachanabali","Vol 15","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 15","","","","","","Shelf","A1","Good","","","","","","",""),
    ("ACC-0016","Rabindra Rachanabali","Vol 16","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 16","","","","","","Shelf","A1","Good","","","","","","",""),
    ("ACC-0017","Rabindra Rachanabali","Vol 17","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 17","","","","","","Shelf","A1","Good","","","","","","",""),
    ("ACC-0018","Rabindra Rachanabali","Vol 18","Rabindranath Tagore","Fiction (FIC)","Literature (LIT)","Bengali","Vol 18","","","","","","Shelf","A1","Good","","","","","","",""),
    ("ACC-0019","নীলকণ্ঠ পাখির খোঁজে","Vol 1","অতীন বন্দ্যোপাধ্যায়","Fiction (FIC)","Literature (LIT)","Bengali","Vol 1","","","","","","Shelf","B1","Good","","","","","",""),
    ("ACC-0020","Desh","","","Magazine","Literature (LIT)","Bengali","","","","","","","Magazine Rack","MR1","Good","","","","","","Monthly issues"),
]
for i, row in enumerate(samples):
    r = i + 2
    (acc, title, sub, creator, typ, genre, lang, vol, ed, pub, pubplace, pubyear,
     isbn, loctype, locid, cond, acqdate, acqtype, price, donor, notes, desc) = (
        row + ("",) * (22 - len(row))
    )[:22]
    ws.cell(r, 2, acc)
    ws.cell(r, 4, title)
    ws.cell(r, 5, sub)
    ws.cell(r, 6, creator)
    ws.cell(r, 11, typ)
    ws.cell(r, 12, genre)
    ws.cell(r, 13, lang)
    ws.cell(r, 14, vol)
    ws.cell(r, 15, ed)
    ws.cell(r, 16, pub)
    ws.cell(r, 17, pubplace)
    ws.cell(r, 18, pubyear)
    ws.cell(r, 19, isbn)
    ws.cell(r, 23, loctype)
    ws.cell(r, 24, locid)
    ws.cell(r, 25, cond)
    ws.cell(r, 27, acqdate)
    ws.cell(r, 28, acqtype)
    ws.cell(r, 29, price)
    ws.cell(r, 30, donor)
    ws.cell(r, 31, notes)
    ws.cell(r, 33, desc)

ws.protection.sheet = True
ws.protection.password = "library"

# ═════════════════════════════════════════════════════════════════
# 4. DIGITAL_CATALOG
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Digital_Catalog")
dig_cols = [
    ("Resource ID", 8), ("Digital Call Number", 26),
    ("Title", 32), ("Creator", 26), ("Contributor", 22),
    ("Resource Type", 18), ("Genre", 18), ("Language", 12),
    ("Format", 12), ("MIME Type", 20), ("File Extension", 12),
    ("File Size (bytes)", 14), ("Storage Path", 32), ("URI", 30),
    ("OCR Status", 20), ("Searchable", 10),
    ("Date Created", 14), ("Date Added", 14),
    ("Rights", 16), ("Access Level", 16),
    ("Description", 30),
    ("Related Physical ID", 14), ("Related Work ID", 12),
    ("Archive ID", 12), ("Checksum", 24),
    ("Reading/Usage Status", 14), ("Preservation Status", 14),
]
for i, (n, w) in enumerate(dig_cols, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
    hdr(ws.cell(1, i, n))
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(dig_cols))}{ROWS_DIG+1}"

for r in range(2, ROWS_DIG + 2):
    ws.cell(r, 1, f'=IF(ISBLANK(C{r}),"",ROW()-1)'); gen_cell(ws.cell(r, 1))
    ws.cell(r, 2, (
        f'=IF(ISBLANK(C{r}),"","DIG-"&'
        f'MID(F{r},SEARCH("(",F{r})+1,3)&"-"&'
        f'MID(G{r},SEARCH("(",G{r})+1,3)&"-"&'
        f'UPPER(LEFT(TRIM(RIGHT(SUBSTITUTE(TRIM(D{r})," ",REPT(" ",100)),100)),3))&"-"&'
        f'TEXT(A{r},"000"))'
    )); gen_cell(ws.cell(r, 2))
    for col in range(3, len(dig_cols) + 1):
        input_cell(ws.cell(r, col))
    gen_cell(ws.cell(r, 18))  # Date Added — locked though input-like

# validations
for fml, rng in [
    ("=ResourceTypeList", f"F2:F{ROWS_DIG+1}"),
    ("=GenreList",        f"G2:G{ROWS_DIG+1}"),
    ("=LanguageList",     f"H2:H{ROWS_DIG+1}"),
    ("=FormatList",       f"I2:I{ROWS_DIG+1}"),
    ("=OCRStatusList",    f"O2:O{ROWS_DIG+1}"),
    ("=RightsList",       f"S2:S{ROWS_DIG+1}"),
    ("=AccessLevelList",  f"T2:T{ROWS_DIG+1}"),
    ("=ReadingStatusList",f"Z2:Z{ROWS_DIG+1}"),
    ("=PreservationList", f"AA2:AA{ROWS_DIG+1}"),
]:
    dv = DataValidation(type="list", formula1=fml, allow_blank=True)
    ws.add_data_validation(dv); dv.add(rng)

# sample data
dig_samples = [
    ("The Pocket Oracle","Baltasar Gracian","","Book","Philosophy (PHI)","English",
     "PDF","application/pdf",".pdf","","Drive/Books/Philosophy","","Searchable (OCR)","Yes",
     "2024-01-15","2024-01-15","Copyright","Private","A pocket edition of Gracian's aphorisms.",
     "","","","","Unread","Master"),
    ("নীলকণ্ঠ পাখির খোঁজে","অতীন বন্দ্যোপাধ্যায়","","Book","Literature (LIT)","Bengali",
     "PDF","application/pdf",".pdf","","Drive/Books/Bengali_Classics","","Scanned (No OCR)","No",
     "2023-06-01","2023-06-01","Copyright","Private","Scanned copy of Bengali novel.",
     "","","","","Unread","Master"),
    ("The Great Derangement","Amitav Ghosh","","Book","Essays (ESS)","English",
     "EPUB","application/epub+zip",".epub","","Drive/Books/Essays","","Searchable (OCR)","Yes",
     "2023-02-10","2023-02-10","Copyright","Private","Climate change essays.",
     "","","","","Read","Master"),
    ("Bengali literary reading","","","Audio Recording","Literature (LIT)","Bengali",
     "FLAC","audio/flac",".flac","","Drive/Archive/Audio","","Not Applicable","No",
     "2023-08-12","2023-08-12","User-Owned","Private","Live reading of Tagore poems.",
     "","","","","N/A","Master"),
    ("Interview with an author","","","Video Recording","Literature (LIT)","Bengali",
     "MP4","video/mp4",".mp4","","Drive/Archive/Video","","Not Applicable","No",
     "2024-03-01","2024-03-01","User-Owned","Restricted","Recorded interview.",
     "","","","","N/A","Master"),
    ("Scanned Bengali manuscript","","","Manuscript","Literature (LIT)","Bengali",
     "TIFF","image/tiff",".tiff","","Archive/Manuscripts","","Scanned (No OCR)","No",
     "2022-01-01","2022-01-01","Public Domain","Research Only","High-res scan of 19th-c manuscript.",
     "","","ARC-0001","","N/A","Master"),
]
for i, row in enumerate(dig_samples):
    r = i + 2
    for j, val in enumerate(row, start=3):
        if j == 18:  # skip Date Added formula
            continue
        ws.cell(r, j, val)

ws.protection.sheet = True
ws.protection.password = "library"

# ═════════════════════════════════════════════════════════════════
# 5. ARCHIVE
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Archive")
arc_cols = [
    ("Archive ID", 12), ("Title", 32), ("Creator", 24), ("Date", 14),
    ("Description", 34), ("Resource Type", 18), ("Physical/Digital", 14),
    ("Language", 12), ("Provenance", 26), ("Acquisition", 22),
    ("Condition", 14), ("Location Type", 14), ("Location Identifier", 18),
    ("Digital Surrogate", 20), ("Digitisation Status", 18),
    ("OCR Status", 16), ("Rights", 16), ("Access Level", 16),
    ("Related Resource ID", 14), ("Notes", 30),
]
for i, (n, w) in enumerate(arc_cols, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
    hdr(ws.cell(1, i, n))
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(arc_cols))}{ROWS_ARC+1}"

for r in range(2, ROWS_ARC + 2):
    ws.cell(r, 1, f'=IF(ISBLANK(B{r}),"","ARC-"&TEXT(ROW()-1,"0000"))')
    gen_cell(ws.cell(r, 1))
    for col in range(2, len(arc_cols) + 1):
        input_cell(ws.cell(r, col))

for fml, rng in [
    ("=ResourceTypeList",     f"F2:F{ROWS_ARC+1}"),
    ("=PhysDigiList",         f"G2:G{ROWS_ARC+1}"),
    ("=LanguageList",         f"H2:H{ROWS_ARC+1}"),
    ("=ConditionList",        f"K2:K{ROWS_ARC+1}"),
    ("=LocationTypeList",     f"L2:L{ROWS_ARC+1}"),
    ("=DigitisationList",     f"O2:O{ROWS_ARC+1}"),
    ("=OCRStatusList",        f"P2:P{ROWS_ARC+1}"),
    ("=RightsList",           f"Q2:Q{ROWS_ARC+1}"),
    ("=AccessLevelList",      f"R2:R{ROWS_ARC+1}"),
]:
    dv = DataValidation(type="list", formula1=fml, allow_blank=True)
    ws.add_data_validation(dv); dv.add(rng)

# sample archive data
arc_samples = [
    ("Historical newspaper clipping","","1947-08-15","Bengali newspaper clipping from independence era.",
     "Newspaper Clipping","Physical","Bengali","Family collection","Inherited","Fair","Archive","ARCH-01",
     "","Not Digitised","Not Applicable","Unknown","Private","","Fragile; store in acid-free sleeve."),
    ("Scanned Bengali manuscript","","1890","Digitised manuscript.",
     "Manuscript","Digital","Bengali","Purchased at auction","Purchase","Good","Archive","ARCH-02",
     "","Scanned","OCR Pending","Public Domain","Research Only","","High-res TIFF master."),
]
for i, row in enumerate(arc_samples):
    r = i + 2
    for j, val in enumerate(row, start=2):
        ws.cell(r, j, val)

ws.protection.sheet = True
ws.protection.password = "library"

# ═════════════════════════════════════════════════════════════════
# 6. CIRCULATION
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Circulation")
cir_cols = [
    ("Loan ID", 10), ("Item ID", 10), ("Member ID", 12),
    ("Borrower Name", 22), ("Checkout Date", 14), ("Due Date", 14),
    ("Return Date", 14), ("Status", 12), ("Notes", 30),
]
for i, (n, w) in enumerate(cir_cols, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
    hdr(ws.cell(1, i, n))
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:I{ROWS_CIR+1}"

for r in range(2, ROWS_CIR + 2):
    ws.cell(r, 1, f'=IF(ISBLANK(B{r}),"",ROW()-1)'); gen_cell(ws.cell(r, 1))
    input_cell(ws.cell(r, 2))
    input_cell(ws.cell(r, 3))
    ws.cell(r, 4, f'=IFERROR(IF(ISBLANK(C{r}),"",VLOOKUP(C{r},Members!$A$2:$B${ROWS_MEM+1},2,FALSE)),"")')
    gen_cell(ws.cell(r, 4))
    input_cell(ws.cell(r, 5))
    ws.cell(r, 6, f'=IF(ISBLANK(E{r}),"",E{r}+Settings!$B$7)'); gen_cell(ws.cell(r, 6))
    input_cell(ws.cell(r, 7))
    ws.cell(r, 8, (
        f'=IF(ISBLANK(B{r}),"",'
        f'IF(NOT(ISBLANK(G{r})),"Returned",'
        f'IF(TODAY()>F{r},"Overdue","Active")))'
    ))
    gen_cell(ws.cell(r, 8))
    input_cell(ws.cell(r, 9))

ws.protection.sheet = True
ws.protection.password = "library"

# ═════════════════════════════════════════════════════════════════
# 7. MEMBERS
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Members")
mem_cols = [("Member ID",12),("Name",26),("Contact",26),
            ("Membership Type",18),("Join Date",14),("Status",12),("Notes",30)]
for i, (n, w) in enumerate(mem_cols, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
    hdr(ws.cell(1, i, n))
ws.freeze_panes = "A2"
for r in range(2, ROWS_MEM + 2):
    for col in range(1, 8): input_cell(ws.cell(r, col))
for fml, rng in [("=MembershipList", f"D2:D{ROWS_MEM+1}"), ("=MemberStatusList", f"F2:F{ROWS_MEM+1}")]:
    dv = DataValidation(type="list", formula1=fml, allow_blank=True)
    ws.add_data_validation(dv); dv.add(rng)

ws.cell(2, 1, "M001"); ws.cell(2, 2, "Sample Borrower")
ws.cell(2, 4, "Public"); ws.cell(2, 5, "2024-01-01"); ws.cell(2, 6, "Active")
ws.protection.sheet = True
ws.protection.password = "library"

# ═════════════════════════════════════════════════════════════════
# 8. SPINE_LABELS
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Spine_Labels")
spn_cols = [("Item ID",10),("Call Number",28),("Title",34),("Volume/Issue",12),
            ("Creator",24),("Location",20),("Label Type",18)]
for i, (n, w) in enumerate(spn_cols, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
    hdr(ws.cell(1, i, n))
ws.freeze_panes = "A2"
for r in range(2, ROWS_SPN + 2):
    ws.cell(r, 1, f'=Catalog!A{r}'); gen_cell(ws.cell(r, 1))
    ws.cell(r, 2, f'=Catalog!C{r}'); gen_cell(ws.cell(r, 2))
    ws.cell(r, 3, f'=Catalog!D{r}'); gen_cell(ws.cell(r, 3))
    ws.cell(r, 4, f'=Catalog!N{r}'); gen_cell(ws.cell(r, 4))
    ws.cell(r, 5, f'=Catalog!F{r}'); gen_cell(ws.cell(r, 5))
    ws.cell(r, 6, f'=IF(Catalog!W{r}="","",Catalog!W{r}&" "&Catalog!X{r})'); gen_cell(ws.cell(r, 6))
    input_cell(ws.cell(r, 7))
dv = DataValidation(type="list", formula1='"Standard Card,Library Block,Minimal Spine"', allow_blank=True)
ws.add_data_validation(dv); dv.add(f"G2:G{ROWS_SPN+1}")
ws.protection.sheet = True
ws.protection.password = "library"

# ═════════════════════════════════════════════════════════════════
# 9. DUBLIN_CORE
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Dublin_Core")
ws.column_dimensions["A"].width = 20
ws.column_dimensions["B"].width = 32
ws.column_dimensions["C"].width = 40
ws["A1"] = "DUBLIN CORE METADATA MAPPING"
ws["A1"].font = Font(name=FONT, bold=True, size=14, color=HEADER_BG)
ws.merge_cells("A1:C1")
for i, h in enumerate(["Dublin Core Element","Catalog Field","Notes"], start=1):
    hdr(ws.cell(3, i, h))
mapping = [
    ("title","Title","Standard DC element"),
    ("creator","Creator","Standard DC element"),
    ("contributor","Editor, Translator","Standard DC element; application-specific split"),
    ("publisher","Publisher","Standard DC element"),
    ("date","Publication Year","Standard DC element"),
    ("language","Language","Standard DC element"),
    ("type","Type","Standard DC element; project vocabulary"),
    ("format","Format, Pages","Standard DC element"),
    ("identifier","ISBN, ISSN, Accession Number, Call Number","Standard DC element"),
    ("subject","Genre","Standard DC element"),
    ("description","Notes, Description","Standard DC element"),
    ("source","Donor/Vendor, Acquisition Type","Standard DC element"),
    ("relation","Volume, Series, Related ID","Standard DC element; isPartOf qualifier"),
    ("coverage","Publication Place","Standard DC element; spatial qualifier"),
    ("rights","Rights","Standard DC element"),
    ("—","Item ID, Call Number, Author Code, Location","Application-specific extension (not DC)"),
    ("—","Circulation Status, Condition","Application-specific extension (not DC)"),
    ("—","Digital storage path, OCR status, Checksum","Application-specific extension (not DC)"),
]
for i, row in enumerate(mapping, start=4):
    for j, val in enumerate(row, start=1):
        c = ws.cell(i, j, val); c.font = Font(name=FONT, size=10); c.border = BORDER

# Export-oriented view
ws["A25"] = "EXPORT VIEW — Dublin Core element columns"
ws["A25"].font = Font(name=FONT, bold=True, size=12, color=HEADER_BG)
dc_view = ["dc.title","dc.creator","dc.contributor","dc.publisher","dc.date",
           "dc.language","dc.type","dc.format","dc.identifier","dc.subject",
           "dc.description","dc.source","dc.relation","dc.coverage","dc.rights"]
for j, name in enumerate(dc_view, start=1):
    hdr(ws.cell(27, j, name))
for r in range(2, ROWS_CAT + 2):
    row = r + 26
    ws.cell(row, 1, f'=Catalog!D{r}')
    ws.cell(row, 2, f'=Catalog!F{r}')
    ws.cell(row, 3, f'=IF(Catalog!I{r}="","",Catalog!I{r})&IF(AND(Catalog!I{r}<>"",Catalog!J{r}<>""),"; ","")&IF(Catalog!J{r}="","",Catalog!J{r})')
    ws.cell(row, 4, f'=Catalog!P{r}')
    ws.cell(row, 5, f'=Catalog!R{r}')
    ws.cell(row, 6, f'=Catalog!M{r}')
    ws.cell(row, 7, f'=Catalog!K{r}')
    ws.cell(row, 8, f'=IF(Catalog!S{r}<>"","ISBN:"&Catalog!S{r},IF(Catalog!T{r}<>"","ISSN:"&Catalog!T{r},""))')
    ws.cell(row, 9, f'=Catalog!B{r}')
    ws.cell(row, 10, f'=Catalog!L{r}')
    ws.cell(row, 11, f'=IF(Catalog!AG{r}<>"",Catalog!AG{r},Catalog!AE{r})')
    ws.cell(row, 12, f'=Catalog!AD{r}')
    ws.cell(row, 13, f'=Catalog!V{r}')
    ws.cell(row, 14, f'=Catalog!Q{r}')
    ws.cell(row, 15, f'=Catalog!AF{r}')
    for j in range(1, 16): gen_cell(ws.cell(row, j))

ws.protection.sheet = True
ws.protection.password = "library"

# ═════════════════════════════════════════════════════════════════
# 10. DATA_QUALITY
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Data_Quality")
ws.column_dimensions["A"].width = 32
ws.column_dimensions["B"].width = 16
ws["A1"] = "DATA QUALITY REPORT"
ws["A1"].font = Font(name=FONT, bold=True, size=14, color=HEADER_BG)
ws.merge_cells("A1:B1")

checks = [
    ("Total Catalog records",       f'=COUNTA(Catalog!$D$2:$D${ROWS_CAT+1})'),
    ("Records missing Title",       f'=COUNTA(Catalog!$A$2:$A${ROWS_CAT+1})-COUNTA(Catalog!$D$2:$D${ROWS_CAT+1})'),
    ("Records missing Creator",     f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$F$2:$F${ROWS_CAT+1}=""))'),
    ("Records missing Type",        f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$K$2:$K${ROWS_CAT+1}=""))'),
    ("Records missing Genre",       f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$L$2:$L${ROWS_CAT+1}=""))'),
    ("Records missing Location Type", f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$W$2:$W${ROWS_CAT+1}=""))'),
    ("Records missing Location Identifier", f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$X$2:$X${ROWS_CAT+1}=""))'),
    ("Duplicate ISBNs",             f'=SUMPRODUCT((Catalog!$S$2:$S${ROWS_CAT+1}<>"")*(COUNTIF(Catalog!$S$2:$S${ROWS_CAT+1},Catalog!$S$2:$S${ROWS_CAT+1})>1))/2'),
    ("Duplicate ISSNs",             f'=SUMPRODUCT((Catalog!$T$2:$T${ROWS_CAT+1}<>"")*(COUNTIF(Catalog!$T$2:$T${ROWS_CAT+1},Catalog!$T$2:$T${ROWS_CAT+1})>1))/2'),
    ("Duplicate Accession Numbers", f'=SUMPRODUCT((Catalog!$B$2:$B${ROWS_CAT+1}<>"")*(COUNTIF(Catalog!$B$2:$B${ROWS_CAT+1},Catalog!$B$2:$B${ROWS_CAT+1})>1))/2'),
    ("Digital resources",           f'=COUNTA(Digital_Catalog!$C$2:$C${ROWS_DIG+1})'),
    ("Archive objects",             f'=COUNTA(Archive!$B$2:$B${ROWS_ARC+1})'),
    ("Active loans",                f'=COUNTIF(Circulation!$H$2:$H${ROWS_CIR+1},"Active")'),
    ("Overdue loans",               f'=COUNTIF(Circulation!$H$2:$H${ROWS_CIR+1},"Overdue")'),
]
ws["A3"] = "Metric"; ws["B3"] = "Value"
hdr(ws["A3"]); hdr(ws["B3"])
for i, (label, f) in enumerate(checks, start=4):
    ws.cell(i, 1, label).font = Font(name=FONT, bold=True, size=10)
    ws.cell(i, 2, f).font = Font(name=FONT, size=10)
    ws.cell(i, 2).alignment = Alignment(horizontal="right")
    ws.cell(i, 2).border = BORDER

# Row-level checks
ws["A25"] = "RECORD-LEVEL CHECKS"
ws["A25"].font = Font(name=FONT, bold=True, size=12, color=HEADER_BG)
for j, name in enumerate(["Item ID","Title","Issues"], start=1):
    hdr(ws.cell(26, j, name))
for r in range(2, ROWS_CAT + 2):
    row = r + 25
    ws.cell(row, 1, f'=Catalog!A{r}').font = Font(name=FONT, size=9)
    ws.cell(row, 2, f'=Catalog!D{r}').font = Font(name=FONT, size=9)
    ws.cell(row, 3, (
        f'=IF(Catalog!D{r}="","",'
        f'IF(ISBLANK(Catalog!F{r}),"Missing Creator; ","")&'
        f'IF(ISBLANK(Catalog!K{r}),"Missing Type; ","")&'
        f'IF(ISBLANK(Catalog!L{r}),"Missing Genre; ","")&'
        f'IF(ISBLANK(Catalog!W{r}),"Missing Location Type; ","")&'
        f'IF(ISBLANK(Catalog!X{r}),"Missing Location Identifier; ","")&'
        f'IF(AND(Catalog!S{r}<>"",COUNTIF(Catalog!$S$2:$S${ROWS_CAT+1},Catalog!S{r})>1),"Duplicate ISBN; ","")&'
        f'IF(AND(Catalog!B{r}<>"",COUNTIF(Catalog!$B$2:$B${ROWS_CAT+1},Catalog!B{r})>1),"Duplicate Accession; ",""))'
    )).font = Font(name=FONT, size=9, color="C0392B")
ws.freeze_panes = "A27"
ws.protection.sheet = True
ws.protection.password = "library"

# ═════════════════════════════════════════════════════════════════
# 11. DASHBOARD
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Dashboard")
ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 34
ws.column_dimensions["C"].width = 16
ws.column_dimensions["D"].width = 4
ws.column_dimensions["E"].width = 34
ws.column_dimensions["F"].width = 16

ws["B2"] = "PERSONAL LIBRARY MANAGER & DIGITAL ARCHIVE — DASHBOARD"
ws["B2"].font = Font(name=FONT, bold=True, size=16, color=HEADER_BG)
ws.merge_cells("B2:F2")

def dash_block(anchor_row, title, items, col_letter):
    ws.cell(anchor_row, ord(col_letter)-64, title)
    c = ws.cell(anchor_row, ord(col_letter)-64)
    c.font = Font(name=FONT, bold=True, size=12, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=SECTION_BG)
    c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    end_col = get_column_letter(ord(col_letter)-64+1)
    ws.merge_cells(f"{col_letter}{anchor_row}:{end_col}{anchor_row}")
    for i, (label, f) in enumerate(items, start=1):
        r = anchor_row + i
        ws.cell(r, ord(col_letter)-64, label).font = Font(name=FONT, size=10)
        ws.cell(r, ord(col_letter)-64+1, f).font = Font(name=FONT, bold=True, size=10)
        ws.cell(r, ord(col_letter)-64+1).alignment = Alignment(horizontal="right")

dash_block(4, "COLLECTION", [
    ("Physical items",   f'=COUNTA(Catalog!A2:A{ROWS_CAT+1})'),
    ("Digital resources",f'=COUNTA(Digital_Catalog!A2:A{ROWS_DIG+1})'),
    ("Archive objects",  f'=COUNTA(Archive!A2:A{ROWS_ARC+1})'),
    ("Books",            f'=COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Book")'),
    ("Magazines",        f'=COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Magazine")'),
    ("Journals",         f'=COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Journal")'),
    ("Other materials",  f'=COUNTA(Catalog!A2:A{ROWS_CAT+1})-COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Book")-COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Magazine")-COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Journal")'),
], "B")

dash_block(4, "METADATA", [
    ("Missing titles",       f'=COUNTA(Catalog!A2:A{ROWS_CAT+1})-COUNTA(Catalog!D2:D{ROWS_CAT+1})'),
    ("Missing creators",     f'=SUMPRODUCT((Catalog!D2:D{ROWS_CAT+1}<>"")*(Catalog!F2:F{ROWS_CAT+1}=""))'),
    ("Missing locations",    f'=SUMPRODUCT((Catalog!D2:D{ROWS_CAT+1}<>"")*(Catalog!W2:W{ROWS_CAT+1}=""))'),
    ("Duplicate ISBNs",      f'=SUMPRODUCT((Catalog!S2:S{ROWS_CAT+1}<>"")*(COUNTIF(Catalog!S2:S{ROWS_CAT+1},Catalog!S2:S{ROWS_CAT+1})>1))/2'),
    ("Duplicate Accessions", f'=SUMPRODUCT((Catalog!B2:B{ROWS_CAT+1}<>"")*(COUNTIF(Catalog!B2:B{ROWS_CAT+1},Catalog!B2:B{ROWS_CAT+1})>1))/2'),
], "E")

dash_block(14, "READING STATUS", [
    ("Read",    f'=COUNTIF(Digital_Catalog!Z2:Z{ROWS_DIG+1},"Read")'),
    ("Reading", f'=COUNTIF(Digital_Catalog!Z2:Z{ROWS_DIG+1},"Reading")'),
    ("Unread",  f'=COUNTIF(Digital_Catalog!Z2:Z{ROWS_DIG+1},"Unread")'),
], "B")

dash_block(14, "CIRCULATION", [
    ("Available",  f'=COUNTIF(Catalog!Z2:Z{ROWS_CAT+1},"Available")'),
    ("On Loan",    f'=COUNTIF(Catalog!Z2:Z{ROWS_CAT+1},"On Loan")'),
    ("Overdue",    f'=COUNTIF(Circulation!H2:H{ROWS_CIR+1},"Overdue")'),
], "E")

dash_block(19, "DIGITISATION", [
    ("Not Digitised",    f'=COUNTIF(Digital_Catalog!O2:O{ROWS_DIG+1},"Not Digitised")+COUNTIF(Archive!O2:O{ROWS_ARC+1},"Not Digitised")'),
    ("OCR Pending",      f'=COUNTIF(Digital_Catalog!O2:O{ROWS_DIG+1},"OCR Pending")+COUNTIF(Archive!P2:P{ROWS_ARC+1},"OCR Pending")'),
    ("OCR Complete",     f'=COUNTIF(Digital_Catalog!O2:O{ROWS_DIG+1},"OCR Complete")+COUNTIF(Archive!P2:P{ROWS_ARC+1},"OCR Complete")'),
], "B")

dash_block(19, "ARCHIVE", [
    ("Physical",   f'=COUNTIF(Archive!G2:G{ROWS_ARC+1},"Physical")'),
    ("Digital",    f'=COUNTIF(Archive!G2:G{ROWS_ARC+1},"Digital")'),
    ("Both",       f'=COUNTIF(Archive!G2:G{ROWS_ARC+1},"Both")'),
], "E")

ws.protection.sheet = True
ws.protection.password = "library"

# ═════════════════════════════════════════════════════════════════
# 12. JSON_EXPORT
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("JSON_Export")
ws.column_dimensions["A"].width = 10
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 24
ws.column_dimensions["D"].width = 40
ws["A1"] = "STRUCTURED EXPORT FOR JSON / CSV PIPELINE"
ws["A1"].font = Font(name=FONT, bold=True, size=14, color=HEADER_BG)
ws.merge_cells("A1:D1")
ws["A3"] = ("This sheet provides a flattened, structured view of the catalog "
            "suitable for conversion to JSON by a future Python/Node script. "
            "Field names use snake_case.")
ws["A3"].font = Font(name=FONT, italic=True, size=10)
ws.merge_cells("A3:D3")
for j, name in enumerate(["item_id","title","creator","json_snippet"], start=1):
    hdr(ws.cell(5, j, name))
for r in range(2, ROWS_CAT + 2):
    row = r + 4
    ws.cell(row, 1, f'=Catalog!A{r}').font = Font(name=FONT, size=9)
    ws.cell(row, 2, f'=Catalog!D{r}').font = Font(name=FONT, size=9)
    ws.cell(row, 3, f'=Catalog!F{r}').font = Font(name=FONT, size=9)
    ws.cell(row, 4, (
        f'=IF(Catalog!D{r}="","",'
        f'"{{""item_id"":"&Catalog!A{r}&",'
        f'""title"":"""&SUBSTITUTE(Catalog!D{r},"""","\\""")&""",'
        f'""creator"":"""&SUBSTITUTE(Catalog!F{r},"""","\\""")&""",'
        f'""language"":"""&Catalog!M{r}&""",'
        f'""identifier"":"""&Catalog!C{r}&"""}}"'
        f')'
    )).font = Font(name="Consolas", size=9)
ws.freeze_panes = "A6"
ws.protection.sheet = True
ws.protection.password = "library"

# ═════════════════════════════════════════════════════════════════
# 13. README
# ═════════════════════════════════════════════════════════════════
ws = wb.create_sheet("README", 0)  # first sheet
ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 100
readme_lines = [
    ("H1", "Personal Library Manager & Digital Archive"),
    ("H2", "Version 2.0 — Excel Application"),
    ("P",  "Excel is the primary structured-data and metadata layer. The future HTML application is a separate presentation/discovery/printing layer."),
    ("H3", "Quick Start"),
    ("P",  "1. Open the Settings sheet and set your library name, loan period, and defaults."),
    ("P",  "2. Open the Catalog sheet. Fill in Accession Number, Title, Creator, Type, Genre, Language, Location Type, and Location Identifier. Item ID, Author Code, Call Number, and Circulation Status are generated automatically."),
    ("P",  "3. Use the Digital_Catalog sheet for PDFs, images, audio, video, datasets, and other digital resources."),
    ("P",  "4. Use the Archive sheet for physical and digital archival objects."),
    ("P",  "5. Record loans in the Circulation sheet. The Catalog's Circulation Status updates automatically."),
    ("P",  "6. Monitor completeness and issues on the Dashboard and Data_Quality sheets."),
    ("H3", "Worksheet Guide"),
    ("P",  "README — this documentation page."),
    ("P",  "Dashboard — metrics for collection, metadata, reading, circulation, digitisation, archive."),
    ("P",  "Catalog — physical books, magazines, journals, printed materials."),
    ("P",  "Digital_Catalog — digital resources of all formats."),
    ("P",  "Archive — physical and digital archival objects."),
    ("P",  "Circulation — loans, borrowers, due dates, returns."),
    ("P",  "Members — borrower registry."),
    ("P",  "Spine_Labels — derived label data for printing."),
    ("P",  "Dublin_Core — metadata mapping and export-oriented view."),
    ("P",  "Lookup_Codes — controlled vocabularies for data validation."),
    ("P",  "Data_Quality — duplicate, missing, and integrity checks."),
    ("P",  "Settings — configurable parameters used by formulas."),
    ("P",  "JSON_Export — flattened view for future JSON conversion."),
    ("H3", "Author Code Logic"),
    ("P",  "Author Code is generated from the last word of the Creator field (surname for most names). Example: 'Rabindranath Tagore' → TAG. Bengali and other Unicode scripts are preserved."),
    ("P",  "To override the automatic code, enter a value in the 'Author Code Override' column. The override takes precedence."),
    ("H3", "Call Number System"),
    ("P",  "Format: TYPE-GENRE-AUTHORCODE-ITEM-VOLUME. Example: FIC-LIT-TAG-001-Vol1."),
    ("P",  "This is the project's own structured call-number system. It is not Dewey, Library of Congress, or Cutter classification."),
    ("H3", "Location System"),
    ("P",  "Location Type is a controlled dropdown (Shelf, Magazine Rack, Journal Rack, Archive, Cabinet, Box, Other)."),
    ("P",  "Location Identifier is free text (A1, A2, MR1, ARCH-01, BOX-17, etc.). This semi-controlled model lets you control the category while keeping the physical identifier flexible."),
    ("H3", "Dublin Core"),
    ("P",  "Dublin Core is the primary metadata framework. See the Dublin_Core sheet for the explicit mapping between catalog fields and DC elements."),
    ("P",  "Project-specific fields (Item ID, Call Number, Author Code, Location, Circulation Status) are documented as application-specific extensions — they are not Dublin Core elements."),
    ("H3", "Interoperability"),
    ("P",  "CSV: every data sheet exports cleanly with UTF-8 encoding. Bengali, Hindi, and other Unicode scripts are preserved."),
    ("P",  "JSON: the JSON_Export sheet provides a flattened, structured view for future Python/Node conversion."),
    ("P",  "TEI: the catalog preserves structured fields (title, author, editor, translator, publisher, date, language, series, volume, ISBN, genre) suitable for later TEI transformation."),
    ("P",  "Bibliographic APIs: the workbook can receive ISBN/ISSN/DOI/OCLC metadata. Live lookups are deferred to the future HTML or Python layer."),
    ("H3", "Protection"),
    ("P",  "Generated columns (Item ID, Author Code, Call Number, Circulation Status, Dashboard metrics) are locked. Editable columns are unlocked."),
    ("P",  "Sheet protection password: 'library'. Change it in Settings if desired."),
    ("H3", "Important Notes"),
    ("P",  "• Do not sort the Catalog sheet by columns other than Item ID — Call Number formulas reference row position. Use AutoFilter instead."),
    ("P",  "• Accession Number is the permanent identifier. Item ID is internal."),
    ("P",  "• Bengali display uses Nirmala UI where available. On macOS, change the font to 'Noto Sans Bengali' if characters do not render."),
    ("P",  "• No VBA macros are used. The workbook is compatible with Microsoft Excel, Google Sheets, and LibreOffice Calc (minor date-format differences may occur)."),
]
row = 1
for kind, text in readme_lines:
    c = ws.cell(row, 2, text)
    if kind == "H1":
        c.font = Font(name=FONT, bold=True, size=18, color=HEADER_BG)
    elif kind == "H2":
        c.font = Font(name=FONT, bold=True, size=13, color=SECTION_BG)
    elif kind == "H3":
        c.font = Font(name=FONT, bold=True, size=11, color=HEADER_BG)
    else:
        c.font = Font(name=FONT, size=10)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if kind == "P":
        ws.row_dimensions[row].height = 30
    row += 1

# ═════════════════════════════════════════════════════════════════
# SAVE
# ═════════════════════════════════════════════════════════════════
wb.active = 0
out = "Personal Library Manager & Digital Archive.xlsx"
wb.save(out)
print(f"Workbook written: {out}")
print(f"Sheets: {wb.sheetnames}")