"""
Personal Library Manager & Digital Archive — Complete Workbook Generator
=========================================================================

Typography: Garamond, Baskerville Old Face, Book Antiqua, Courier New
Colour:     DarkSlateBlue (#483D8B), Brown (#A52A2A), Aquamarine (#7FFFD4)

Usage:
    python build_library.py            → full workbook with sample data
    python build_library.py --blank    → blank template

Requires: openpyxl
    pip install openpyxl
"""

import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════════════════════
# TYPOGRAPHY
# ═══════════════════════════════════════════════════════════════════════
FONT_TITLE    = "Garamond"                 # workbook title, dashboard banner
FONT_SECTION  = "Baskerville Old Face"      # section headers, column headers
FONT_BODY     = "Book Antiqua"              # body text, data (Caslon substitute)
FONT_MONO     = "Courier New"               # call numbers, IDs, code
FONT_SMALL    = "Book Antiqua"              # labels, notes, meta

# Sizes
SZ_TITLE   = 22
SZ_H1      = 16
SZ_H2      = 13
SZ_BODY    = 12
SZ_TABLE   = 11
SZ_SMALL   = 10
SZ_META    = 9

# ═══════════════════════════════════════════════════════════════════════
# COLOUR PALETTE
# ═══════════════════════════════════════════════════════════════════════
# Primary trio
C_DARKSLATEBLUE = "483D8B"
C_BROWN         = "A52A2A"
C_AQUAMARINE    = "7FFFD4"

# Complements
C_CREAM         = "FFF8E7"   # warm ivory — input cells
C_PALETURQ      = "AFEEEE"   # pale turquoise — subtle highlight
C_LAVENDER      = "E6E6FA"   # section dividers
C_PERU          = "CD853F"   # secondary accent
C_DARKSLATEGRAY = "2F4F4F"   # primary text
C_IVORY         = "FFFFF0"   # surface background

# Status tones
C_ERROR_SOFT    = "FFE4E1"   # MistyRose
C_WARN_SOFT     = "FFFACD"   # LemonChiffon
C_SUCCESS_SOFT  = "F0FFF0"   # Honeydew
C_GENERATED     = "E6F2F2"   # very pale aqua tint for formula cells
C_BORDER        = "B8B8C8"

thin   = Side(style="thin",   color=C_BORDER)
medium = Side(style="medium", color=C_DARKSLATEBLUE)
BORDER_THIN   = Border(left=thin, right=thin, top=thin, bottom=thin)
BORDER_HEADER = Border(left=thin, right=thin, top=thin, bottom=medium)


# ═══════════════════════════════════════════════════════════════════════
# CELL STYLE HELPERS
# ═══════════════════════════════════════════════════════════════════════
def hdr(cell):
    """Column header — Baskerville, DarkSlateBlue background."""
    cell.font = Font(name=FONT_SECTION, bold=True, color="FFFFFF",
                     size=SZ_TABLE)
    cell.fill = PatternFill("solid", fgColor=C_DARKSLATEBLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center",
                               wrap_text=True)
    cell.border = BORDER_HEADER


def hdr_brown(cell):
    """Alternate column header — Brown background."""
    cell.font = Font(name=FONT_SECTION, bold=True, color="FFFFFF",
                     size=SZ_TABLE)
    cell.fill = PatternFill("solid", fgColor=C_BROWN)
    cell.alignment = Alignment(horizontal="center", vertical="center",
                               wrap_text=True)
    cell.border = BORDER_HEADER


def input_cell(cell):
    """Editable cell — Cream background, Book Antiqua."""
    cell.fill = PatternFill("solid", fgColor=C_CREAM)
    cell.protection = Protection(locked=False)
    cell.border = BORDER_THIN
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    cell.font = Font(name=FONT_BODY, size=SZ_BODY, color=C_DARKSLATEGRAY)


def input_mono(cell):
    """Editable cell with Courier New — for IDs, codes."""
    input_cell(cell)
    cell.font = Font(name=FONT_MONO, size=SZ_BODY, color=C_DARKSLATEGRAY)


def gen_cell(cell):
    """Generated (formula) cell — pale aqua tint, locked."""
    cell.fill = PatternFill("solid", fgColor=C_GENERATED)
    cell.protection = Protection(locked=True)
    cell.border = BORDER_THIN
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    cell.font = Font(name=FONT_MONO, size=SZ_BODY, color=C_DARKSLATEBLUE)


def gen_body(cell):
    """Generated cell rendered in body font."""
    cell.fill = PatternFill("solid", fgColor=C_GENERATED)
    cell.protection = Protection(locked=True)
    cell.border = BORDER_THIN
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    cell.font = Font(name=FONT_BODY, size=SZ_BODY, color=C_DARKSLATEGRAY)


def title_cell(cell):
    """Large workbook title — Garamond, DarkSlateBlue."""
    cell.font = Font(name=FONT_TITLE, bold=True, size=SZ_TITLE,
                     color=C_DARKSLATEBLUE)


def h1_cell(cell):
    cell.font = Font(name=FONT_TITLE, bold=True, size=SZ_H1,
                     color=C_DARKSLATEBLUE)


def h2_cell(cell):
    cell.font = Font(name=FONT_SECTION, bold=True, size=SZ_H2,
                     color=C_BROWN)


def h3_cell(cell):
    cell.font = Font(name=FONT_SECTION, bold=True, size=SZ_H2,
                     color=C_DARKSLATEBLUE)


def body_cell(cell):
    cell.font = Font(name=FONT_BODY, size=SZ_BODY, color=C_DARKSLATEGRAY)


def small_cell(cell):
    cell.font = Font(name=FONT_SMALL, italic=True, size=SZ_SMALL,
                     color=C_PERU)


def meta_cell(cell):
    cell.font = Font(name=FONT_SMALL, italic=True, size=SZ_META,
                     color=C_PERU)


def section_banner(cell):
    """Section banner — DarkSlateBlue full-width."""
    cell.font = Font(name=FONT_SECTION, bold=True, size=SZ_H2,
                     color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=C_DARKSLATEBLUE)
    cell.alignment = Alignment(horizontal="left", vertical="center",
                               indent=1)


def section_banner_brown(cell):
    """Section banner — Brown full-width."""
    cell.font = Font(name=FONT_SECTION, bold=True, size=SZ_H2,
                     color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=C_BROWN)
    cell.alignment = Alignment(horizontal="left", vertical="center",
                               indent=1)


def metric_label(cell):
    cell.font = Font(name=FONT_BODY, size=SZ_BODY, color=C_DARKSLATEGRAY)
    cell.fill = PatternFill("solid", fgColor=C_CREAM)


def metric_value(cell):
    cell.font = Font(name=FONT_MONO, bold=True, size=SZ_BODY,
                     color=C_DARKSLATEBLUE)
    cell.fill = PatternFill("solid", fgColor=C_AQUAMARINE)
    cell.alignment = Alignment(horizontal="right", vertical="center")


# ═══════════════════════════════════════════════════════════════════════
# ROW LIMITS & MODE
# ═══════════════════════════════════════════════════════════════════════
ROWS_CAT = 500
ROWS_DIG = 500
ROWS_ARC = 200
ROWS_CIR = 500
ROWS_MEM = 100
ROWS_SPN = 500

BLANK_MODE = "--blank" in sys.argv

# ═══════════════════════════════════════════════════════════════════════
# WORKBOOK
# ═══════════════════════════════════════════════════════════════════════
wb = Workbook()
wb.remove(wb.active)


# ═══════════════════════════════════════════════════════════════════════
# 1. SETTINGS
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Settings")
ws.column_dimensions["A"].width = 36
ws.column_dimensions["B"].width = 42
ws.column_dimensions["C"].width = 55

title_cell(ws["A1"])
ws["A1"] = "SETTINGS — Personal Library Manager"
ws.merge_cells("A1:C1")
ws.row_dimensions[1].height = 32

ws["A3"] = "Setting"
ws["B3"] = "Value"
ws["C3"] = "Purpose"
hdr(ws["A3"]); hdr(ws["B3"]); hdr_brown(ws["C3"])

settings_data = [
    ("Library Name",                   "Personal Library",       "Display name of the library"),
    ("Collection Name",                "Main Collection",        "Name of the primary collection"),
    ("Default Language",               "English",                "Default language for new records"),
    ("Default Loan Period (days)",     14,                       "Circulation loan period in days"),
    ("Accession Prefix",               "ACC",                    "Prefix for accession numbers"),
    ("Author Code Length",             3,                        "Number of letters for auto author code"),
    ("Date Format",                    "YYYY-MM-DD",             "Display format for dates"),
    ("Default Resource Type",          "Book",                   "Default resource type"),
    ("Metadata Profile Version",       "2.1",                    "Dublin Core profile version"),
    ("Workbook Protection Password",   "library",                "Password to unprotect sheets"),
]
for i, (k, v, p) in enumerate(settings_data, start=4):
    label = ws.cell(i, 1, k)
    label.font = Font(name=FONT_BODY, bold=True, size=SZ_BODY,
                      color=C_DARKSLATEBLUE)
    label.fill = PatternFill("solid", fgColor=C_LAVENDER)
    label.border = BORDER_THIN
    val = ws.cell(i, 2, v)
    input_cell(val)
    purpose = ws.cell(i, 3, p)
    small_cell(purpose)
    purpose.border = BORDER_THIN
    purpose.alignment = Alignment(vertical="center", wrap_text=True)

wb.defined_names.add(DefinedName("LibraryName",    attr_text="Settings!$B$4"))
wb.defined_names.add(DefinedName("CollectionName", attr_text="Settings!$B$5"))
wb.defined_names.add(DefinedName("DefaultLang",    attr_text="Settings!$B$6"))
wb.defined_names.add(DefinedName("LoanPeriod",     attr_text="Settings!$B$7"))
wb.defined_names.add(DefinedName("AccPrefix",      attr_text="Settings!$B$8"))
wb.defined_names.add(DefinedName("AuthorCodeLen",  attr_text="Settings!$B$9"))


# ═══════════════════════════════════════════════════════════════════════
# 2. LOOKUP_CODES
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Lookup_Codes")

vocab = {
    "A": ("Resource Type", [
        "Book", "Magazine", "Journal", "Periodical", "Newsletter", "Annual",
        "Special Issue", "Article", "Newspaper", "Newspaper Clipping",
        "Manuscript", "Personal Papers", "Pamphlet", "Brochure", "Paper Ephemera",
        "Image", "Photograph", "Poster", "Map",
        "Sound Recording", "Audio Recording", "AV Recording", "Video Recording",
        "Memento", "Sculpture", "Painting", "Textile", "Coin", "Medal",
        "Furniture",
        "Dataset", "Website", "Other",
    ]),
    "B": ("Genre", [
        "Literature (LIT)", "History (HIS)", "Philosophy (PHI)",
        "Science & Tech (SCI)", "Politics (POL)", "Essays (ESS)",
        "Biography / Memoir (BIO)", "Mystery / Thriller (THR)",
        "Religion & Mythology (REL)", "Art & Cinema (ART)", "Other",
    ]),
    "C": ("Language", ["English", "Bengali", "Hindi", "Other"]),
    "D": ("Format", [
        "PDF", "EPUB", "MOBI", "DJVU", "TXT", "DOCX", "HTML", "XML", "TEI XML",
        "Markdown", "RTF",
        "JPG", "JPEG", "PNG", "TIFF", "WebP", "SVG",
        "MP3", "WAV", "FLAC", "M4A", "AAC", "OGG",
        "MP4", "MKV", "MOV", "WebM", "AVI",
        "ZIP", "Other",
    ]),
    "E": ("OCR Status", [
        "Not Applicable", "Not Digitised", "Scanned (No OCR)", "OCR Pending",
        "OCR Complete", "Searchable (OCR)", "Image Only",
    ]),
    "F": ("Condition", [
        "New", "Excellent", "Good", "Fair", "Poor", "Damaged", "Under Repair",
    ]),
    "G": ("Circulation Status", [
        "Available", "On Loan", "Reference Only", "Lost", "Withdrawn",
    ]),
    "H": ("Location Type", [
        "Shelf", "Magazine Rack", "Journal Rack", "Archive", "Cabinet", "Box",
        "Display Case", "Wall", "Floor", "Other",
    ]),
    "I": ("Access Level", [
        "Public", "Private", "Restricted", "Research Only",
        "Permission Required",
    ]),
    "J": ("Rights", [
        "Copyright", "Public Domain", "Licensed", "Unknown", "User-Owned",
        "Institutional", "Other",
    ]),
    "K": ("Acquisition Type", [
        "Purchase", "Gift", "Donation", "Exchange", "Legacy", "Inherited",
        "Unknown",
    ]),
    "L": ("Frequency", [
        "Daily", "Weekly", "Fortnightly", "Monthly", "Bi-Monthly",
        "Quarterly", "Annual", "Irregular",
    ]),
    "M": ("Digitisation Status", [
        "Not Digitised", "Digitisation Planned", "Scanned", "OCR Pending",
        "OCR Complete", "OCR Verified",
    ]),
    "N": ("Physical/Digital", ["Physical", "Digital", "Both"]),
    "O": ("Membership Type", [
        "Student", "Faculty", "Public", "Institutional", "Other",
    ]),
    "P": ("Member Status", ["Active", "Inactive", "Suspended"]),
    "Q": ("Reading/Usage Status", ["Read", "Reading", "Unread", "N/A"]),
    "R": ("Preservation Status", ["Master", "Derivative", "Archived", "At Risk"]),
    "S": ("Collection", [
        "Books", "Manuscripts", "Photographs", "Audio", "Video",
        "Mementos", "Sculptures", "Paintings", "Textiles",
        "Coins & Medals", "Ephemera", "Newspapers", "Maps",
        "Furniture", "Other",
    ]),
    "T": ("Media Type", ["Image", "Audio", "Video", "Document", "Link", "None"]),
}

for col, (title, values) in vocab.items():
    ws.column_dimensions[col].width = max(24, len(title) + 6)
    c = ws[f"{col}1"]
    c.value = title
    hdr(c)
    for i, v in enumerate(values, start=2):
        cell = ws[f"{col}{i}"]
        cell.value = v
        cell.font = Font(name=FONT_BODY, size=SZ_BODY,
                         color=C_DARKSLATEGRAY)
        if i % 2 == 0:
            cell.fill = PatternFill("solid", fgColor=C_CREAM)

named = [
    ("ResourceTypeList",  "A", 40),
    ("GenreList",         "B", 15),
    ("LanguageList",      "C", 10),
    ("FormatList",        "D", 40),
    ("OCRStatusList",     "E", 10),
    ("ConditionList",     "F", 10),
    ("CirculationList",   "G", 10),
    ("LocationTypeList",  "H", 12),
    ("AccessLevelList",   "I", 10),
    ("RightsList",        "J", 10),
    ("AcqTypeList",       "K", 10),
    ("FrequencyList",     "L", 12),
    ("DigitisationList",  "M", 10),
    ("PhysDigiList",      "N", 5),
    ("MembershipList",    "O", 10),
    ("MemberStatusList",  "P", 6),
    ("ReadingStatusList", "Q", 6),
    ("PreservationList",  "R", 6),
    ("CollectionList",    "S", 18),
    ("MediaTypeList",     "T", 8),
]
for name, col, max_row in named:
    wb.defined_names.add(
        DefinedName(name, attr_text=f"Lookup_Codes!${col}$2:${col}${max_row}")
    )


# ═══════════════════════════════════════════════════════════════════════
# 3. CATALOG
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Catalog")
cat_cols = [
    ("Item ID", 8), ("Accession Number", 16), ("Call Number", 26),
    ("Title", 32), ("Subtitle", 22), ("Creator", 26),
    ("Author Code Override", 12), ("Author Code", 12),
    ("Editor", 20), ("Translator", 20),
    ("Type", 18), ("Genre", 18), ("Language", 12),
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
    hdr(ws.cell(1, i, name))
ws.row_dimensions[1].height = 34
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(cat_cols))}{ROWS_CAT + 1}"

for r in range(2, ROWS_CAT + 2):
    ws.cell(r, 1, f'=IF(ISBLANK(D{r}),"",ROW()-1)')
    gen_cell(ws.cell(r, 1))

    input_mono(ws.cell(r, 2))

    ws.cell(r, 3, (
        f'=IF(ISBLANK(D{r}),"",'
        f'MID(K{r},SEARCH("(",K{r})+1,SEARCH(")",K{r})-SEARCH("(",K{r})-1)&"-"&'
        f'MID(L{r},SEARCH("(",L{r})+1,SEARCH(")",L{r})-SEARCH("(",L{r})-1)&"-"&'
        f'H{r}&"-"&TEXT(A{r},"000")&'
        f'IF(ISBLANK(N{r}),"","-"&SUBSTITUTE(N{r}," ","")))'
    ))
    gen_cell(ws.cell(r, 3))

    for col in (4, 5, 6):
        input_cell(ws.cell(r, col))

    input_cell(ws.cell(r, 7))

    ws.cell(r, 8, (
        f'=IF(ISBLANK(F{r}),"",'
        f'IF(G{r}<>"",UPPER(TRIM(G{r})),'
        f'IF(ISNUMBER(SEARCH(" ",TRIM(F{r}))),'
        f'UPPER(LEFT(TRIM(RIGHT(SUBSTITUTE(TRIM(F{r})," ",REPT(" ",100)),100)),Settings!$B$9)),'
        f'UPPER(LEFT(TRIM(F{r}),Settings!$B$9)))))'
    ))
    gen_cell(ws.cell(r, 8))

    for col in (9, 10):
        input_cell(ws.cell(r, col))
    for col in (11, 12, 13):
        input_cell(ws.cell(r, col))
    for col in (14, 15, 16, 17, 18):
        input_cell(ws.cell(r, col))
    for col in (19, 20, 21, 22):
        input_cell(ws.cell(r, col))
    for col in (23, 24):
        input_cell(ws.cell(r, col))

    input_cell(ws.cell(r, 25))

    ws.cell(r, 26, (
        f'=IF(ISBLANK(D{r}),"",'
        f'IF(COUNTIFS(Circulation!$B$2:$B${ROWS_CIR},A{r},'
        f'Circulation!$H$2:$H${ROWS_CIR},"Active")>0,"On Loan","Available"))'
    ))
    gen_body(ws.cell(r, 26))

    for col in range(27, 34):
        input_cell(ws.cell(r, col))

for fml, rng in [
    ("=ResourceTypeList", f"K2:K{ROWS_CAT+1}"),
    ("=GenreList",        f"L2:L{ROWS_CAT+1}"),
    ("=LanguageList",     f"M2:M{ROWS_CAT+1}"),
    ("=LocationTypeList", f"W2:W{ROWS_CAT+1}"),
    ("=ConditionList",    f"Y2:Y{ROWS_CAT+1}"),
    ("=AcqTypeList",      f"AB2:AB{ROWS_CAT+1}"),
    ("=RightsList",       f"AF2:AF{ROWS_CAT+1}"),
]:
    dv = DataValidation(type="list", formula1=fml, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)

ws.conditional_formatting.add(
    f"D2:D{ROWS_CAT+1}",
    FormulaRule(formula=['ISBLANK($D2)'],
                fill=PatternFill("solid", fgColor=C_ERROR_SOFT)))
ws.conditional_formatting.add(
    f"F2:F{ROWS_CAT+1}",
    FormulaRule(formula=['AND($D2<>"",ISBLANK($F2))'],
                fill=PatternFill("solid", fgColor=C_WARN_SOFT)))
ws.conditional_formatting.add(
    f"Z2:Z{ROWS_CAT+1}",
    FormulaRule(formula=['$Z2="On Loan"'],
                fill=PatternFill("solid", fgColor=C_WARN_SOFT)))

if not BLANK_MODE:
    sample_catalog = [
        ("ACC-0001", "Rabindra Rachanabali", "Vol 1", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 1", "Shelf", "A2", "Good"),
        ("ACC-0002", "Rabindra Rachanabali", "Vol 2", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 2", "Shelf", "A2", "Good"),
        ("ACC-0003", "Rabindra Rachanabali", "Vol 3", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 3", "Shelf", "A2", "Good"),
        ("ACC-0004", "Rabindra Rachanabali", "Vol 4", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 4", "Shelf", "A2", "Good"),
        ("ACC-0005", "Rabindra Rachanabali", "Vol 5", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 5", "Shelf", "A2", "Good"),
        ("ACC-0006", "Rabindra Rachanabali", "Vol 6", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 6", "Shelf", "A2", "Good"),
        ("ACC-0007", "Rabindra Rachanabali", "Vol 7", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 7", "Shelf", "A2", "Good"),
        ("ACC-0008", "Rabindra Rachanabali", "Vol 8", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 8", "Shelf", "A2", "Good"),
        ("ACC-0009", "Rabindra Rachanabali", "Vol 9", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 9", "Shelf", "A1", "Good"),
        ("ACC-0010", "Rabindra Rachanabali", "Vol 10", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 10", "Shelf", "A1", "Good"),
        ("ACC-0011", "Rabindra Rachanabali", "Vol 11", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 11", "Shelf", "A1", "Good"),
        ("ACC-0012", "Rabindra Rachanabali", "Vol 12", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 12", "Shelf", "A1", "Good"),
        ("ACC-0013", "Rabindra Rachanabali", "Vol 13", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 13", "Shelf", "A1", "Good"),
        ("ACC-0014", "Rabindra Rachanabali", "Vol 14", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 14", "Shelf", "A1", "Good"),
        ("ACC-0015", "Rabindra Rachanabali", "Vol 15", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 15", "Shelf", "A1", "Good"),
        ("ACC-0016", "Rabindra Rachanabali", "Vol 16", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 16", "Shelf", "A1", "Good"),
        ("ACC-0017", "Rabindra Rachanabali", "Vol 17", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 17", "Shelf", "A1", "Good"),
        ("ACC-0018", "Rabindra Rachanabali", "Vol 18", "Rabindranath Tagore", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 18", "Shelf", "A1", "Good"),
        ("ACC-0019", "নীলকণ্ঠ পাখির খোঁজে", "Vol 1", "অতীন বন্দ্যোপাধ্যায়", "Fiction (FIC)", "Literature (LIT)", "Bengali", "Vol 1", "Shelf", "B1", "Good"),
        ("ACC-0020", "Desh", "", "", "Magazine", "Literature (LIT)", "Bengali", "", "Magazine Rack", "MR1", "Good"),
    ]
    for i, row in enumerate(sample_catalog):
        r = i + 2
        ws.cell(r, 2, row[0])
        ws.cell(r, 4, row[1])
        ws.cell(r, 5, row[2])
        ws.cell(r, 6, row[3])
        ws.cell(r, 11, row[4])
        ws.cell(r, 12, row[5])
        ws.cell(r, 13, row[6])
        ws.cell(r, 14, row[7])
        ws.cell(r, 23, row[8])
        ws.cell(r, 24, row[9])
        ws.cell(r, 25, row[10])


# ═══════════════════════════════════════════════════════════════════════
# 4. DIGITAL_CATALOG
# ═══════════════════════════════════════════════════════════════════════
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
ws.row_dimensions[1].height = 34
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(dig_cols))}{ROWS_DIG + 1}"

for r in range(2, ROWS_DIG + 2):
    ws.cell(r, 1, f'=IF(ISBLANK(C{r}),"",ROW()-1)')
    gen_cell(ws.cell(r, 1))

    ws.cell(r, 2, (
        f'=IF(ISBLANK(C{r}),"","DIG-"&'
        f'MID(F{r},SEARCH("(",F{r})+1,3)&"-"&'
        f'MID(G{r},SEARCH("(",G{r})+1,3)&"-"&'
        f'UPPER(LEFT(TRIM(RIGHT(SUBSTITUTE(TRIM(D{r})," ",REPT(" ",100)),100)),3))&"-"&'
        f'TEXT(A{r},"000"))'
    ))
    gen_cell(ws.cell(r, 2))

    for col in range(3, len(dig_cols) + 1):
        input_cell(ws.cell(r, col))
    gen_body(ws.cell(r, 18))

for fml, rng in [
    ("=ResourceTypeList",  f"F2:F{ROWS_DIG+1}"),
    ("=GenreList",         f"G2:G{ROWS_DIG+1}"),
    ("=LanguageList",      f"H2:H{ROWS_DIG+1}"),
    ("=FormatList",        f"I2:I{ROWS_DIG+1}"),
    ("=OCRStatusList",     f"O2:O{ROWS_DIG+1}"),
    ("=RightsList",        f"S2:S{ROWS_DIG+1}"),
    ("=AccessLevelList",   f"T2:T{ROWS_DIG+1}"),
    ("=ReadingStatusList", f"Z2:Z{ROWS_DIG+1}"),
    ("=PreservationList",  f"AA2:AA{ROWS_DIG+1}"),
]:
    dv = DataValidation(type="list", formula1=fml, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)

if not BLANK_MODE:
    dig_samples = [
        ("The Pocket Oracle", "Baltasar Gracian", "", "Book", "Philosophy (PHI)", "English",
         "PDF", "application/pdf", ".pdf", "", "Drive/Books/Philosophy", "",
         "Searchable (OCR)", "Yes", "2024-01-15", "2024-01-15", "Copyright", "Private",
         "A pocket edition of Gracian's aphorisms.", "", "", "", "", "Unread", "Master"),
        ("নীলকণ্ঠ পাখির খোঁজে", "অতীন বন্দ্যোপাধ্যায়", "", "Book", "Literature (LIT)", "Bengali",
         "PDF", "application/pdf", ".pdf", "", "Drive/Books/Bengali_Classics", "",
         "Scanned (No OCR)", "No", "2023-06-01", "2023-06-01", "Copyright", "Private",
         "Scanned copy of Bengali novel.", "", "", "", "", "Unread", "Master"),
        ("The Great Derangement", "Amitav Ghosh", "", "Book", "Essays (ESS)", "English",
         "EPUB", "application/epub+zip", ".epub", "", "Drive/Books/Essays", "",
         "Searchable (OCR)", "Yes", "2023-02-10", "2023-02-10", "Copyright", "Private",
         "Climate change essays.", "", "", "", "", "Read", "Master"),
        ("Bengali literary reading", "", "", "Audio Recording", "Literature (LIT)", "Bengali",
         "FLAC", "audio/flac", ".flac", "", "Drive/Archive/Audio", "",
         "Not Applicable", "No", "2023-08-12", "2023-08-12", "User-Owned", "Private",
         "Live reading of Tagore poems.", "", "", "", "", "N/A", "Master"),
        ("Interview with an author", "", "", "Video Recording", "Literature (LIT)", "Bengali",
         "MP4", "video/mp4", ".mp4", "", "Drive/Archive/Video", "",
         "Not Applicable", "No", "2024-03-01", "2024-03-01", "User-Owned", "Restricted",
         "Recorded interview.", "", "", "", "", "N/A", "Master"),
        ("Scanned Bengali manuscript", "", "", "Manuscript", "Literature (LIT)", "Bengali",
         "TIFF", "image/tiff", ".tiff", "", "Archive/Manuscripts", "",
         "Scanned (No OCR)", "No", "2022-01-01", "2022-01-01", "Public Domain",
         "Research Only", "High-res scan of 19th-c manuscript.", "", "", "ARC-0001", "",
         "N/A", "Master"),
    ]
    for i, row in enumerate(dig_samples):
        r = i + 2
        for j, val in enumerate(row, start=3):
            if j == 18:
                continue
            ws.cell(r, j, val)


# ═══════════════════════════════════════════════════════════════════════
# 5. ARCHIVE
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Archive")
arc_cols = [
    ("Archive ID", 12), ("Title", 32), ("Creator", 24), ("Date", 14),
    ("Description", 34), ("Resource Type", 20), ("Physical/Digital", 14),
    ("Language", 12), ("Provenance", 26), ("Acquisition", 22),
    ("Condition", 14), ("Location Type", 14), ("Location Identifier", 18),
    ("Digital Surrogate", 20), ("Digitisation Status", 18),
    ("OCR Status", 16), ("Rights", 16), ("Access Level", 16),
    ("Related Resource ID", 14), ("Notes", 30),
    ("Collection", 18), ("Dimensions", 22), ("Material", 22),
    ("Place", 18), ("Subject / Keywords", 26), ("Digital File Name", 22),
    ("Public Link", 40), ("Media Preview", 16),
]
for i, (n, w) in enumerate(arc_cols, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
    hdr(ws.cell(1, i, n))
ws.row_dimensions[1].height = 34
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(arc_cols))}{ROWS_ARC + 1}"

for r in range(2, ROWS_ARC + 2):
    ws.cell(r, 1, f'=IF(ISBLANK(B{r}),"","ARC-"&TEXT(ROW()-1,"0000"))')
    gen_cell(ws.cell(r, 1))
    for col in range(2, len(arc_cols) + 1):
        input_cell(ws.cell(r, col))
    ws.cell(r, 28, f'=IF(AA{r}="","",HYPERLINK(AA{r},"Open \u2197"))')
    gen_body(ws.cell(r, 28))

for fml, rng in [
    ("=ResourceTypeList", f"F2:F{ROWS_ARC+1}"),
    ("=PhysDigiList",     f"G2:G{ROWS_ARC+1}"),
    ("=LanguageList",     f"H2:H{ROWS_ARC+1}"),
    ("=ConditionList",    f"K2:K{ROWS_ARC+1}"),
    ("=LocationTypeList", f"L2:L{ROWS_ARC+1}"),
    ("=DigitisationList", f"O2:O{ROWS_ARC+1}"),
    ("=OCRStatusList",    f"P2:P{ROWS_ARC+1}"),
    ("=RightsList",       f"Q2:Q{ROWS_ARC+1}"),
    ("=AccessLevelList",  f"R2:R{ROWS_ARC+1}"),
    ("=CollectionList",   f"U2:U{ROWS_ARC+1}"),
]:
    dv = DataValidation(type="list", formula1=fml, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)

if not BLANK_MODE:
    arc_samples = [
        ("Historical newspaper clipping", "", "1947-08-15",
         "Bengali newspaper clipping from independence era.",
         "Newspaper Clipping", "Physical", "Bengali", "Family collection", "Inherited",
         "Fair", "Archive", "ARCH-01", "", "Not Digitised", "Not Applicable",
         "Unknown", "Private", "", "Fragile; store in acid-free sleeve.",
         "Newspapers", "18 × 12 cm", "newsprint", "Kolkata",
         "independence, 1947, Bengal", "", ""),
        ("Scanned Bengali manuscript", "", "1890",
         "Digitised manuscript with traditional binding.",
         "Manuscript", "Digital", "Bengali", "Purchased at auction", "Purchase",
         "Good", "Archive", "ARCH-02", "", "Scanned", "OCR Pending",
         "Public Domain", "Research Only", "", "High-res TIFF master.",
         "Manuscripts", "24 × 16 × 3 cm", "paper, ink, cloth",
         "Bengal", "manuscript, 19th century, Bengali",
         "manuscript_1890_scan.tif", ""),
        ("Brass Nataraja statuette", "", "c. 1950",
         "Small brass sculpture of Nataraja.",
         "Sculpture", "Physical", "Bengali", "Family heirloom", "Inherited",
         "Excellent", "Display Case", "CAB-01", "", "Digitisation Planned",
         "Not Applicable", "User-Owned", "Private",
         "", "Do not handle with bare hands.",
         "Sculptures", "12 × 8 × 8 cm", "brass", "South India",
         "Nataraja, Shiva, brass", "nataraja_front.jpg", ""),
        ("Grandmother's wedding saree", "", "c. 1930",
         "Handwoven silk saree with zari border.",
         "Textile", "Physical", "Bengali", "Family collection", "Inherited",
         "Fair", "Box", "BOX-07", "", "Digitisation Planned", "Not Applicable",
         "User-Owned", "Private", "", "Preserve in muslin wrap.",
         "Textiles", "5.5 × 1.2 m", "silk, zari", "Bengal",
         "saree, wedding, silk, zari", "saree_detail.jpg", ""),
        ("Field recording — Baul songs", "", "2019-11-20",
         "Audio field recording of Baul singers in Birbhum.",
         "Sound Recording", "Digital", "Bengali", "Own field research", "Recorded",
         "Excellent", "Archive", "ARCH-03", "", "Scanned", "Not Applicable",
         "User-Owned", "Research Only", "", "Recorded with Zoom H5.",
         "Audio", "", "digital (WAV 48 kHz)", "Birbhum, West Bengal",
         "Baul, folk music, oral tradition",
         "baul_session_2019.wav", ""),
        ("Interview with Shri Nabin Das", "", "2022-03-15",
         "Video interview about local printing history.",
         "Video Recording", "Digital", "Bengali", "Own research", "Recorded",
         "Excellent", "Archive", "ARCH-04", "", "Scanned", "Not Applicable",
         "User-Owned", "Research Only", "", "",
         "Video", "", "digital (MP4 H.264)", "Kolkata",
         "interview, printing, oral history",
         "interview_nabin_2022.mp4", ""),
    ]
    for i, row in enumerate(arc_samples):
        r = i + 2
        for j, val in enumerate(row, start=2):
            ws.cell(r, j, val)


# ═══════════════════════════════════════════════════════════════════════
# 6. CIRCULATION
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Circulation")
cir_cols = [
    ("Loan ID", 10), ("Item ID", 10), ("Member ID", 12),
    ("Borrower Name", 22), ("Checkout Date", 14), ("Due Date", 14),
    ("Return Date", 14), ("Status", 12), ("Notes", 30),
]
for i, (n, w) in enumerate(cir_cols, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
    hdr(ws.cell(1, i, n))
ws.row_dimensions[1].height = 30
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:I{ROWS_CIR + 1}"

for r in range(2, ROWS_CIR + 2):
    ws.cell(r, 1, f'=IF(ISBLANK(B{r}),"",ROW()-1)')
    gen_cell(ws.cell(r, 1))
    input_cell(ws.cell(r, 2))
    input_cell(ws.cell(r, 3))
    ws.cell(r, 4, f'=IFERROR(IF(ISBLANK(C{r}),"",VLOOKUP(C{r},Members!$A$2:$B${ROWS_MEM+1},2,FALSE)),"")')
    gen_body(ws.cell(r, 4))
    input_cell(ws.cell(r, 5))
    ws.cell(r, 6, f'=IF(ISBLANK(E{r}),"",E{r}+Settings!$B$7)')
    gen_cell(ws.cell(r, 6))
    input_cell(ws.cell(r, 7))
    ws.cell(r, 8, (
        f'=IF(ISBLANK(B{r}),"",'
        f'IF(NOT(ISBLANK(G{r})),"Returned",'
        f'IF(TODAY()>F{r},"Overdue","Active")))'
    ))
    gen_body(ws.cell(r, 8))
    input_cell(ws.cell(r, 9))


# ═══════════════════════════════════════════════════════════════════════
# 7. MEMBERS
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Members")
mem_cols = [
    ("Member ID", 12), ("Name", 26), ("Contact", 26),
    ("Membership Type", 18), ("Join Date", 14), ("Status", 12), ("Notes", 30),
]
for i, (n, w) in enumerate(mem_cols, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
    hdr(ws.cell(1, i, n))
ws.row_dimensions[1].height = 30
ws.freeze_panes = "A2"
for r in range(2, ROWS_MEM + 2):
    for col in range(1, 8):
        input_cell(ws.cell(r, col))

for fml, rng in [("=MembershipList", f"D2:D{ROWS_MEM+1}"),
                 ("=MemberStatusList", f"F2:F{ROWS_MEM+1}")]:
    dv = DataValidation(type="list", formula1=fml, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)

if not BLANK_MODE:
    ws.cell(2, 1, "M001")
    ws.cell(2, 2, "Sample Borrower")
    ws.cell(2, 4, "Public")
    ws.cell(2, 5, "2024-01-01")
    ws.cell(2, 6, "Active")


# ═══════════════════════════════════════════════════════════════════════
# 8. SPINE_LABELS
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Spine_Labels")
spn_cols = [
    ("Item ID", 10), ("Call Number", 28), ("Title", 34),
    ("Volume/Issue", 12), ("Creator", 24), ("Location", 20),
    ("Label Type", 18),
]
for i, (n, w) in enumerate(spn_cols, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
    hdr(ws.cell(1, i, n))
ws.row_dimensions[1].height = 30
ws.freeze_panes = "A2"

for r in range(2, ROWS_SPN + 2):
    ws.cell(r, 1, f'=Catalog!A{r}'); gen_cell(ws.cell(r, 1))
    ws.cell(r, 2, f'=Catalog!C{r}'); gen_cell(ws.cell(r, 2))
    ws.cell(r, 3, f'=Catalog!D{r}'); gen_body(ws.cell(r, 3))
    ws.cell(r, 4, f'=Catalog!N{r}'); gen_body(ws.cell(r, 4))
    ws.cell(r, 5, f'=Catalog!F{r}'); gen_body(ws.cell(r, 5))
    ws.cell(r, 6, f'=IF(Catalog!W{r}="","",Catalog!W{r}&" "&Catalog!X{r})')
    gen_body(ws.cell(r, 6))
    input_cell(ws.cell(r, 7))

dv = DataValidation(
    type="list",
    formula1='"Standard Card,Library Block,Minimal Spine"',
    allow_blank=True)
ws.add_data_validation(dv)
dv.add(f"G2:G{ROWS_SPN+1}")


# ═══════════════════════════════════════════════════════════════════════
# 9. DUBLIN_CORE
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Dublin_Core")
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 36
ws.column_dimensions["C"].width = 48

ws["A1"] = "DUBLIN CORE METADATA MAPPING"
title_cell(ws["A1"])
ws.merge_cells("A1:C1")
ws.row_dimensions[1].height = 32

for i, h in enumerate(["Dublin Core Element", "Catalog Field", "Notes"], start=1):
    hdr(ws.cell(3, i, h))

mapping = [
    ("title", "Title", "Standard DC element"),
    ("creator", "Creator", "Standard DC element"),
    ("contributor", "Editor, Translator", "Standard DC element; split"),
    ("publisher", "Publisher", "Standard DC element"),
    ("date", "Publication Year", "Standard DC element"),
    ("language", "Language", "Standard DC element"),
    ("type", "Type", "Standard DC element; project vocabulary"),
    ("format", "Format, Pages", "Standard DC element"),
    ("identifier", "ISBN, ISSN, Accession, Call Number", "Standard DC element"),
    ("subject", "Genre, Subject / Keywords", "Standard DC element"),
    ("description", "Notes, Description", "Standard DC element"),
    ("source", "Donor/Vendor, Acquisition Type, Provenance", "Standard DC element"),
    ("relation", "Volume, Series, Related ID", "isPartOf qualifier"),
    ("coverage", "Publication Place, Place", "spatial qualifier"),
    ("rights", "Rights", "Standard DC element"),
    ("—", "Item ID, Call Number, Author Code", "Application-specific"),
    ("—", "Location, Circulation Status, Condition", "Application-specific"),
    ("—", "Dimensions, Material, Collection", "Application-specific"),
    ("—", "Storage path, OCR status, Public Link, Checksum", "Application-specific"),
]
for i, row in enumerate(mapping, start=4):
    for j, val in enumerate(row, start=1):
        c = ws.cell(i, j, val)
        body_cell(c)
        c.border = BORDER_THIN
        if j == 1:
            c.font = Font(name=FONT_MONO, size=SZ_BODY,
                          color=C_DARKSLATEBLUE, bold=True)

ws["A26"] = "EXPORT VIEW — Dublin Core element columns"
h1_cell(ws["A26"])

dc_view = [
    "dc.title", "dc.creator", "dc.contributor", "dc.publisher", "dc.date",
    "dc.language", "dc.type", "dc.format", "dc.identifier", "dc.subject",
    "dc.description", "dc.source", "dc.relation", "dc.coverage", "dc.rights",
]
for j, name in enumerate(dc_view, start=1):
    c = ws.cell(28, j, name)
    c.font = Font(name=FONT_MONO, bold=True, color="FFFFFF",
                  size=SZ_SMALL)
    c.fill = PatternFill("solid", fgColor=C_DARKSLATEBLUE)
    c.alignment = Alignment(horizontal="center", vertical="center",
                            wrap_text=True)
    c.border = BORDER_THIN

for r in range(2, ROWS_CAT + 2):
    row = r + 27
    formulas = [
        f'=Catalog!D{r}',
        f'=Catalog!F{r}',
        f'=IF(Catalog!I{r}="","",Catalog!I{r})&IF(AND(Catalog!I{r}<>"",Catalog!J{r}<>""),"; ","")&IF(Catalog!J{r}="","",Catalog!J{r})',
        f'=Catalog!P{r}',
        f'=Catalog!R{r}',
        f'=Catalog!M{r}',
        f'=Catalog!K{r}',
        f'=IF(Catalog!S{r}<>"","ISBN:"&Catalog!S{r},IF(Catalog!T{r}<>"","ISSN:"&Catalog!T{r},""))',
        f'=Catalog!B{r}',
        f'=Catalog!L{r}',
        f'=IF(Catalog!AG{r}<>"",Catalog!AG{r},Catalog!AE{r})',
        f'=Catalog!AD{r}',
        f'=Catalog!V{r}',
        f'=Catalog!Q{r}',
        f'=Catalog!AF{r}',
    ]
    for j, fml in enumerate(formulas, start=1):
        cell = ws.cell(row, j, fml)
        gen_body(cell)


# ═══════════════════════════════════════════════════════════════════════
# 10. DATA_QUALITY
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Data_Quality")
ws.column_dimensions["A"].width = 38
ws.column_dimensions["B"].width = 14

ws["A1"] = "DATA QUALITY REPORT"
title_cell(ws["A1"])
ws.merge_cells("A1:B1")
ws.row_dimensions[1].height = 32

ws["A3"] = "Metric"
ws["B3"] = "Value"
hdr(ws["A3"])
hdr_brown(ws["B3"])

checks = [
    ("Total Catalog records",
     f'=COUNTA(Catalog!$D$2:$D${ROWS_CAT+1})'),
    ("Records missing Title",
     f'=COUNTA(Catalog!$A$2:$A${ROWS_CAT+1})-COUNTA(Catalog!$D$2:$D${ROWS_CAT+1})'),
    ("Records missing Creator",
     f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$F$2:$F${ROWS_CAT+1}=""))'),
    ("Records missing Type",
     f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$K$2:$K${ROWS_CAT+1}=""))'),
    ("Records missing Genre",
     f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$L$2:$L${ROWS_CAT+1}=""))'),
    ("Records missing Location Type",
     f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$W$2:$W${ROWS_CAT+1}=""))'),
    ("Records missing Location Id",
     f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$X$2:$X${ROWS_CAT+1}=""))'),
    ("Duplicate ISBNs",
     f'=SUMPRODUCT((Catalog!$S$2:$S${ROWS_CAT+1}<>"")*(COUNTIF(Catalog!$S$2:$S${ROWS_CAT+1},Catalog!$S$2:$S${ROWS_CAT+1})>1))/2'),
    ("Duplicate ISSNs",
     f'=SUMPRODUCT((Catalog!$T$2:$T${ROWS_CAT+1}<>"")*(COUNTIF(Catalog!$T$2:$T${ROWS_CAT+1},Catalog!$T$2:$T${ROWS_CAT+1})>1))/2'),
    ("Duplicate Accession Numbers",
     f'=SUMPRODUCT((Catalog!$B$2:$B${ROWS_CAT+1}<>"")*(COUNTIF(Catalog!$B$2:$B${ROWS_CAT+1},Catalog!$B$2:$B${ROWS_CAT+1})>1))/2'),
    ("Digital resources",
     f'=COUNTA(Digital_Catalog!$C$2:$C${ROWS_DIG+1})'),
    ("Archive objects",
     f'=COUNTA(Archive!$B$2:$B${ROWS_ARC+1})'),
    ("Archive missing Collection",
     f'=SUMPRODUCT((Archive!$B$2:$B${ROWS_ARC+1}<>"")*(Archive!$U$2:$U${ROWS_ARC+1}=""))'),
    ("Archive missing Dimensions",
     f'=SUMPRODUCT((Archive!$B$2:$B${ROWS_ARC+1}<>"")*(Archive!$V$2:$V${ROWS_ARC+1}=""))'),
    ("Archive missing Public Link",
     f'=SUMPRODUCT((Archive!$B$2:$B${ROWS_ARC+1}<>"")*(Archive!$AA$2:$AA${ROWS_ARC+1}=""))'),
    ("Active loans",
     f'=COUNTIF(Circulation!$H$2:$H${ROWS_CIR+1},"Active")'),
    ("Overdue loans",
     f'=COUNTIF(Circulation!$H$2:$H${ROWS_CIR+1},"Overdue")'),
]
for i, (label, f) in enumerate(checks, start=4):
    lbl = ws.cell(i, 1, label)
    lbl.font = Font(name=FONT_BODY, bold=True, size=SZ_BODY,
                    color=C_DARKSLATEGRAY)
    lbl.border = BORDER_THIN
    if i % 2 == 0:
        lbl.fill = PatternFill("solid", fgColor=C_LAVENDER)
    val = ws.cell(i, 2, f)
    val.font = Font(name=FONT_MONO, bold=True, size=SZ_BODY,
                    color=C_DARKSLATEBLUE)
    val.alignment = Alignment(horizontal="right", vertical="center")
    val.border = BORDER_THIN
    val.fill = PatternFill("solid", fgColor=C_AQUAMARINE)

ws["A24"] = "RECORD-LEVEL CHECKS"
h1_cell(ws["A24"])

for j, name in enumerate(["Item ID", "Title", "Issues"], start=1):
    hdr(ws.cell(25, j, name))

for r in range(2, ROWS_CAT + 2):
    row = r + 24
    c1 = ws.cell(row, 1, f'=Catalog!A{r}')
    c1.font = Font(name=FONT_MONO, size=SZ_SMALL, color=C_DARKSLATEGRAY)
    c2 = ws.cell(row, 2, f'=Catalog!D{r}')
    c2.font = Font(name=FONT_BODY, size=SZ_SMALL, color=C_DARKSLATEGRAY)
    c3 = ws.cell(row, 3, (
        f'=IF(Catalog!D{r}="","",'
        f'IF(ISBLANK(Catalog!F{r}),"Missing Creator; ","")&'
        f'IF(ISBLANK(Catalog!K{r}),"Missing Type; ","")&'
        f'IF(ISBLANK(Catalog!L{r}),"Missing Genre; ","")&'
        f'IF(ISBLANK(Catalog!W{r}),"Missing Location Type; ","")&'
        f'IF(ISBLANK(Catalog!X{r}),"Missing Location Id; ","")&'
        f'IF(AND(Catalog!S{r}<>"",COUNTIF(Catalog!$S$2:$S${ROWS_CAT+1},Catalog!S{r})>1),"Duplicate ISBN; ","")&'
        f'IF(AND(Catalog!B{r}<>"",COUNTIF(Catalog!$B$2:$B${ROWS_CAT+1},Catalog!B{r})>1),"Duplicate Accession; ",""))'
    ))
    c3.font = Font(name=FONT_SMALL, size=SZ_SMALL, color=C_BROWN)

ws.freeze_panes = "A26"


# ═══════════════════════════════════════════════════════════════════════
# 11. DASHBOARD
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Dashboard")
ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 36
ws.column_dimensions["C"].width = 16
ws.column_dimensions["D"].width = 4
ws.column_dimensions["E"].width = 36
ws.column_dimensions["F"].width = 16

ws["B2"] = "PERSONAL LIBRARY MANAGER & DIGITAL ARCHIVE"
title_cell(ws["B2"])
ws.merge_cells("B2:F2")
ws.row_dimensions[2].height = 36

ws["B3"] = "Dashboard"
ws["B3"].font = Font(name=FONT_SECTION, bold=True, size=SZ_H2,
                     color=C_BROWN)
ws.merge_cells("B3:F3")
ws.row_dimensions[3].height = 24


def dash_block(anchor_row, title, items, col_letter, brown=False):
    col_offset = ord(col_letter) - 64
    c = ws.cell(anchor_row, col_offset, title)
    if brown:
        section_banner_brown(c)
    else:
        section_banner(c)
    end_col = get_column_letter(col_offset + 1)
    ws.merge_cells(f"{col_letter}{anchor_row}:{end_col}{anchor_row}")
    ws.row_dimensions[anchor_row].height = 26
    for i, (label, f) in enumerate(items, start=1):
        r = anchor_row + i
        lbl = ws.cell(r, col_offset, label)
        metric_label(lbl)
        lbl.border = BORDER_THIN
        val = ws.cell(r, col_offset + 1, f)
        metric_value(val)
        val.border = BORDER_THIN


dash_block(5, "COLLECTION", [
    ("Physical items",    f'=COUNTA(Catalog!A2:A{ROWS_CAT+1})'),
    ("Digital resources", f'=COUNTA(Digital_Catalog!A2:A{ROWS_DIG+1})'),
    ("Archive objects",   f'=COUNTA(Archive!A2:A{ROWS_ARC+1})'),
    ("Books",             f'=COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Book")'),
    ("Magazines",         f'=COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Magazine")'),
    ("Journals",          f'=COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Journal")'),
    ("Other materials",
     f'=COUNTA(Catalog!A2:A{ROWS_CAT+1})-COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Book")-COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Magazine")-COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Journal")'),
], "B")

dash_block(5, "METADATA", [
    ("Missing titles",
     f'=COUNTA(Catalog!A2:A{ROWS_CAT+1})-COUNTA(Catalog!D2:D{ROWS_CAT+1})'),
    ("Missing creators",
     f'=SUMPRODUCT((Catalog!D2:D{ROWS_CAT+1}<>"")*(Catalog!F2:F{ROWS_CAT+1}=""))'),
    ("Missing locations",
     f'=SUMPRODUCT((Catalog!D2:D{ROWS_CAT+1}<>"")*(Catalog!W2:W{ROWS_CAT+1}=""))'),
    ("Duplicate ISBNs",
     f'=SUMPRODUCT((Catalog!S2:S{ROWS_CAT+1}<>"")*(COUNTIF(Catalog!S2:S{ROWS_CAT+1},Catalog!S2:S{ROWS_CAT+1})>1))/2'),
    ("Duplicate Accessions",
     f'=SUMPRODUCT((Catalog!B2:B{ROWS_CAT+1}<>"")*(COUNTIF(Catalog!B2:B{ROWS_CAT+1},Catalog!B2:B{ROWS_CAT+1})>1))/2'),
], "E", brown=True)

dash_block(15, "READING STATUS", [
    ("Read",    f'=COUNTIF(Digital_Catalog!Z2:Z{ROWS_DIG+1},"Read")'),
    ("Reading", f'=COUNTIF(Digital_Catalog!Z2:Z{ROWS_DIG+1},"Reading")'),
    ("Unread",  f'=COUNTIF(Digital_Catalog!Z2:Z{ROWS_DIG+1},"Unread")'),
], "B")

dash_block(15, "CIRCULATION", [
    ("Available", f'=COUNTIF(Catalog!Z2:Z{ROWS_CAT+1},"Available")'),
    ("On Loan",   f'=COUNTIF(Catalog!Z2:Z{ROWS_CAT+1},"On Loan")'),
    ("Overdue",   f'=COUNTIF(Circulation!H2:H{ROWS_CIR+1},"Overdue")'),
], "E", brown=True)

dash_block(20, "ARCHIVE BY COLLECTION", [
    ("Manuscripts", f'=COUNTIF(Archive!U2:U{ROWS_ARC+1},"Manuscripts")'),
    ("Photographs", f'=COUNTIF(Archive!U2:U{ROWS_ARC+1},"Photographs")'),
    ("Audio",       f'=COUNTIF(Archive!U2:U{ROWS_ARC+1},"Audio")'),
    ("Video",       f'=COUNTIF(Archive!U2:U{ROWS_ARC+1},"Video")'),
    ("Sculptures",  f'=COUNTIF(Archive!U2:U{ROWS_ARC+1},"Sculptures")'),
    ("Textiles",    f'=COUNTIF(Archive!U2:U{ROWS_ARC+1},"Textiles")'),
    ("Other",
     f'=COUNTA(Archive!B2:B{ROWS_ARC+1})-SUM(COUNTIF(Archive!U2:U{ROWS_ARC+1},{{"Manuscripts","Photographs","Audio","Video","Sculptures","Textiles"}}))'),
], "B")

dash_block(20, "DIGITISATION", [
    ("Physical archive only", f'=COUNTIF(Archive!G2:G{ROWS_ARC+1},"Physical")'),
    ("Digital archive only",  f'=COUNTIF(Archive!G2:G{ROWS_ARC+1},"Digital")'),
    ("Both",                  f'=COUNTIF(Archive!G2:G{ROWS_ARC+1},"Both")'),
    ("OCR Pending",           f'=COUNTIF(Archive!P2:P{ROWS_ARC+1},"OCR Pending")'),
    ("OCR Complete",          f'=COUNTIF(Archive!P2:P{ROWS_ARC+1},"OCR Complete")'),
], "E", brown=True)


# ═══════════════════════════════════════════════════════════════════════
# 12. JSON_EXPORT
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("JSON_Export")
ws.column_dimensions["A"].width = 10
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 24
ws.column_dimensions["D"].width = 70

ws["A1"] = "STRUCTURED EXPORT FOR JSON / CSV PIPELINE"
title_cell(ws["A1"])
ws.merge_cells("A1:D1")
ws.row_dimensions[1].height = 32

ws["A3"] = ("Flattened view of the Catalog for conversion to JSON by a "
            "future Python or Node script. Field names use snake_case.")
small_cell(ws["A3"])
ws.merge_cells("A3:D3")

for j, name in enumerate(["item_id", "title", "creator", "json_snippet"],
                          start=1):
    c = ws.cell(5, j, name)
    c.font = Font(name=FONT_MONO, bold=True, color="FFFFFF",
                  size=SZ_SMALL)
    c.fill = PatternFill("solid", fgColor=C_DARKSLATEBLUE)
    c.alignment = Alignment(horizontal="center", vertical="center",
                            wrap_text=True)
    c.border = BORDER_THIN

for r in range(2, ROWS_CAT + 2):
    row = r + 4
    c1 = ws.cell(row, 1, f'=Catalog!A{r}')
    c1.font = Font(name=FONT_MONO, size=SZ_SMALL, color=C_DARKSLATEGRAY)
    c2 = ws.cell(row, 2, f'=Catalog!D{r}')
    c2.font = Font(name=FONT_BODY, size=SZ_SMALL, color=C_DARKSLATEGRAY)
    c3 = ws.cell(row, 3, f'=Catalog!F{r}')
    c3.font = Font(name=FONT_BODY, size=SZ_SMALL, color=C_DARKSLATEGRAY)
    c4 = ws.cell(row, 4, (
        f'=IF(Catalog!D{r}="","",'
        f'"{{""item_id"":"&Catalog!A{r}&",'
        f'""title"":"""&SUBSTITUTE(Catalog!D{r},"""","\\""")&""",'
        f'""creator"":"""&SUBSTITUTE(Catalog!F{r},"""","\\""")&""",'
        f'""language"":"""&Catalog!M{r}&""",'
        f'""identifier"":"""&Catalog!C{r}&"""}}"'
        f')'
    ))
    c4.font = Font(name=FONT_MONO, size=SZ_SMALL, color=C_DARKSLATEBLUE)
    c4.fill = PatternFill("solid", fgColor=C_GENERATED)

ws.freeze_panes = "A6"


# ═══════════════════════════════════════════════════════════════════════
# 13. README
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("README", 0)
ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 108

readme_lines = [
    ("TITLE", "Personal Library Manager & Digital Archive"),
    ("H2",    "Version 2.1 — Excel Application"),
    ("P",     "The Excel workbook is the primary structured-data and metadata layer. "
              "The web application is a separate presentation and printing layer."),

    ("H3",    "Quick Start"),
    ("P",     "1. Open Settings and set your library name, loan period, and defaults."),
    ("P",     "2. Use Catalog for physical books, magazines, and journals. Yellow cells "
              "are editable; pale-aqua cells auto-generate."),
    ("P",     "3. Use Digital_Catalog for PDFs, images, audio, video, datasets, and other "
              "digital resources."),
    ("P",     "4. Use Archive for manuscripts, photographs, sound recordings, videos, "
              "sculptures, textiles, mementos, and other objects."),
    ("P",     "5. Record loans in Circulation. The Catalog's Circulation Status updates "
              "automatically."),
    ("P",     "6. Monitor Dashboard and Data_Quality for completeness and problems."),

    ("H3",    "Worksheet Guide"),
    ("P",     "README          — this page"),
    ("P",     "Dashboard       — collection, metadata, reading, circulation, archive metrics"),
    ("P",     "Catalog         — physical books, magazines, journals, printed materials"),
    ("P",     "Digital_Catalog — digital resources of all formats"),
    ("P",     "Archive         — archival objects, both physical and digital"),
    ("P",     "Circulation     — loans, borrowers, due dates, returns"),
    ("P",     "Members         — borrower registry"),
    ("P",     "Spine_Labels    — derived label data for printing"),
    ("P",     "Dublin_Core     — metadata mapping and export view"),
    ("P",     "Lookup_Codes    — controlled vocabularies for dropdowns"),
    ("P",     "Data_Quality    — validation and completeness checks"),
    ("P",     "Settings        — configurable parameters"),
    ("P",     "JSON_Export     — flattened view for JSON conversion"),

    ("H3",    "Typography"),
    ("P",     "Titles       — Garamond"),
    ("P",     "Section headers — Baskerville Old Face"),
    ("P",     "Body text    — Book Antiqua (Caslon substitute)"),
    ("P",     "Call numbers, IDs — Courier New"),

    ("H3",    "Colour System"),
    ("P",     "DarkSlateBlue (#483D8B) — primary headers, generated values"),
    ("P",     "Brown (#A52A2A)         — secondary headers, alternates"),
    ("P",     "Aquamarine (#7FFFD4)    — metric values, highlights"),
    ("P",     "Cream (#FFF8E7)         — editable cells"),
    ("P",     "Lavender (#E6E6FA)      — label backgrounds"),

    ("H3",    "Author Code Logic"),
    ("P",     "Author Code is derived from the last word of the Creator field. "
              "'Rabindranath Tagore' → TAG. Bengali and other Unicode scripts are preserved."),
    ("P",     "Use Author Code Override to supply a manual code; it takes precedence."),

    ("H3",    "Call Number System"),
    ("P",     "Format: TYPE-GENRE-AUTHORCODE-ITEM-VOLUME. Example: FIC-LIT-TAG-001-Vol1."),
    ("P",     "This is the project's own identifier scheme. It is not Dewey, LC, or Cutter."),

    ("H3",    "Location System"),
    ("P",     "Location Type is controlled (Shelf, Magazine Rack, Journal Rack, Archive, "
              "Cabinet, Box, Display Case, Wall, Floor, Other)."),
    ("P",     "Location Identifier is free text (A1, A2, MR1, JR1, ARCH-01, BOX-07)."),

    ("H3",    "Archive Sheet — Extended Fields"),
    ("P",     "Collection         — Books, Manuscripts, Photographs, Audio, Video, "
              "Mementos, Sculptures, Paintings, Textiles, Coins & Medals, Ephemera, "
              "Newspapers, Maps, Furniture, Other"),
    ("P",     "Dimensions         — physical size, e.g. '24 × 18 cm' or '12 × 8 × 3 in'"),
    ("P",     "Material           — e.g. 'paper, ink', 'oil on canvas', 'brass'"),
    ("P",     "Place              — geographic origin"),
    ("P",     "Subject / Keywords — comma-separated tags"),
    ("P",     "Digital File Name  — e.g. 'manuscript_1890_scan.tif'"),
    ("P",     "Public Link        — Google Drive, Dropbox, or any public URL"),
    ("P",     "Media Preview      — auto-generated hyperlink"),

    ("H3",    "Dublin Core"),
    ("P",     "Dublin Core is the primary metadata framework. See the Dublin_Core sheet."),
    ("P",     "Project-specific fields are documented as application extensions."),

    ("H3",    "Interoperability"),
    ("P",     "CSV — every data sheet exports cleanly with UTF-8 encoding."),
    ("P",     "JSON — see JSON_Export sheet."),
    ("P",     "TEI — catalog preserves structured fields suitable for TEI transformation."),
    ("P",     "Bibliographic APIs — ISBN, ISSN, DOI, OCLC fields ready to receive enrichment."),

    ("H3",    "Protection"),
    ("P",     "Generated columns are locked. Input columns are editable."),
    ("P",     "Sheet protection password: 'library'. Change in Settings if desired."),

    ("H3",    "Important Notes"),
    ("P",     "• Accession Number is the permanent identifier. Item ID is internal."),
    ("P",     "• Do not sort the Catalog sheet by columns other than Item ID — Call Number "
              "formulas reference row position. Use AutoFilter instead."),
    ("P",     "• Open the workbook in Excel and press Ctrl+S before uploading to the web "
              "app so formula values are cached."),
    ("P",     "• No VBA macros are used. Compatible with Microsoft Excel, Google Sheets, "
              "and LibreOffice Calc."),
]

row = 1
for kind, text in readme_lines:
    c = ws.cell(row, 2, text)
    if kind == "TITLE":
        c.font = Font(name=FONT_TITLE, bold=True, size=SZ_TITLE,
                      color=C_DARKSLATEBLUE)
    elif kind == "H2":
        c.font = Font(name=FONT_SECTION, bold=True, size=SZ_H2,
                      color=C_BROWN)
    elif kind == "H3":
        c.font = Font(name=FONT_SECTION, bold=True, size=SZ_H2,
                      color=C_DARKSLATEBLUE)
    else:
        c.font = Font(name=FONT_BODY, size=SZ_BODY,
                      color=C_DARKSLATEGRAY)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if kind == "P":
        ws.row_dimensions[row].height = 26
    if kind in ("TITLE",):
        ws.row_dimensions[row].height = 36
    if kind in ("H2", "H3"):
        ws.row_dimensions[row].height = 24
    row += 1


# ═══════════════════════════════════════════════════════════════════════
# SHEET PROTECTION
# ═══════════════════════════════════════════════════════════════════════
PW = "library"
for sname in ["Catalog", "Digital_Catalog", "Archive", "Circulation",
              "Members", "Spine_Labels", "Dublin_Core", "Data_Quality",
              "Dashboard", "JSON_Export"]:
    wb[sname].protection.sheet = True
    wb[sname].protection.password = PW

wb["Settings"].protection.sheet = True
wb["Settings"].protection.password = PW


# ═══════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════
wb.active = 0

if BLANK_MODE:
    out = "Personal Library Template.xlsx"
else:
    out = "Personal Library Manager & Digital Archive.xlsx"

wb.save(out)

print()
print("=" * 68)
print(f"Workbook written: {out}")
print(f"Mode:  {'BLANK TEMPLATE' if BLANK_MODE else 'FULL (with sample data)'}")
print(f"Sheets ({len(wb.sheetnames)}):")
for s in wb.sheetnames:
    print(f"  • {s}")
print()
print("Typography: Garamond · Baskerville Old Face · Book Antiqua · Courier New")
print("Colours:    DarkSlateBlue · Brown · Aquamarine · Cream · Lavender")
print()
print("Next: open the workbook in Excel and press Ctrl+S to cache formula values.")
print("=" * 68)