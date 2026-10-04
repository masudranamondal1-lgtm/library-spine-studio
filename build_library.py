"""
Personal Library Manager & Digital Archive — Workbook Generator
================================================================

Three separate Item Type vocabularies:
    Catalog         — books and print (Book, Magazine, Journal, …)
    Digital_Catalog — digital resources (E-Book, Article, Image, Audio, …)
    Archive         — archival objects  (Sculpture, Coin, Photograph, …)

Typographic system:
    Garamond            — workbook title, dashboard banner
    Baskerville Old Face — section headers, column headers
    Book Antiqua        — body text, data (Caslon substitute)
    Courier New         — call numbers, IDs, codes, formulas

Colour system:
    DarkSlateBlue (#483D8B)  — primary headers, generated values
    Brown        (#A52A2A)  — secondary headers, alternates
    Aquamarine   (#7FFFD4)  — metric values, highlights
    Cream        (#FFF8E7)  — editable cells
    Lavender     (#E6E6FA)  — label backgrounds

Usage:
    python build_library.py            → full workbook with sample data
    python build_library.py --blank    → blank template

Requires: openpyxl
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
FONT_TITLE   = "Garamond"
FONT_SECTION = "Baskerville Old Face"
FONT_BODY    = "Book Antiqua"
FONT_MONO    = "Courier New"
FONT_SMALL   = "Book Antiqua"

SZ_TITLE = 22
SZ_H1    = 16
SZ_H2    = 13
SZ_BODY  = 12
SZ_TABLE = 11
SZ_SMALL = 10
SZ_META  = 9


# ═══════════════════════════════════════════════════════════════════════
# COLOUR PALETTE
# ═══════════════════════════════════════════════════════════════════════
C_DARKSLATEBLUE = "483D8B"
C_BROWN         = "A52A2A"
C_AQUAMARINE    = "7FFFD4"
C_CREAM         = "FFF8E7"
C_PALETURQ      = "AFEEEE"
C_LAVENDER      = "E6E6FA"
C_PERU          = "CD853F"
C_DARKSLATEGRAY = "2F4F4F"
C_ERROR_SOFT    = "FFE4E1"
C_WARN_SOFT     = "FFFACD"
C_GENERATED     = "E6F2F2"
C_BORDER        = "B8B8C8"

thin   = Side(style="thin",   color=C_BORDER)
medium = Side(style="medium", color=C_DARKSLATEBLUE)
BORDER_THIN   = Border(left=thin, right=thin, top=thin, bottom=thin)
BORDER_HEADER = Border(left=thin, right=thin, top=thin, bottom=medium)


# ═══════════════════════════════════════════════════════════════════════
# STYLE HELPERS
# ═══════════════════════════════════════════════════════════════════════
def hdr(cell):
    cell.font = Font(name=FONT_SECTION, bold=True, color="FFFFFF", size=SZ_TABLE)
    cell.fill = PatternFill("solid", fgColor=C_DARKSLATEBLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER_HEADER


def hdr_brown(cell):
    cell.font = Font(name=FONT_SECTION, bold=True, color="FFFFFF", size=SZ_TABLE)
    cell.fill = PatternFill("solid", fgColor=C_BROWN)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER_HEADER


def input_cell(cell):
    cell.fill = PatternFill("solid", fgColor=C_CREAM)
    cell.protection = Protection(locked=False)
    cell.border = BORDER_THIN
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    cell.font = Font(name=FONT_BODY, size=SZ_BODY, color=C_DARKSLATEGRAY)


def input_mono(cell):
    input_cell(cell)
    cell.font = Font(name=FONT_MONO, size=SZ_BODY, color=C_DARKSLATEGRAY)


def gen_cell(cell):
    cell.fill = PatternFill("solid", fgColor=C_GENERATED)
    cell.protection = Protection(locked=True)
    cell.border = BORDER_THIN
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    cell.font = Font(name=FONT_MONO, size=SZ_BODY, color=C_DARKSLATEBLUE)


def gen_body(cell):
    cell.fill = PatternFill("solid", fgColor=C_GENERATED)
    cell.protection = Protection(locked=True)
    cell.border = BORDER_THIN
    cell.alignment = Alignment(vertical="center", wrap_text=True)
    cell.font = Font(name=FONT_BODY, size=SZ_BODY, color=C_DARKSLATEGRAY)


def title_cell(cell):
    cell.font = Font(name=FONT_TITLE, bold=True, size=SZ_TITLE, color=C_DARKSLATEBLUE)


def h1_cell(cell):
    cell.font = Font(name=FONT_TITLE, bold=True, size=SZ_H1, color=C_DARKSLATEBLUE)


def h3_cell(cell):
    cell.font = Font(name=FONT_SECTION, bold=True, size=SZ_H2, color=C_DARKSLATEBLUE)


def body_cell(cell):
    cell.font = Font(name=FONT_BODY, size=SZ_BODY, color=C_DARKSLATEGRAY)


def small_cell(cell):
    cell.font = Font(name=FONT_SMALL, italic=True, size=SZ_SMALL, color=C_PERU)


def section_banner(cell):
    cell.font = Font(name=FONT_SECTION, bold=True, size=SZ_H2, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=C_DARKSLATEBLUE)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)


def section_banner_brown(cell):
    cell.font = Font(name=FONT_SECTION, bold=True, size=SZ_H2, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor=C_BROWN)
    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)


def metric_label(cell):
    cell.font = Font(name=FONT_BODY, size=SZ_BODY, color=C_DARKSLATEGRAY)
    cell.fill = PatternFill("solid", fgColor=C_CREAM)


def metric_value(cell):
    cell.font = Font(name=FONT_MONO, bold=True, size=SZ_BODY, color=C_DARKSLATEBLUE)
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

ws["A3"] = "Setting"; ws["B3"] = "Value"; ws["C3"] = "Purpose"
hdr(ws["A3"]); hdr(ws["B3"]); hdr_brown(ws["C3"])

settings_data = [
    ("Library Name",                   "Personal Library",  "Display name of the library"),
    ("Collection Name",                "Main Collection",   "Name of the primary collection"),
    ("Default Language",               "English",           "Default language for new records"),
    ("Default Loan Period (days)",     14,                  "Circulation loan period in days"),
    ("Accession Prefix",               "ACC",               "Prefix for accession numbers"),
    ("Author Code Length",             3,                   "Letters used for auto author code"),
    ("Date Format",                    "YYYY-MM-DD",        "Display format for dates"),
    ("Default Resource Type",          "Book",              "Default resource type"),
    ("Metadata Profile Version",       "3.1",               "Dublin Core profile version"),
    ("Workbook Protection Password",   "library",           "Password to unprotect sheets"),
]
for i, (k, v, p) in enumerate(settings_data, start=4):
    lbl = ws.cell(i, 1, k)
    lbl.font = Font(name=FONT_BODY, bold=True, size=SZ_BODY, color=C_DARKSLATEBLUE)
    lbl.fill = PatternFill("solid", fgColor=C_LAVENDER)
    lbl.border = BORDER_THIN
    val = ws.cell(i, 2, v); input_cell(val)
    pu = ws.cell(i, 3, p); small_cell(pu); pu.border = BORDER_THIN
    pu.alignment = Alignment(vertical="center", wrap_text=True)

wb.defined_names.add(DefinedName("LibraryName",    attr_text="Settings!$B$4"))
wb.defined_names.add(DefinedName("CollectionName", attr_text="Settings!$B$5"))
wb.defined_names.add(DefinedName("DefaultLang",    attr_text="Settings!$B$6"))
wb.defined_names.add(DefinedName("LoanPeriod",     attr_text="Settings!$B$7"))
wb.defined_names.add(DefinedName("AccPrefix",      attr_text="Settings!$B$8"))
wb.defined_names.add(DefinedName("AuthorCodeLen",  attr_text="Settings!$B$9"))


# ═══════════════════════════════════════════════════════════════════════
# 2. LOOKUP_CODES — organised by category
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Lookup_Codes")

# ─── Three Item Type code tables ─────────────────────────────────────
# A:B  → Catalog Item Type (books and print)
# C:D  → Digital Item Type (digital resources)
# E:F  → Archive Item Type (archival objects)
# G:H  → Subject (all sheets)
# I:J  → Genre / Form (all sheets)
# K    → Language
# L    → Format
# M    → OCR Status
# N    → Condition
# O    → Location Type
# P    → Access Level
# Q    → Rights
# R    → Acquisition Type
# S    → Physical/Digital
# T    → Membership Type
# U    → Member Status
# V    → Reading/Usage Status
# W    → Preservation Status
# X    → Collection (archive)
# Y    → Media Type
# Z    → Circulation Status
# AA   → Frequency
# AB   → Digitisation Status

cat_item_types = [
    ("Book",            "BK"),
    ("Magazine",        "MG"),
    ("Journal",         "JN"),
    ("Periodical",      "PR"),
    ("Newspaper",       "NP"),
    ("Newsletter",      "NL"),
    ("Annual",          "AN"),
    ("Special Issue",   "SI"),
    ("Pamphlet",        "PM"),
    ("Brochure",        "BR"),
    ("Manuscript",      "MS"),
    ("Personal Papers", "PP"),
    ("Ephemera",        "EP"),
    ("Other",           "OT"),
]

dig_item_types = [
    ("E-Book",           "EB"),
    ("E-Journal",        "EJ"),
    ("E-Magazine",       "EM"),
    ("Article",          "AR"),
    ("Image",            "IM"),
    ("Photograph",       "PH"),
    ("Poster",           "PO"),
    ("Map",              "MP"),
    ("Manuscript",       "MS"),
    ("Sound Recording",  "SR"),
    ("Audio Recording",  "AU"),
    ("Video Recording",  "VD"),
    ("AV Recording",     "AV"),
    ("Dataset",          "DS"),
    ("Website",          "WB"),
    ("Software",         "SW"),
    ("Other",            "OT"),
]

arc_item_types = [
    ("Manuscript",          "MS"),
    ("Newspaper",           "NP"),
    ("Newspaper Clipping",  "NC"),
    ("Photograph",          "PH"),
    ("Poster",              "PO"),
    ("Map",                 "MP"),
    ("Painting",            "PT"),
    ("Sculpture",           "SC"),
    ("Textile",             "TX"),
    ("Coin",                "CN"),
    ("Medal",               "MD"),
    ("Furniture",           "FR"),
    ("Memento",             "MN"),
    ("Personal Papers",     "PP"),
    ("Ephemera",            "EP"),
    ("Pamphlet",            "PM"),
    ("Brochure",            "BR"),
    ("Sound Recording",     "SR"),
    ("Audio Recording",     "AU"),
    ("Video Recording",     "VD"),
    ("AV Recording",        "AV"),
    ("Other",               "OT"),
]

subjects = [
    ("Literature",             "LIT"),
    ("History",                "HIS"),
    ("Philosophy",             "PHI"),
    ("Science & Technology",   "SCI"),
    ("Politics",               "POL"),
    ("Economics",              "ECO"),
    ("Religion & Mythology",   "REL"),
    ("Art & Cinema",           "ART"),
    ("Music",                  "MUS"),
    ("Language",               "LAN"),
    ("Medicine",               "MED"),
    ("Law",                    "LAW"),
    ("Education",              "EDU"),
    ("Biography",              "BIO"),
    ("Reference",              "REF"),
    ("Other",                  "OTH"),
]

genres = [
    ("Novel",             "NOV"),
    ("Poem",              "POE"),
    ("Play",              "PLY"),
    ("Short Story",       "SHT"),
    ("Essay",             "ESS"),
    ("Article",           "ART"),
    ("Biography",         "BIO"),
    ("Autobiography",     "AUT"),
    ("Travelogue",        "TRV"),
    ("Science Fiction",   "SCF"),
    ("Fantasy",           "FAN"),
    ("Mystery",           "MYS"),
    ("Thriller",          "THR"),
    ("Romance",           "ROM"),
    ("Children's",        "CHI"),
    ("Reference",         "REF"),
    ("Collected Works",   "COL"),
    ("Anthology",         "ANT"),
    ("Interview",         "INT"),
    ("Other",             "OTH"),
]

languages = ["English", "Bengali", "Hindi", "Sanskrit", "Other"]

formats = [
    "PDF", "EPUB", "MOBI", "DJVU", "TXT", "DOCX", "HTML", "XML", "TEI XML",
    "Markdown", "RTF",
    "JPG", "JPEG", "PNG", "TIFF", "WebP", "SVG",
    "MP3", "WAV", "FLAC", "M4A", "AAC", "OGG",
    "MP4", "MKV", "MOV", "WebM", "AVI",
    "ZIP", "Other",
]

ocr_status = [
    "Not Applicable", "Not Digitised", "Scanned (No OCR)", "OCR Pending",
    "OCR Complete", "Searchable (OCR)", "Image Only",
]

conditions = ["New", "Excellent", "Good", "Fair", "Poor", "Damaged", "Under Repair"]

location_types = [
    "Shelf", "Magazine Rack", "Journal Rack", "Archive", "Cabinet", "Box",
    "Display Case", "Wall", "Floor", "Other",
]

access_levels = ["Public", "Private", "Restricted", "Research Only", "Permission Required"]

rights_vals = ["Copyright", "Public Domain", "Licensed", "Unknown",
               "User-Owned", "Institutional", "Other"]

acq_types = ["Purchase", "Gift", "Donation", "Exchange", "Legacy", "Inherited", "Unknown"]

phys_digi = ["Physical", "Digital", "Both"]

membership_types = ["Student", "Faculty", "Public", "Institutional", "Other"]

member_status = ["Active", "Inactive", "Suspended"]

reading_status = ["Read", "Reading", "Unread", "N/A"]

preservation_status = ["Master", "Derivative", "Archived", "At Risk"]

collection_vals = [
    "Books", "Manuscripts", "Photographs", "Audio", "Video",
    "Mementos", "Sculptures", "Paintings", "Textiles",
    "Coins & Medals", "Ephemera", "Newspapers", "Maps",
    "Furniture", "Other",
]

media_types = ["Image", "Audio", "Video", "Document", "Link", "None"]

circ_status = ["Available", "On Loan", "Reference Only", "Lost", "Withdrawn"]

frequency_vals = ["Daily", "Weekly", "Fortnightly", "Monthly", "Bi-Monthly",
                  "Quarterly", "Annual", "Irregular"]

digitisation_status = ["Not Digitised", "Digitisation Planned", "Scanned",
                       "OCR Pending", "OCR Complete", "OCR Verified"]


def write_coded_table(start_col_idx, title, rows, brown_code=False):
    col1 = get_column_letter(start_col_idx)
    col2 = get_column_letter(start_col_idx + 1)
    ws.column_dimensions[col1].width = 24
    ws.column_dimensions[col2].width = 8
    c1 = ws[f"{col1}1"]; c1.value = title + " (Value)"; hdr(c1)
    c2 = ws[f"{col2}1"]; c2.value = "Code"
    (hdr_brown if brown_code else hdr)(c2)
    for i, (val, code) in enumerate(rows, start=2):
        a = ws[f"{col1}{i}"]; a.value = val
        a.font = Font(name=FONT_BODY, size=SZ_BODY, color=C_DARKSLATEGRAY)
        if i % 2 == 0: a.fill = PatternFill("solid", fgColor=C_CREAM)
        b = ws[f"{col2}{i}"]; b.value = code
        b.font = Font(name=FONT_MONO, size=SZ_BODY, color=C_DARKSLATEBLUE, bold=True)
        if i % 2 == 0: b.fill = PatternFill("solid", fgColor=C_CREAM)


def write_plain_table(col_idx, title, values):
    col = get_column_letter(col_idx)
    ws.column_dimensions[col].width = max(22, len(title) + 4)
    c = ws[f"{col}1"]; c.value = title; hdr(c)
    for i, v in enumerate(values, start=2):
        cell = ws[f"{col}{i}"]; cell.value = v
        cell.font = Font(name=FONT_BODY, size=SZ_BODY, color=C_DARKSLATEGRAY)
        if i % 2 == 0: cell.fill = PatternFill("solid", fgColor=C_CREAM)


# Coded tables
write_coded_table(1, "Catalog Item Type", cat_item_types)         # A:B
write_coded_table(3, "Digital Item Type", dig_item_types, True)   # C:D
write_coded_table(5, "Archive Item Type", arc_item_types)         # E:F
write_coded_table(7, "Subject", subjects, True)                   # G:H
write_coded_table(9, "Genre / Form", genres)                      # I:J

# Plain vocabularies (K onward)
write_plain_table(11, "Language", languages)                     # K
write_plain_table(12, "Format", formats)                         # L
write_plain_table(13, "OCR Status", ocr_status)                  # M
write_plain_table(14, "Condition", conditions)                   # N
write_plain_table(15, "Location Type", location_types)           # O
write_plain_table(16, "Access Level", access_levels)             # P
write_plain_table(17, "Rights", rights_vals)                     # Q
write_plain_table(18, "Acquisition Type", acq_types)             # R
write_plain_table(19, "Physical/Digital", phys_digi)             # S
write_plain_table(20, "Membership Type", membership_types)       # T
write_plain_table(21, "Member Status", member_status)             # U
write_plain_table(22, "Reading/Usage Status", reading_status)    # V
write_plain_table(23, "Preservation Status", preservation_status) # W
write_plain_table(24, "Collection", collection_vals)             # X
write_plain_table(25, "Media Type", media_types)                 # Y
write_plain_table(26, "Circulation Status", circ_status)         # Z
write_plain_table(27, "Frequency", frequency_vals)               # AA
write_plain_table(28, "Digitisation Status", digitisation_status) # AB


# ═══════════════════════════════════════════════════════════════════════
# 3. NAMED RANGES
# ═══════════════════════════════════════════════════════════════════════
n_cat = len(cat_item_types)
n_dig = len(dig_item_types)
n_arc = len(arc_item_types)
n_sub = len(subjects)
n_gen = len(genres)
n_lang = len(languages)
n_fmt = len(formats)
n_ocr = len(ocr_status)
n_cond = len(conditions)
n_loc = len(location_types)
n_acc = len(access_levels)
n_rgt = len(rights_vals)
n_acq = len(acq_types)
n_pd = len(phys_digi)
n_mem = len(membership_types)
n_mstat = len(member_status)
n_read = len(reading_status)
n_pres = len(preservation_status)
n_coll = len(collection_vals)
n_media = len(media_types)
n_circ = len(circ_status)
n_freq = len(frequency_vals)
n_digi = len(digitisation_status)

named_ranges = [
    # Item Type — three separate lookup tables
    ("CatItemTypeList",   f"Lookup_Codes!$A$2:$A${n_cat+1}"),
    ("CatItemTypeLookup", f"Lookup_Codes!$A$2:$B${n_cat+1}"),
    ("DigItemTypeList",   f"Lookup_Codes!$C$2:$C${n_dig+1}"),
    ("DigItemTypeLookup", f"Lookup_Codes!$C$2:$D${n_dig+1}"),
    ("ArcItemTypeList",   f"Lookup_Codes!$E$2:$E${n_arc+1}"),
    ("ArcItemTypeLookup", f"Lookup_Codes!$E$2:$F${n_arc+1}"),
    # Subject & Genre
    ("SubjectList",       f"Lookup_Codes!$G$2:$G${n_sub+1}"),
    ("SubjectLookup",     f"Lookup_Codes!$G$2:$H${n_sub+1}"),
    ("GenreList",         f"Lookup_Codes!$I$2:$I${n_gen+1}"),
    ("GenreLookup",       f"Lookup_Codes!$I$2:$J${n_gen+1}"),
    # Plain vocabularies
    ("LanguageList",      f"Lookup_Codes!$K$2:$K${n_lang+1}"),
    ("FormatList",        f"Lookup_Codes!$L$2:$L${n_fmt+1}"),
    ("OCRStatusList",     f"Lookup_Codes!$M$2:$M${n_ocr+1}"),
    ("ConditionList",     f"Lookup_Codes!$N$2:$N${n_cond+1}"),
    ("LocationTypeList",  f"Lookup_Codes!$O$2:$O${n_loc+1}"),
    ("AccessLevelList",   f"Lookup_Codes!$P$2:$P${n_acc+1}"),
    ("RightsList",        f"Lookup_Codes!$Q$2:$Q${n_rgt+1}"),
    ("AcqTypeList",       f"Lookup_Codes!$R$2:$R${n_acq+1}"),
    ("PhysDigiList",      f"Lookup_Codes!$S$2:$S${n_pd+1}"),
    ("MembershipList",    f"Lookup_Codes!$T$2:$T${n_mem+1}"),
    ("MemberStatusList",  f"Lookup_Codes!$U$2:$U${n_mstat+1}"),
    ("ReadingStatusList", f"Lookup_Codes!$V$2:$V${n_read+1}"),
    ("PreservationList",  f"Lookup_Codes!$W$2:$W${n_pres+1}"),
    ("CollectionList",    f"Lookup_Codes!$X$2:$X${n_coll+1}"),
    ("MediaTypeList",     f"Lookup_Codes!$Y$2:$Y${n_media+1}"),
    ("CirculationList",   f"Lookup_Codes!$Z$2:$Z${n_circ+1}"),
    ("FrequencyList",     f"Lookup_Codes!$AA$2:$AA${n_freq+1}"),
    ("DigitisationList",  f"Lookup_Codes!$AB$2:$AB${n_digi+1}"),
]
for name, ref in named_ranges:
    wb.defined_names.add(DefinedName(name, attr_text=ref))


# ═══════════════════════════════════════════════════════════════════════
# 4. CATALOG — books and print only
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Catalog")
cat_cols = [
    ("Item ID", 8), ("Accession Number", 16), ("Call Number", 26),
    ("Title", 32), ("Subtitle", 22), ("Creator", 26),
    ("Author Code Override", 12), ("Author Code", 12),
    ("Editor", 20), ("Translator", 20),
    ("Item Type", 16), ("Subject", 18), ("Genre / Form", 18), ("Language", 14),
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
    ws.column_dimensions[get_column_letter(i)].width = w
    hdr(ws.cell(1, i, name))
ws.row_dimensions[1].height = 34
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(cat_cols))}{ROWS_CAT + 1}"

for r in range(2, ROWS_CAT + 2):
    # A — Item ID
    ws.cell(r, 1, f'=IF(ISBLANK(D{r}),"",ROW()-1)'); gen_cell(ws.cell(r, 1))
    # B — Accession Number
    input_mono(ws.cell(r, 2))
    # C — Call Number  (CatItemTypeLookup + SubjectLookup + GenreLookup)
    ws.cell(r, 3, (
        f'=IF(OR(ISBLANK(D{r}),ISBLANK(K{r})),"",'
        f'IFERROR(VLOOKUP(K{r},CatItemTypeLookup,2,FALSE),"?")&"-"&'
        f'IF(L{r}="","?",IFERROR(VLOOKUP(L{r},SubjectLookup,2,FALSE),"="))&"-"&'
        f'IF(M{r}="","?",IFERROR(VLOOKUP(M{r},GenreLookup,2,FALSE),"="))&"-"&'
        f'IF(H{r}="","NOC",H{r})&"-"&'
        f'TEXT(A{r},"000")&'
        f'IF(ISBLANK(O{r}),"","-"&SUBSTITUTE(O{r}," ","")))'
    ))
    gen_cell(ws.cell(r, 3))
    # D–F
    for col in (4, 5, 6): input_cell(ws.cell(r, col))
    # G — Author Code Override
    input_cell(ws.cell(r, 7))
    # H — Author Code
    ws.cell(r, 8, (
        f'=IF(ISBLANK(F{r}),"",'
        f'IF(G{r}<>"",UPPER(TRIM(G{r})),'
        f'IF(ISNUMBER(SEARCH(" ",TRIM(F{r}))),'
        f'UPPER(LEFT(TRIM(RIGHT(SUBSTITUTE(TRIM(F{r})," ",REPT(" ",100)),100)),Settings!$B$9)),'
        f'UPPER(LEFT(TRIM(F{r}),Settings!$B$9)))))'
    ))
    gen_cell(ws.cell(r, 8))
    # I–N
    for col in (9, 10, 11, 12, 13, 14):
        input_cell(ws.cell(r, col))
    # O–S
    for col in (15, 16, 17, 18, 19):
        input_cell(ws.cell(r, col))
    # T–W
    for col in (20, 21, 22, 23):
        input_cell(ws.cell(r, col))
    # X–Z
    for col in (24, 25, 26):
        input_cell(ws.cell(r, col))
    # AA — Circulation Status
    ws.cell(r, 27, (
        f'=IF(ISBLANK(D{r}),"",'
        f'IF(COUNTIFS(Circulation!$B$2:$B${ROWS_CIR},A{r},'
        f'Circulation!$H$2:$H${ROWS_CIR},"Active")>0,"On Loan","Available"))'
    ))
    gen_body(ws.cell(r, 27))
    # AB–AH
    for col in range(28, 35):
        input_cell(ws.cell(r, col))

for fml, rng in [
    ("=CatItemTypeList",  f"K2:K{ROWS_CAT+1}"),
    ("=SubjectList",      f"L2:L{ROWS_CAT+1}"),
    ("=GenreList",        f"M2:M{ROWS_CAT+1}"),
    ("=LanguageList",     f"N2:N{ROWS_CAT+1}"),
    ("=LocationTypeList", f"X2:X{ROWS_CAT+1}"),
    ("=ConditionList",    f"Z2:Z{ROWS_CAT+1}"),
    ("=AcqTypeList",      f"AC2:AC{ROWS_CAT+1}"),
    ("=RightsList",       f"AG2:AG{ROWS_CAT+1}"),
]:
    dv = DataValidation(type="list", formula1=fml, allow_blank=True)
    ws.add_data_validation(dv); dv.add(rng)

ws.conditional_formatting.add(
    f"D2:D{ROWS_CAT+1}",
    FormulaRule(formula=['ISBLANK($D2)'],
                fill=PatternFill("solid", fgColor=C_ERROR_SOFT)))
ws.conditional_formatting.add(
    f"F2:F{ROWS_CAT+1}",
    FormulaRule(formula=['AND($D2<>"",ISBLANK($F2))'],
                fill=PatternFill("solid", fgColor=C_WARN_SOFT)))
ws.conditional_formatting.add(
    f"AA2:AA{ROWS_CAT+1}",
    FormulaRule(formula=['$AA2="On Loan"'],
                fill=PatternFill("solid", fgColor=C_WARN_SOFT)))

if not BLANK_MODE:
    rows = []
    for vol in range(1, 19):
        shelf = "A2" if vol <= 8 else "A1"
        rows.append((
            f"ACC-{vol:04d}", "Rabindra Rachanabali", f"Vol {vol}",
            "Rabindranath Tagore", "Book", "Literature", "Collected Works",
            "Bengali", f"Vol {vol}", "Shelf", shelf, "Good",
        ))
    rows.append((
        "ACC-0019", "নীলকণ্ঠ পাখির খোঁজে", "Vol 1",
        "অতীন বন্দ্যোপাধ্যায়", "Book", "Literature", "Novel",
        "Bengali", "Vol 1", "Shelf", "B1", "Good",
    ))
    rows.append((
        "ACC-0020", "Desh", "", "", "Magazine", "Literature", "Other",
        "Bengali", "", "Magazine Rack", "MR1", "Good",
    ))
    for i, row in enumerate(rows):
        r = i + 2
        (acc, title, sub, creator, itype, subject, genre,
         lang, vol, loc_type, loc_id, cond) = row
        ws.cell(r, 2, acc)
        ws.cell(r, 4, title)
        ws.cell(r, 5, sub)
        ws.cell(r, 6, creator)
        ws.cell(r, 11, itype)
        ws.cell(r, 12, subject)
        ws.cell(r, 13, genre)
        ws.cell(r, 14, lang)
        ws.cell(r, 15, vol)
        ws.cell(r, 24, loc_type)
        ws.cell(r, 25, loc_id)
        ws.cell(r, 26, cond)


# ═══════════════════════════════════════════════════════════════════════
# 5. DIGITAL_CATALOG — digital resources only
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Digital_Catalog")
dig_cols = [
    ("Resource ID", 8), ("Digital Call Number", 30),
    ("Title", 32), ("Creator", 26), ("Contributor", 22),
    ("Item Type", 18), ("Subject", 18), ("Genre / Form", 18), ("Language", 14),
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
    ws.cell(r, 1, f'=IF(ISBLANK(C{r}),"",ROW()-1)'); gen_cell(ws.cell(r, 1))
    ws.cell(r, 2, (
        f'=IF(OR(ISBLANK(C{r}),ISBLANK(F{r})),"","DIG-"&'
        f'IFERROR(VLOOKUP(F{r},DigItemTypeLookup,2,FALSE),"?")&"-"&'
        f'IF(G{r}="","?",IFERROR(VLOOKUP(G{r},SubjectLookup,2,FALSE),"="))&"-"&'
        f'IF(H{r}="","?",IFERROR(VLOOKUP(H{r},GenreLookup,2,FALSE),"="))&"-"&'
        f'IF(ISBLANK(D{r}),"NOC",'
        f'UPPER(LEFT(TRIM(RIGHT(SUBSTITUTE(TRIM(D{r})," ",REPT(" ",100)),100)),3)))&"-"&'
        f'TEXT(A{r},"000"))'
    ))
    gen_cell(ws.cell(r, 2))
    for col in range(3, len(dig_cols) + 1):
        input_cell(ws.cell(r, col))

for fml, rng in [
    ("=DigItemTypeList",  f"F2:F{ROWS_DIG+1}"),
    ("=SubjectList",      f"G2:G{ROWS_DIG+1}"),
    ("=GenreList",        f"H2:H{ROWS_DIG+1}"),
    ("=LanguageList",     f"I2:I{ROWS_DIG+1}"),
    ("=FormatList",       f"J2:J{ROWS_DIG+1}"),
    ("=OCRStatusList",    f"P2:P{ROWS_DIG+1}"),
    ("=RightsList",       f"T2:T{ROWS_DIG+1}"),
    ("=AccessLevelList",  f"U2:U{ROWS_DIG+1}"),
    ("=ReadingStatusList",f"AA2:AA{ROWS_DIG+1}"),
    ("=PreservationList", f"AB2:AB{ROWS_DIG+1}"),
]:
    dv = DataValidation(type="list", formula1=fml, allow_blank=True)
    ws.add_data_validation(dv); dv.add(rng)

if not BLANK_MODE:
    dig_samples = [
        ("The Pocket Oracle", "Baltasar Gracian", "", "E-Book", "Philosophy", "Reference",
         "English", "PDF", "application/pdf", ".pdf", "", "Drive/Books/Philosophy", "",
         "Searchable (OCR)", "Yes", "2024-01-15", "2024-01-15", "Copyright", "Private",
         "A pocket edition of Gracian's aphorisms.", "", "", "", "", "Unread", "Master"),
        ("নীলকণ্ঠ পাখির খোঁজে", "অতীন বন্দ্যোপাধ্যায়", "", "E-Book", "Literature", "Novel",
         "Bengali", "PDF", "application/pdf", ".pdf", "", "Drive/Books/Bengali_Classics", "",
         "Scanned (No OCR)", "No", "2023-06-01", "2023-06-01", "Copyright", "Private",
         "Scanned copy of Bengali novel.", "", "", "", "", "Unread", "Master"),
        ("The Great Derangement", "Amitav Ghosh", "", "E-Book", "Literature", "Essay",
         "English", "EPUB", "application/epub+zip", ".epub", "", "Drive/Books/Essays", "",
         "Searchable (OCR)", "Yes", "2023-02-10", "2023-02-10", "Copyright", "Private",
         "Climate change essays.", "", "", "", "", "Read", "Master"),
        ("Bengali literary reading", "", "", "Audio Recording", "Literature", "Poem",
         "Bengali", "FLAC", "audio/flac", ".flac", "", "Drive/Archive/Audio", "",
         "Not Applicable", "No", "2023-08-12", "2023-08-12", "User-Owned", "Private",
         "Live reading of Tagore poems.", "", "", "", "", "N/A", "Master"),
        ("Interview with an author", "", "", "Video Recording", "Literature", "Interview",
         "Bengali", "MP4", "video/mp4", ".mp4", "", "Drive/Archive/Video", "",
         "Not Applicable", "No", "2024-03-01", "2024-03-01", "User-Owned", "Restricted",
         "Recorded interview.", "", "", "", "", "N/A", "Master"),
        ("Scanned Bengali manuscript", "", "", "Manuscript", "Literature", "Other",
         "Bengali", "TIFF", "image/tiff", ".tiff", "", "Archive/Manuscripts", "",
         "Scanned (No OCR)", "No", "2022-01-01", "2022-01-01", "Public Domain",
         "Research Only", "High-res scan of 19th-c manuscript.", "", "", "ARC-0001", "",
         "N/A", "Master"),
    ]
    for i, row in enumerate(dig_samples):
        r = i + 2
        for j, val in enumerate(row, start=3):
            ws.cell(r, j, val)


# ═══════════════════════════════════════════════════════════════════════
# 6. ARCHIVE — archival objects (physical, digital, or both)
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Archive")
arc_cols = [
    ("Archive ID", 12), ("Title", 32), ("Creator", 24), ("Date", 14),
    ("Description", 34), ("Resource Type", 22), ("Physical/Digital", 14),
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
    ("=ArcItemTypeList",  f"F2:F{ROWS_ARC+1}"),
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
    ws.add_data_validation(dv); dv.add(rng)

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
# 7. CIRCULATION
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
    ws.cell(r, 1, f'=IF(ISBLANK(B{r}),"",ROW()-1)'); gen_cell(ws.cell(r, 1))
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
# 8. MEMBERS
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
    ws.add_data_validation(dv); dv.add(rng)

if not BLANK_MODE:
    ws.cell(2, 1, "M001")
    ws.cell(2, 2, "Sample Borrower")
    ws.cell(2, 4, "Public")
    ws.cell(2, 5, "2024-01-01")
    ws.cell(2, 6, "Active")


# ═══════════════════════════════════════════════════════════════════════
# 9. SPINE_LABELS
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
    ws.cell(r, 4, f'=Catalog!O{r}'); gen_body(ws.cell(r, 4))
    ws.cell(r, 5, f'=Catalog!F{r}'); gen_body(ws.cell(r, 5))
    ws.cell(r, 6, f'=IF(Catalog!X{r}="","",Catalog!X{r}&" "&Catalog!Y{r})')
    gen_body(ws.cell(r, 6))
    input_cell(ws.cell(r, 7))

dv = DataValidation(type="list",
                    formula1='"Standard Card,Library Block,Minimal Spine"',
                    allow_blank=True)
ws.add_data_validation(dv); dv.add(f"G2:G{ROWS_SPN+1}")


# ═══════════════════════════════════════════════════════════════════════
# 10. DUBLIN_CORE
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Dublin_Core")
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 36
ws.column_dimensions["C"].width = 48

title_cell(ws["A1"])
ws["A1"] = "DUBLIN CORE METADATA MAPPING"
ws.merge_cells("A1:C1")
ws.row_dimensions[1].height = 32

for i, h in enumerate(["Dublin Core Element", "Catalog Field", "Notes"], start=1):
    hdr(ws.cell(3, i, h))

mapping = [
    ("title",       "Title",                          "Standard DC element"),
    ("creator",     "Creator",                        "Standard DC element"),
    ("contributor", "Editor, Translator",             "Standard DC element; split"),
    ("publisher",   "Publisher",                      "Standard DC element"),
    ("date",        "Publication Year",               "Standard DC element"),
    ("language",    "Language",                       "Standard DC element"),
    ("type",        "Item Type / Genre / Form",       "Standard DC element; two columns"),
    ("format",      "Print medium",                   "Standard DC element"),
    ("identifier",  "ISBN, ISSN, Accession, Call No", "Standard DC element"),
    ("subject",     "Subject",                        "Standard DC element"),
    ("description", "Notes, Description",             "Standard DC element"),
    ("source",      "Donor/Vendor, Acquisition Type", "Standard DC element"),
    ("relation",    "Volume, Series, Related ID",     "isPartOf qualifier"),
    ("coverage",    "Publication Place",              "spatial qualifier"),
    ("rights",      "Rights",                         "Standard DC element"),
    ("—",           "Item ID, Call Number, Author Code",  "Application-specific"),
    ("—",           "Location, Circulation Status, Condition", "Application-specific"),
    ("—",           "Dimensions, Material, Collection", "Application-specific"),
    ("—",           "Storage path, OCR status, Public Link, Checksum", "Application-specific"),
]
for i, row in enumerate(mapping, start=4):
    for j, val in enumerate(row, start=1):
        c = ws.cell(i, j, val); body_cell(c); c.border = BORDER_THIN
        if j == 1:
            c.font = Font(name=FONT_MONO, size=SZ_BODY, color=C_DARKSLATEBLUE, bold=True)

h1_cell(ws["A26"]); ws["A26"] = "EXPORT VIEW — Dublin Core element columns"

dc_view = [
    "dc.title", "dc.creator", "dc.contributor", "dc.publisher", "dc.date",
    "dc.language", "dc.type", "dc.format", "dc.identifier", "dc.subject",
    "dc.description", "dc.source", "dc.relation", "dc.coverage", "dc.rights",
]
for j, name in enumerate(dc_view, start=1):
    c = ws.cell(28, j, name)
    c.font = Font(name=FONT_MONO, bold=True, color="FFFFFF", size=SZ_SMALL)
    c.fill = PatternFill("solid", fgColor=C_DARKSLATEBLUE)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER_THIN

for r in range(2, ROWS_CAT + 2):
    row = r + 27
    formulas = [
        f'=Catalog!D{r}',                                    # title
        f'=Catalog!F{r}',                                    # creator
        f'=IF(Catalog!I{r}="","",Catalog!I{r})&IF(AND(Catalog!I{r}<>"",Catalog!J{r}<>""),"; ","")&IF(Catalog!J{r}="","",Catalog!J{r})',
        f'=Catalog!Q{r}',                                    # publisher
        f'=Catalog!S{r}',                                    # date
        f'=Catalog!N{r}',                                    # language
        f'=Catalog!K{r}&" / "&Catalog!M{r}',                 # type
        f'="Print"',                                         # format
        f'=IF(Catalog!T{r}<>"","ISBN:"&Catalog!T{r},IF(Catalog!U{r}<>"","ISSN:"&Catalog!U{r},"ACC:"&Catalog!B{r}))',
        f'=Catalog!L{r}',                                    # subject
        f'=IF(Catalog!AH{r}<>"",Catalog!AH{r},Catalog!AF{r})',  # description
        f'=Catalog!AE{r}',                                   # source
        f'=Catalog!W{r}',                                    # relation
        f'=Catalog!R{r}',                                    # coverage
        f'=Catalog!AG{r}',                                   # rights
    ]
    for j, fml in enumerate(formulas, start=1):
        gen_body(ws.cell(row, j, fml))


# ═══════════════════════════════════════════════════════════════════════
# 11. DATA_QUALITY
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Data_Quality")
ws.column_dimensions["A"].width = 38
ws.column_dimensions["B"].width = 14

title_cell(ws["A1"]); ws["A1"] = "DATA QUALITY REPORT"
ws.merge_cells("A1:B1"); ws.row_dimensions[1].height = 32

ws["A3"] = "Metric"; ws["B3"] = "Value"
hdr(ws["A3"]); hdr_brown(ws["B3"])

checks = [
    ("Total Catalog records", f'=COUNTA(Catalog!$D$2:$D${ROWS_CAT+1})'),
    ("Records missing Title", f'=COUNTA(Catalog!$A$2:$A${ROWS_CAT+1})-COUNTA(Catalog!$D$2:$D${ROWS_CAT+1})'),
    ("Records missing Creator", f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$F$2:$F${ROWS_CAT+1}=""))'),
    ("Records missing Item Type", f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$K$2:$K${ROWS_CAT+1}=""))'),
    ("Records missing Subject", f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$L$2:$L${ROWS_CAT+1}=""))'),
    ("Records missing Genre/Form", f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$M$2:$M${ROWS_CAT+1}=""))'),
    ("Records missing Location Type", f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$X$2:$X${ROWS_CAT+1}=""))'),
    ("Records missing Location Id", f'=SUMPRODUCT((Catalog!$D$2:$D${ROWS_CAT+1}<>"")*(Catalog!$Y$2:$Y${ROWS_CAT+1}=""))'),
    ("Duplicate ISBNs", f'=SUMPRODUCT((Catalog!$T$2:$T${ROWS_CAT+1}<>"")*(COUNTIF(Catalog!$T$2:$T${ROWS_CAT+1},Catalog!$T$2:$T${ROWS_CAT+1})>1))/2'),
    ("Duplicate ISSNs", f'=SUMPRODUCT((Catalog!$U$2:$U${ROWS_CAT+1}<>"")*(COUNTIF(Catalog!$U$2:$U${ROWS_CAT+1},Catalog!$U$2:$U${ROWS_CAT+1})>1))/2'),
    ("Duplicate Accessions", f'=SUMPRODUCT((Catalog!$B$2:$B${ROWS_CAT+1}<>"")*(COUNTIF(Catalog!$B$2:$B${ROWS_CAT+1},Catalog!$B$2:$B${ROWS_CAT+1})>1))/2'),
    ("Digital resources", f'=COUNTA(Digital_Catalog!$C$2:$C${ROWS_DIG+1})'),
    ("Archive objects", f'=COUNTA(Archive!$B$2:$B${ROWS_ARC+1})'),
    ("Archive missing Collection", f'=SUMPRODUCT((Archive!$B$2:$B${ROWS_ARC+1}<>"")*(Archive!$U$2:$U${ROWS_ARC+1}=""))'),
    ("Archive missing Dimensions", f'=SUMPRODUCT((Archive!$B$2:$B${ROWS_ARC+1}<>"")*(Archive!$V$2:$V${ROWS_ARC+1}=""))'),
    ("Active loans", f'=COUNTIF(Circulation!$H$2:$H${ROWS_CIR+1},"Active")'),
    ("Overdue loans", f'=COUNTIF(Circulation!$H$2:$H${ROWS_CIR+1},"Overdue")'),
]
for i, (label, f) in enumerate(checks, start=4):
    lbl = ws.cell(i, 1, label)
    lbl.font = Font(name=FONT_BODY, bold=True, size=SZ_BODY, color=C_DARKSLATEGRAY)
    lbl.border = BORDER_THIN
    if i % 2 == 0:
        lbl.fill = PatternFill("solid", fgColor=C_LAVENDER)
    val = ws.cell(i, 2, f)
    val.font = Font(name=FONT_MONO, bold=True, size=SZ_BODY, color=C_DARKSLATEBLUE)
    val.alignment = Alignment(horizontal="right", vertical="center")
    val.border = BORDER_THIN
    val.fill = PatternFill("solid", fgColor=C_AQUAMARINE)

h1_cell(ws["A24"]); ws["A24"] = "RECORD-LEVEL CHECKS"
for j, name in enumerate(["Item ID", "Title", "Issues"], start=1):
    hdr(ws.cell(25, j, name))

for r in range(2, ROWS_CAT + 2):
    row = r + 24
    ws.cell(row, 1, f'=Catalog!A{r}').font = Font(name=FONT_MONO, size=SZ_SMALL, color=C_DARKSLATEGRAY)
    ws.cell(row, 2, f'=Catalog!D{r}').font = Font(name=FONT_BODY, size=SZ_SMALL, color=C_DARKSLATEGRAY)
    c3 = ws.cell(row, 3, (
        f'=IF(Catalog!D{r}="","",'
        f'IF(ISBLANK(Catalog!F{r}),"Missing Creator; ","")&'
        f'IF(ISBLANK(Catalog!K{r}),"Missing Item Type; ","")&'
        f'IF(ISBLANK(Catalog!L{r}),"Missing Subject; ","")&'
        f'IF(ISBLANK(Catalog!M{r}),"Missing Genre/Form; ","")&'
        f'IF(ISBLANK(Catalog!X{r}),"Missing Location Type; ","")&'
        f'IF(ISBLANK(Catalog!Y{r}),"Missing Location Id; ","")&'
        f'IF(AND(Catalog!T{r}<>"",COUNTIF(Catalog!$T$2:$T${ROWS_CAT+1},Catalog!T{r})>1),"Duplicate ISBN; ","")&'
        f'IF(AND(Catalog!B{r}<>"",COUNTIF(Catalog!$B$2:$B${ROWS_CAT+1},Catalog!B{r})>1),"Duplicate Accession; ",""))'
    ))
    c3.font = Font(name=FONT_SMALL, size=SZ_SMALL, color=C_BROWN)
ws.freeze_panes = "A26"


# ═══════════════════════════════════════════════════════════════════════
# 12. DASHBOARD
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Dashboard")
ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 36
ws.column_dimensions["C"].width = 16
ws.column_dimensions["D"].width = 4
ws.column_dimensions["E"].width = 36
ws.column_dimensions["F"].width = 16

title_cell(ws["B2"]); ws["B2"] = "PERSONAL LIBRARY MANAGER & DIGITAL ARCHIVE"
ws.merge_cells("B2:F2"); ws.row_dimensions[2].height = 36
ws["B3"] = "Dashboard"
ws["B3"].font = Font(name=FONT_SECTION, bold=True, size=SZ_H2, color=C_BROWN)
ws.merge_cells("B3:F3"); ws.row_dimensions[3].height = 24


def dash_block(anchor_row, title, items, col_letter, brown=False):
    col_offset = ord(col_letter) - 64
    c = ws.cell(anchor_row, col_offset, title)
    (section_banner_brown if brown else section_banner)(c)
    end_col = get_column_letter(col_offset + 1)
    ws.merge_cells(f"{col_letter}{anchor_row}:{end_col}{anchor_row}")
    ws.row_dimensions[anchor_row].height = 26
    for i, (label, f) in enumerate(items, start=1):
        r = anchor_row + i
        lbl = ws.cell(r, col_offset, label)
        metric_label(lbl); lbl.border = BORDER_THIN
        val = ws.cell(r, col_offset + 1, f)
        metric_value(val); val.border = BORDER_THIN


dash_block(5, "COLLECTION", [
    ("Physical items",    f'=COUNTA(Catalog!A2:A{ROWS_CAT+1})'),
    ("Digital resources", f'=COUNTA(Digital_Catalog!A2:A{ROWS_DIG+1})'),
    ("Archive objects",   f'=COUNTA(Archive!A2:A{ROWS_ARC+1})'),
    ("Books",             f'=COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Book")'),
    ("Magazines",         f'=COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Magazine")'),
    ("Journals",          f'=COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Journal")'),
    ("Newspapers",        f'=COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Newspaper")'),
    ("Manuscripts",       f'=COUNTIF(Catalog!K2:K{ROWS_CAT+1},"Manuscript")'),
    ("Other print",
     f'=COUNTA(Catalog!A2:A{ROWS_CAT+1})'
     f'-SUM(COUNTIF(Catalog!K2:K{ROWS_CAT+1},'
     f'{{"Book","Magazine","Journal","Newspaper","Manuscript"}}))'),
], "B")

dash_block(5, "METADATA", [
    ("Missing titles",
     f'=COUNTA(Catalog!A2:A{ROWS_CAT+1})-COUNTA(Catalog!D2:D{ROWS_CAT+1})'),
    ("Missing creators",
     f'=SUMPRODUCT((Catalog!D2:D{ROWS_CAT+1}<>"")*(Catalog!F2:F{ROWS_CAT+1}=""))'),
    ("Missing Item Type",
     f'=SUMPRODUCT((Catalog!D2:D{ROWS_CAT+1}<>"")*(Catalog!K2:K{ROWS_CAT+1}=""))'),
    ("Missing Subject",
     f'=SUMPRODUCT((Catalog!D2:D{ROWS_CAT+1}<>"")*(Catalog!L2:L{ROWS_CAT+1}=""))'),
    ("Missing Genre/Form",
     f'=SUMPRODUCT((Catalog!D2:D{ROWS_CAT+1}<>"")*(Catalog!M2:M{ROWS_CAT+1}=""))'),
    ("Duplicate ISBNs",
     f'=SUMPRODUCT((Catalog!T2:T{ROWS_CAT+1}<>"")*(COUNTIF(Catalog!T2:T{ROWS_CAT+1},Catalog!T2:T{ROWS_CAT+1})>1))/2'),
], "E", brown=True)

dash_block(15, "READING STATUS", [
    ("Read",    f'=COUNTIF(Digital_Catalog!AA2:AA{ROWS_DIG+1},"Read")'),
    ("Reading", f'=COUNTIF(Digital_Catalog!AA2:AA{ROWS_DIG+1},"Reading")'),
    ("Unread",  f'=COUNTIF(Digital_Catalog!AA2:AA{ROWS_DIG+1},"Unread")'),
], "B")

dash_block(15, "CIRCULATION", [
    ("Available", f'=COUNTIF(Catalog!AA2:AA{ROWS_CAT+1},"Available")'),
    ("On Loan",   f'=COUNTIF(Catalog!AA2:AA{ROWS_CAT+1},"On Loan")'),
    ("Overdue",   f'=COUNTIF(Circulation!H2:H{ROWS_CIR+1},"Overdue")'),
], "E", brown=True)

dash_block(21, "ARCHIVE BY COLLECTION", [
    ("Manuscripts", f'=COUNTIF(Archive!U2:U{ROWS_ARC+1},"Manuscripts")'),
    ("Photographs", f'=COUNTIF(Archive!U2:U{ROWS_ARC+1},"Photographs")'),
    ("Audio",       f'=COUNTIF(Archive!U2:U{ROWS_ARC+1},"Audio")'),
    ("Video",       f'=COUNTIF(Archive!U2:U{ROWS_ARC+1},"Video")'),
    ("Sculptures",  f'=COUNTIF(Archive!U2:U{ROWS_ARC+1},"Sculptures")'),
    ("Textiles",    f'=COUNTIF(Archive!U2:U{ROWS_ARC+1},"Textiles")'),
], "B")

dash_block(21, "DIGITISATION", [
    ("Physical only", f'=COUNTIF(Archive!G2:G{ROWS_ARC+1},"Physical")'),
    ("Digital only",  f'=COUNTIF(Archive!G2:G{ROWS_ARC+1},"Digital")'),
    ("Both",          f'=COUNTIF(Archive!G2:G{ROWS_ARC+1},"Both")'),
    ("OCR Pending",   f'=COUNTIF(Archive!P2:P{ROWS_ARC+1},"OCR Pending")'),
    ("OCR Complete",  f'=COUNTIF(Archive!P2:P{ROWS_ARC+1},"OCR Complete")'),
], "E", brown=True)


# ═══════════════════════════════════════════════════════════════════════
# 13. FORMULAS
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("Formulas")
ws.column_dimensions["A"].width = 18
ws.column_dimensions["B"].width = 24
ws.column_dimensions["C"].width = 80
ws.column_dimensions["D"].width = 44

title_cell(ws["A1"])
ws["A1"] = "FORMULA REFERENCE — Every generated cell in this workbook"
ws.merge_cells("A1:D1")
ws.row_dimensions[1].height = 32

ws["A3"] = ("Row 4 and below list each generated cell, its purpose, and the exact formula pattern "
            "(row 2 as example). Only cells with pale-aqua background contain formulas.")
small_cell(ws["A3"]); ws.merge_cells("A3:D3")

for j, h in enumerate(["Sheet", "Column / Cell", "Formula", "Purpose"], start=1):
    hdr(ws.cell(5, j, h))

formulas_doc = [
    # Catalog
    ("Catalog", "A — Item ID",
     '=IF(ISBLANK(D2),"",ROW()-1)',
     "Sequential ID; blank if no Title."),
    ("Catalog", "C — Call Number",
     '=IF(OR(ISBLANK(D2),ISBLANK(K2)),"",'
     'IFERROR(VLOOKUP(K2,CatItemTypeLookup,2,FALSE),"?")&"-"&'
     'IF(L2="","?",IFERROR(VLOOKUP(L2,SubjectLookup,2,FALSE),"="))&"-"&'
     'IF(M2="","?",IFERROR(VLOOKUP(M2,GenreLookup,2,FALSE),"="))&"-"&'
     'IF(H2="","NOC",H2)&"-"&TEXT(A2,"000")&'
     'IF(ISBLANK(O2),"","-"&SUBSTITUTE(O2," ","")))',
     "Compose call number: Item Type code + Subject code + Genre code + Author Code + ID [+ Volume]."),
    ("Catalog", "H — Author Code",
     '=IF(ISBLANK(F2),"",'
     'IF(G2<>"",UPPER(TRIM(G2)),'
     'IF(ISNUMBER(SEARCH(" ",TRIM(F2))),'
     'UPPER(LEFT(TRIM(RIGHT(SUBSTITUTE(TRIM(F2)," ",REPT(" ",100)),100)),Settings!$B$9)),'
     'UPPER(LEFT(TRIM(F2),Settings!$B$9)))))',
     "Auto 3-letter code from surname; manual override takes precedence."),
    ("Catalog", "AA — Circulation Status",
     '=IF(ISBLANK(D2),"",'
     'IF(COUNTIFS(Circulation!$B$2:$B$501,A2,Circulation!$H$2:$H$501,"Active")>0,'
     '"On Loan","Available"))',
     "'On Loan' if an active loan exists for this item."),

    # Digital_Catalog
    ("Digital_Catalog", "A — Resource ID",
     '=IF(ISBLANK(C2),"",ROW()-1)',
     "Sequential ID; blank if no Title."),
    ("Digital_Catalog", "B — Digital Call Number",
     '=IF(OR(ISBLANK(C2),ISBLANK(F2)),"",'
     '"DIG-" &'
     'IFERROR(VLOOKUP(F2,DigItemTypeLookup,2,FALSE),"?") & "-" &'
     'IF(G2="","?",IFERROR(VLOOKUP(G2,SubjectLookup,2,FALSE),"=")) & "-" &'
     'IF(H2="","?",IFERROR(VLOOKUP(H2,GenreLookup,2,FALSE),"=")) & "-" &'
     'IF(ISBLANK(D2),"NOC",'
     'UPPER(LEFT(TRIM(RIGHT(SUBSTITUTE(TRIM(D2)," ",REPT(" ",100)),100)),3))) & "-" &'
     'TEXT(A2,"000"))',
     "Same code pattern with DIG- prefix."),

    # Archive
    ("Archive", "A — Archive ID",
     '=IF(ISBLANK(B2),"","ARC-"&TEXT(ROW()-1,"0000"))',
     "Auto zero-padded archive ID."),
    ("Archive", "AB — Media Preview",
     '=IF(AA2="","",HYPERLINK(AA2,"Open ↗"))',
     "Turn Public Link (AA) into a clickable hyperlink."),

    # Circulation
    ("Circulation", "A — Loan ID",
     '=IF(ISBLANK(B2),"",ROW()-1)',
     "Sequential loan number."),
    ("Circulation", "D — Borrower Name",
     '=IFERROR(IF(ISBLANK(C2),"",VLOOKUP(C2,Members!$A$2:$B$101,2,FALSE)),"")',
     "Look up member name from Member ID."),
    ("Circulation", "F — Due Date",
     '=IF(ISBLANK(E2),"",E2+Settings!$B$7)',
     "Checkout date + loan period."),
    ("Circulation", "H — Status",
     '=IF(ISBLANK(B2),"",'
     'IF(NOT(ISBLANK(G2)),"Returned",'
     'IF(TODAY()>F2,"Overdue","Active")))',
     "Returned / Overdue / Active."),

    # Spine_Labels
    ("Spine_Labels", "A — Item ID",     '=Catalog!A2',  "Mirror catalog item ID."),
    ("Spine_Labels", "B — Call Number", '=Catalog!C2',  "Mirror catalog call number."),
    ("Spine_Labels", "C — Title",       '=Catalog!D2',  "Mirror title."),
    ("Spine_Labels", "D — Volume",      '=Catalog!O2',  "Mirror volume."),
    ("Spine_Labels", "E — Creator",     '=Catalog!F2',  "Mirror author."),
    ("Spine_Labels", "F — Location",
     '=IF(Catalog!X2="","",Catalog!X2&" "&Catalog!Y2)',
     "Location Type + Location Identifier."),

    # Dublin_Core
    ("Dublin_Core", "Columns A–O",
     '=Catalog!D2 / F2 / (I2&J2) / Q2 / S2 / N2 / (K2&" / "&M2) / "Print" / identifier / L2 / description / AE2 / W2 / R2 / AG2',
     "Fifteen-column Dublin Core export view, aligned with Catalog."),

    # Data_Quality
    ("Data_Quality", "Summary metrics",
     '=COUNTA(...) or SUMPRODUCT(...)',
     "Counts missing, duplicate, and total records."),
    ("Data_Quality", "C — Issues",
     '=IF(Catalog!D2="","",'
     'IF(ISBLANK(Catalog!F2),"Missing Creator; ","") & ... )',
     "Concatenated warning string for each catalog row."),

    # Dashboard
    ("Dashboard", "COLLECTION",
     '=COUNTA(...) or COUNTIF(Catalog!K:K,"Book")',
     "Counts by item type."),
    ("Dashboard", "METADATA",
     '=SUMPRODUCT((Catalog!$D$...<>"")*(Catalog!$F$...=""))',
     "Completeness counters."),
    ("Dashboard", "CIRCULATION",
     '=COUNTIF(Catalog!AA:AA, "Available") etc.',
     "Available / On Loan / Overdue."),
    ("Dashboard", "DIGITISATION",
     '=COUNTIF(Archive!G:G, ...)',
     "Physical/Digital/Both distribution."),
]

for i, row in enumerate(formulas_doc, start=6):
    for j, val in enumerate(row, start=1):
        c = ws.cell(i, j, val)
        c.border = BORDER_THIN
        c.alignment = Alignment(vertical="top", wrap_text=True)
        if j == 1:
            c.font = Font(name=FONT_BODY, bold=True, size=SZ_SMALL, color=C_DARKSLATEBLUE)
        elif j == 2:
            c.font = Font(name=FONT_BODY, bold=True, size=SZ_SMALL, color=C_BROWN)
        elif j == 3:
            c.font = Font(name=FONT_MONO, size=SZ_META, color=C_DARKSLATEGRAY)
        else:
            c.font = Font(name=FONT_SMALL, italic=True, size=SZ_META, color=C_PERU)
    ws.row_dimensions[i].height = 46


# ═══════════════════════════════════════════════════════════════════════
# 14. JSON_EXPORT
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("JSON_Export")
ws.column_dimensions["A"].width = 10
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 24
ws.column_dimensions["D"].width = 70

title_cell(ws["A1"]); ws["A1"] = "STRUCTURED EXPORT FOR JSON / CSV PIPELINE"
ws.merge_cells("A1:D1"); ws.row_dimensions[1].height = 32

small_cell(ws["A3"])
ws["A3"] = "Flattened view of Catalog for future JSON conversion. Field names are snake_case."
ws.merge_cells("A3:D3")

for j, name in enumerate(["item_id", "title", "creator", "json_snippet"], start=1):
    c = ws.cell(5, j, name)
    c.font = Font(name=FONT_MONO, bold=True, color="FFFFFF", size=SZ_SMALL)
    c.fill = PatternFill("solid", fgColor=C_DARKSLATEBLUE)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER_THIN

for r in range(2, ROWS_CAT + 2):
    row = r + 4
    ws.cell(row, 1, f'=Catalog!A{r}').font = Font(name=FONT_MONO, size=SZ_SMALL, color=C_DARKSLATEGRAY)
    ws.cell(row, 2, f'=Catalog!D{r}').font = Font(name=FONT_BODY, size=SZ_SMALL, color=C_DARKSLATEGRAY)
    ws.cell(row, 3, f'=Catalog!F{r}').font = Font(name=FONT_BODY, size=SZ_SMALL, color=C_DARKSLATEGRAY)
    c4 = ws.cell(row, 4, (
        f'=IF(Catalog!D{r}="","",'
        f'"{{""item_id"":"&Catalog!A{r}&",'
        f'""title"":"""&SUBSTITUTE(Catalog!D{r},"""","\\""")&""",'
        f'""creator"":"""&SUBSTITUTE(Catalog!F{r},"""","\\""")&""",'
        f'""language"":"""&Catalog!N{r}&""",'
        f'""identifier"":"""&Catalog!C{r}&"""}}"'
        f')'
    ))
    c4.font = Font(name=FONT_MONO, size=SZ_SMALL, color=C_DARKSLATEBLUE)
    c4.fill = PatternFill("solid", fgColor=C_GENERATED)
ws.freeze_panes = "A6"


# ═══════════════════════════════════════════════════════════════════════
# 15. README
# ═══════════════════════════════════════════════════════════════════════
ws = wb.create_sheet("README", 0)
ws.column_dimensions["A"].width = 4
ws.column_dimensions["B"].width = 108

readme_lines = [
    ("TITLE", "Personal Library Manager & Digital Archive"),
    ("H2", "Version 3.1 — Excel Application"),
    ("P", "Excel is the authoritative structured-data and metadata layer. "
          "The web application is a separate presentation and printing layer."),

    ("H3", "Quick Start"),
    ("P", "1. Open Settings; set library name, loan period, defaults."),
    ("P", "2. Catalog — books, magazines, journals, newspapers, manuscripts, ephemera (print only)."),
    ("P", "3. Digital_Catalog — E-books, articles, images, photographs, audio, video, manuscripts (digital)."),
    ("P", "4. Archive — archival objects: sculptures, coins, textiles, photographs, mementos, "
          "manuscripts, and any physical or digital artefact that is not a regular book."),
    ("P", "5. Circulation — loans; Catalog status updates automatically."),
    ("P", "6. Formulas sheet documents every generated cell."),

    ("H3", "Three separate Item Type vocabularies"),
    ("P", "Catalog        — Book, Magazine, Journal, Periodical, Newspaper, Newsletter, "
          "Annual, Special Issue, Pamphlet, Brochure, Manuscript, Personal Papers, Ephemera, Other."),
    ("P", "Digital        — E-Book, E-Journal, E-Magazine, Article, Image, Photograph, Poster, "
          "Map, Manuscript, Sound Recording, Audio Recording, Video Recording, AV Recording, "
          "Dataset, Website, Software, Other."),
    ("P", "Archive        — Manuscript, Newspaper, Newspaper Clipping, Photograph, Poster, Map, "
          "Painting, Sculpture, Textile, Coin, Medal, Furniture, Memento, Personal Papers, "
          "Ephemera, Pamphlet, Brochure, Sound Recording, Audio Recording, Video Recording, "
          "AV Recording, Other."),
    ("P", "Each vocabulary lives in its own code table in Lookup_Codes and has its own dropdown."),

    ("H3", "Subject & Genre / Form (shared across sheets)"),
    ("P", "Subject      — Literature, History, Philosophy, Science & Technology, Politics, "
          "Economics, Religion & Mythology, Art & Cinema, Music, Language, Medicine, Law, "
          "Education, Biography, Reference, Other."),
    ("P", "Genre / Form — Novel, Poem, Play, Short Story, Essay, Article, Biography, "
          "Autobiography, Travelogue, Science Fiction, Fantasy, Mystery, Thriller, Romance, "
          "Children's, Reference, Collected Works, Anthology, Interview, Other."),

    ("H3", "Call Number System"),
    ("P", "Format:  ITEMTYPE-SUBJECT-GENRE-AUTHORCODE-ITEMNUMBER[-VOLUME]"),
    ("P", "Example: BK-LIT-NOV-TAG-001-Vol1"),
    ("P", "Codes are read from Lookup_Codes via VLOOKUP. If a value is not in the code table, "
          "the formula shows '?' in that slot. A missing author becomes 'NOC'."),
    ("P", "Digital_Catalog uses the same pattern with a DIG- prefix."),

    ("H3", "Author Code Logic"),
    ("P", "Author Code uses the last word of Creator (surname for most names). "
          "'Rabindranath Tagore' → TAG. Use Author Code Override to supply a manual code."),

    ("H3", "Location System — semi-controlled"),
    ("P", "Location Type: dropdown (Shelf, Magazine Rack, Journal Rack, Archive, Cabinet, Box, …)."),
    ("P", "Location Identifier: free text (A1, A2, MR1, JR1, ARCH-01, BOX-07)."),

    ("H3", "Archive — Extended Fields"),
    ("P", "Collection, Dimensions, Material, Place, Subject / Keywords, Digital File Name, "
          "Public Link (hyperlink), Media Preview (auto-generated)."),

    ("H3", "Typography"),
    ("P", "Titles          — Garamond"),
    ("P", "Section headers — Baskerville Old Face"),
    ("P", "Body text       — Book Antiqua (Caslon substitute)"),
    ("P", "Call numbers    — Courier New"),

    ("H3", "Colour System"),
    ("P", "DarkSlateBlue #483D8B — primary headers, generated values"),
    ("P", "Brown         #A52A2A — secondary headers, alternates"),
    ("P", "Aquamarine    #7FFFD4 — metric values"),
    ("P", "Cream         #FFF8E7 — editable cells"),
    ("P", "Lavender      #E6E6FA — label backgrounds"),

    ("H3", "Protection"),
    ("P", "Generated columns are locked; input columns remain editable."),
    ("P", "Sheet protection password: 'library'. Change in Settings if desired."),

    ("H3", "Important Notes"),
    ("P", "• Accession Number is the permanent identifier. Item ID is internal."),
    ("P", "• Do not sort by columns other than Item ID — Call Number formulas reference row position."),
    ("P", "• Open in Excel and press Ctrl+S before uploading to the web app so formula values are cached."),
    ("P", "• No VBA macros. Compatible with Excel, Google Sheets, and LibreOffice Calc."),
]

row = 1
for kind, text in readme_lines:
    c = ws.cell(row, 2, text)
    if kind == "TITLE":
        c.font = Font(name=FONT_TITLE, bold=True, size=SZ_TITLE, color=C_DARKSLATEBLUE)
    elif kind == "H2":
        c.font = Font(name=FONT_SECTION, bold=True, size=SZ_H2, color=C_BROWN)
    elif kind == "H3":
        c.font = Font(name=FONT_SECTION, bold=True, size=SZ_H2, color=C_DARKSLATEBLUE)
    else:
        c.font = Font(name=FONT_BODY, size=SZ_BODY, color=C_DARKSLATEGRAY)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if kind == "P":
        ws.row_dimensions[row].height = 26
    if kind == "TITLE":
        ws.row_dimensions[row].height = 36
    if kind in ("H2", "H3"):
        ws.row_dimensions[row].height = 24
    row += 1


# ═══════════════════════════════════════════════════════════════════════
# 16. SHEET PROTECTION
# ═══════════════════════════════════════════════════════════════════════
PW = "library"
for sname in ["Catalog", "Digital_Catalog", "Archive", "Circulation",
              "Members", "Spine_Labels", "Dublin_Core", "Data_Quality",
              "Dashboard", "JSON_Export", "Formulas"]:
    wb[sname].protection.sheet = True
    wb[sname].protection.password = PW

wb["Settings"].protection.sheet = True
wb["Settings"].protection.password = PW


# ═══════════════════════════════════════════════════════════════════════
# 17. SAVE
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
print("Item Type vocabularies:")
print("  Catalog  →", ", ".join(v for v, _ in cat_item_types[:5]), "…")
print("  Digital  →", ", ".join(v for v, _ in dig_item_types[:5]), "…")
print("  Archive  →", ", ".join(v for v, _ in arc_item_types[:5]), "…")
print()
print("Typography: Garamond · Baskerville Old Face · Book Antiqua · Courier New")
print("Colours:    DarkSlateBlue · Brown · Aquamarine · Cream · Lavender")
print()
print("Next: open the workbook in Excel and press Ctrl+S to cache formula values.")
print("=" * 68)
