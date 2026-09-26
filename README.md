# 📚 Personal Library Manager & Spine Studio

<p align="center">
  <strong>A structured Excel catalogue + browser-based library and spine-label studio.</strong>
</p>

<p align="center">
  <a href="https://masudranamondal1-lgtm.github.io/library-spine-studio/">
    <img src="https://img.shields.io/badge/🌐_LIVE_APP-OPEN-2F6F6D?style=for-the-badge" alt="Live App">
  </a>
  <a href="https://github.com/masudranamondal1-lgtm/library-spine-studio">
    <img src="https://img.shields.io/github/stars/masudranamondal1-lgtm/library-spine-studio?style=for-the-badge" alt="GitHub Stars">
  </a>
  <a href="https://github.com/masudranamondal1-lgtm/library-spine-studio/blob/main/Personal%20Library%20Template.xlsx">
    <img src="https://img.shields.io/badge/📊_EXCEL-WORKBOOK-8B5E3C?style=for-the-badge" alt="Excel Workbook">
  </a>
  <a href="LICENSE">
    <img src="https://img.shields.io/badge/📜_LICENSE-MIT-6B4E71?style=for-the-badge" alt="MIT License">
  </a>
</p>

<p align="center">
  📊 <strong>Excel Catalogue</strong> ·
  🌐 <strong>Browser App</strong> ·
  🏷️ <strong>Dublin Core</strong> ·
  🖨️ <strong>A4 Labels</strong> ·
  📚 <strong>Physical + Digital</strong>
</p>

---

## 🧭 Overview

**Personal Library Manager & Spine Studio** combines a structured **Excel workbook** with a browser-based application for cataloguing, searching, browsing, metadata lookup, and spine-label printing.

### 🔄 One Workflow, Two Layers

<table>
<tr>
<th>📊 Excel Workbook</th>
<th>🌐 Spine Studio</th>
</tr>
<tr>
<td>
<strong>Catalogue & data layer</strong><br><br>
Identifiers · Metadata · Classification · Locations · Circulation · Archive
</td>
<td>
<strong>Interface & output layer</strong><br><br>
Search · Filter · Browse · ISBN Lookup · Preview · Print
</td>
</tr>
</table>

```text
📊 Excel Workbook
       │
       │  .xlsx / .csv
       ▼
🌐 Spine Studio
       │
       ├── 🔎 Search & Filter
       ├── 📖 Browse Catalogue
       ├── 📷 ISBN Lookup
       └── 🖨️ Preview & Print
```

> [!IMPORTANT]
> **The Excel workbook is the catalogue of record.** Edit your collection there, save it, and then load it into Spine Studio.

### 🔗 Project Links

| Resource | Link |
|---|---|
| 🌐 **Live Application** | [Open Spine Studio](https://masudranamondal1-lgtm.github.io/library-spine-studio/) |
| 📊 **Excel Workbook** | [Download `Personal Library Template.xlsx`](https://github.com/masudranamondal1-lgtm/library-spine-studio/blob/main/Personal%20Library%20Template.xlsx) |
| 💻 **GitHub Repository** | [View Repository](https://github.com/masudranamondal1-lgtm/library-spine-studio) |
| 🐛 **Issues** | [Report a bug / suggest an improvement](https://github.com/masudranamondal1-lgtm/library-spine-studio/issues) |
| 📜 **License** | [MIT License](https://github.com/masudranamondal1-lgtm/library-spine-studio/blob/main/LICENSE) |

---

## 🚀 Quick Start

### 📊 1. Download the Excel workbook

Download **[Personal Library Template.xlsx](https://github.com/masudranamondal1-lgtm/library-spine-studio/blob/main/Personal%20Library%20Template.xlsx)** and open it in Microsoft Excel, Google Sheets, or LibreOffice Calc.

### ✍️ 2. Catalogue your collection

Open the **`Catalog`** sheet and enter your records.

Typical fields:

`Title · Creator · Type · Genre · Language · Volume · Location Type · Location Identifier`

The workbook generates fields such as:

`Item ID · Author Code · Call Number · Circulation Status`

### 💾 3. Save the workbook

The web application reads **cached spreadsheet values**. It does not evaluate Excel formulas itself.

> [!WARNING]
> **Save the workbook after editing.** This allows your spreadsheet application to recalculate the formulas before Spine Studio imports the file.

### 🌐 4. Open Spine Studio

[**Launch the Live Application →**](https://masudranamondal1-lgtm.github.io/library-spine-studio/)

Choose **Local File** and load your `.xlsx` or `.csv`.

### 🖨️ 5. Browse and print

Search, filter, select a label style, preview the page, and print to A4 or save as PDF.

> [!TIP]
> Print a test page on ordinary paper before using adhesive label stock.

---

## ✨ Features

### 📚 Collection Management

- Physical books, magazines, journals and periodicals
- Digital resources: PDF, EPUB, MOBI, DJVU, TXT, DOCX, HTML, Markdown, audio, video, images and datasets
- Archive objects: physical and digital, with provenance and digitisation information

### 🏷️ Metadata

- **Dublin Core** as the primary metadata framework
- Mapping between catalogue fields and Dublin Core elements
- Project-specific extensions alongside standard metadata
- CSV and structured JSON export
- TEI-ready fields
- Preparation for bibliographic API enrichment

### 🔢 Cataloguing

- Automatic item numbering
- Surname-aware author codes
- Manual author-code override
- Project call numbers:

```text
TYPE-GENRE-AUTHORCODE-ITEM-VOLUME
```

Example:

```text
FIC-LIT-TAG-001-Vol1
```

These are **project-specific call numbers**, not Dewey Decimal or Library of Congress classifications.

### 📍 Locations

A semi-controlled location model:

**Location Type**

`Shelf · Magazine Rack · Journal Rack · Archive · Cabinet · Box · Other`

**Location Identifier**

`A2 · MR1 · ARCH-01 · BOX-07`

You do not need to pre-register every physical shelf.

### 🔄 Circulation

- Loans and borrowers
- Checkout and due dates
- Returns
- Automatic overdue detection
- Configurable loan period

### 🖨️ Printing

- A4 multi-label layout
- **Standard Card**
- **Library Block**
- **Minimal Spine**
- Shelf and accession information
- Optional colour-coded or monochrome output

### 🌐 Google Sheets

A published Google Sheet can be connected as a **read-only catalogue source**.

Changes made in the web application are not automatically written back to the spreadsheet.

---

## 📊 Excel Workbook

The workbook is designed to work **independently** as a structured library-management system.

### 🗂️ Workbook Structure

| Sheet | Purpose |
|---|---|
| `README` | Workbook documentation |
| `Settings` | Configurable parameters |
| `Lookup_Codes` | Controlled vocabularies |
| `Catalog` | Physical collection |
| `Digital_Catalog` | Digital resources |
| `Archive` | Physical and digital archival objects |
| `Circulation` | Loans and borrowers |
| `Members` | Borrower registry |
| `Spine_Labels` | Derived label data |
| `Dublin_Core` | Metadata mapping and export |
| `Data_Quality` | Missing-data and integrity checks |
| `Dashboard` | Collection overview |
| `JSON_Export` | Structured export view |

> **The workbook can be used without the web application.** The web application is an interface around the catalogue, not a replacement for it.

---

## 🗂️ Collection Model

### 📚 Physical Collection

Books, magazines, journals, periodicals and other printed materials.

### 💾 Digital Collection

Digital resources including:

`PDF · EPUB · MOBI · DJVU · TXT · DOCX · HTML · XML · TEI XML · Markdown · Images · Audio · Video · Datasets`

### 🗃️ Archive

The archive supports both **physical originals and digital surrogates**, including:

- manuscripts
- newspaper clippings
- photographs
- maps
- posters
- pamphlets
- personal papers
- ephemera
- rare or historical printed material

---

## 🏷️ Metadata & Interoperability

**Dublin Core is the primary metadata framework.**

Core elements include:

`Title · Creator · Contributor · Publisher · Date · Subject · Description · Type · Format · Identifier · Source · Language · Relation · Coverage · Rights`

The project is designed for:

- 📤 CSV export
- 🧾 JSON export
- 📜 TEI workflows
- 🔎 bibliographic API enrichment
- 🔗 external identifiers
- 🗂️ digital-archive interoperability

---

## 📱 ISBN Scanner

On supported mobile browsers, Spine Studio can scan an ISBN using browser barcode capabilities and request available bibliographic metadata through **Google Books**.

[Google Books API documentation →](https://developers.google.com/books)

> [!NOTE]
> External metadata should be reviewed before being added to the authoritative catalogue.

---

## 🔒 Privacy

Local `.xlsx` and `.csv` files are processed in the browser.

The basic local workflow does not require:

- an account
- a database server
- a subscription

Optional features may connect to:

- Google Sheets
- Google Books
- external CDN resources

> [!CAUTION]
> Do not publish private borrower information, private notes, unpublished material, or other sensitive records in a public Google Sheet or GitHub repository.

---

## 💻 Compatibility

| Platform | Support |
|---|---|
| 🪟 Microsoft Excel | Full |
| 📊 Google Sheets | Catalogue and spreadsheet workflow |
| 🟢 LibreOffice Calc | Catalogue workflow |
| 🌐 Chrome / Edge | Full web-app support |
| 🦊 Firefox / Safari | Web-app support |
| 📱 Mobile browsers | Web-app support; ISBN scanning where supported |

The workbook uses formulas, validation and conditional formatting rather than VBA macros.

---

## 🗂️ Project Structure

```text
library-spine-studio/
├── index.html
├── Personal Library Template.xlsx
├── build_library.py
├── manifest.json
├── service-worker.js
├── qr-code.png
├── donate.html
├── LICENSE
└── README.md
```

The web application is the **front end**. The workbook is the **data layer**.

---

## 🐍 Generate a Workbook with Python

For developers:

```bash
pip install openpyxl
python build_library.py
```

For a blank template:

```bash
python build_library.py --blank
```

---

## 🌍 Supported Languages

The workbook and application support Unicode text, including:

- English
- Bengali `বাংলা`
- Hindi `हिन्दी`
- Other Unicode scripts

Bengali titles and author names are preserved as entered.

---

## 🤝 Contributing

Bug reports, documentation improvements and pull requests are welcome.

When reporting an issue, include:

1. What you were doing
2. Browser or spreadsheet application
3. Workbook version
4. Expected behaviour
5. Actual behaviour

Please do not upload private library data to public issues.

[🐛 Open a GitHub Issue →](https://github.com/masudranamondal1-lgtm/library-spine-studio/issues)

---

## ☕ Support

If the project is useful:

[☕ Buy Me a Coffee →](https://buymeacoffee.com/masudshaon)

---

## 📜 License

Released under the **MIT License**.

[Read the full MIT License →](https://github.com/masudranamondal1-lgtm/library-spine-studio/blob/main/LICENSE)

---

## 🙏 Credits

Built with:

- **[SheetJS](https://sheetjs.com/)** — browser-side spreadsheet parsing
- **[Google Books API](https://developers.google.com/books)** — optional ISBN lookup
- **Python / [openpyxl](https://openpyxl.readthedocs.io/)** — workbook generation
- **GitHub Pages** — hosting
- **Noto Sans Bengali / Nirmala UI** — Unicode rendering

Designed for personal and small-scale library use, not as a replacement for a full institutional ILS.

---

## 📝 Changelog

### Version 1.0.5

- Extended catalogue field recognition
- Combined Location Type + Location Identifier for display
- Added empty-call-number warning
- Improved Catalog sheet detection

### Version 1.0.0

- Initial public release
- Physical, digital and archive cataloguing
- A4 spine-label printing
- Dublin Core metadata mapping

---

## 🗺️ Next Steps

1. 📊 Download and populate the Excel workbook.
2. 🌐 Load it into Spine Studio.
3. 🔎 Search and organise your collection.
4. 🖨️ Generate and test your spine labels.
5. 🗂️ Keep the workbook as your authoritative catalogue.

---

<h2 align="center">👤 Author</h2>

<h3 align="center">Masud Rana Mondal</h3>

<p align="center">
  Writer · Researcher · Digital Humanities · Publishing · Library & Archive Technology
</p>

<p align="center">
  <a href="https://github.com/masudranamondal1-lgtm">
    <img src="https://img.shields.io/badge/GitHub-Masud%20Rana-181717?style=for-the-badge&logo=github" alt="GitHub">
  </a>
</p>
