📚 Personal Library Manager & Spine Studio

<p align="center">
  <strong>A lightweight library catalogue, digital-archive and spine-label toolkit.</strong>
</p>

<p align="center">
  <a href="https://masudranamondal1-lgtm.github.io/library-spine-studio/"><img src="https://img.shields.io/badge/🌐_Live_App-Open-2F6F6D?style=for-the-badge" alt="Live App"></a>
  <a href="https://github.com/masudranamondal1-lgtm/library-spine-studio"><img src="https://img.shields.io/github/stars/masudranamondal1-lgtm/library-spine-studio?style=for-the-badge&label=⭐%20Stars" alt="GitHub Stars"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-7A5C3E?style=for-the-badge" alt="MIT License"></a>
</p>

<p align="center">
  📚 Physical Collection · 💾 Digital Archive · 🏷️ Dublin Core · 🖨️ A4 Labels · 🔒 Local-First
</p>

🧭 Overview

Personal Library Manager & Spine Studio combines a structured Excel catalogue with a browser-based interface for searching, browsing, metadata lookup, and printing spine labels.

[!IMPORTANT]
Excel is the structured data layer. The web application is the discovery and printing layer.

It is designed for personal libraries, researchers, collectors, writers, and small libraries without requiring a database server or subscription.

✨ Core Features

📚 Physical books, magazines, journals and periodicals

💾 Digital resources including ebooks, scans, audio, video and datasets

🗂️ Physical and digital archival objects

🏷️ Dublin Core-based metadata

🔢 Project-specific call numbers and author codes

📍 Semi-controlled shelf/location management

🔄 Basic circulation and member management

🖨️ A4 spine-label printing

📷 ISBN scanning on supported mobile browsers

📤 CSV / JSON / interoperability-ready data

🌐 Optional read-only Google Sheets source

🚀 Quick Start

1. Prepare the workbook

Download Personal Library Template.xlsx, open it in Excel or another compatible spreadsheet application, and enter your collection in the Catalog sheet.

Typical fields include:

Title · Creator · Type · Genre · Language · Volume · Publisher · ISBN · Location · Condition

Save the workbook after editing so spreadsheet formulas are recalculated.

2. Open Spine Studio

🌐 https://masudranamondal1-lgtm.github.io/library-spine-studio/

Choose Local File and load your .xlsx or .csv.

3. Browse and print

Search or filter your collection, select a label style, preview the page, and print to A4 or save as PDF.

[!TIP]
Always print a test page on ordinary paper before using adhesive label stock.

🗂️ Workbook Architecture

Sheet

Purpose

Catalog

Physical collection

Digital_Catalog

Digital resources

Archive

Physical and digital archival objects

Dublin_Core

Metadata mapping/export

Circulation

Loans and returns

Members

Borrowers

Lookup_Codes

Controlled vocabularies

Spine_Labels

Label data

Data_Quality

Integrity checks

Dashboard

Collection overview

JSON_Export

Structured export

Settings

Configuration

The workbook is intended to remain the authoritative catalogue.

📚 Collection Model

Physical

Books, magazines, journals, periodicals, newsletters, printed reference material and other publications.

Digital

PDF, EPUB, MOBI, DJVU, TXT, DOCX, HTML, XML, TEI XML, Markdown, images, audio, video, datasets and other digital resources.

Archive

The archive can contain both physical originals and digital surrogates, including:

manuscripts

newspaper clippings

photographs

maps

posters

pamphlets

personal papers

ephemera

rare or historical printed material

Records can include provenance, condition, location, digitisation status, OCR, rights and access information.

🏷️ Metadata & Interoperability

Dublin Core is the primary metadata framework.

Core elements include:

Title · Creator · Contributor · Publisher · Date · Subject · Description · Type · Format · Identifier · Source · Language · Relation · Coverage · Rights

Project-specific catalogue fields remain available where Dublin Core alone is insufficient.

The system is designed for future:

CSV export

JSON export

TEI workflows

bibliographic API enrichment

external identifiers

digital-archive interoperability

🔢 Call Numbers

The project uses a readable, project-specific structure:

TYPE-GENRE-AUTHOR-ITEM-VOLUME

Example:

FIC-LIT-TAG-001-Vol1

These are not Dewey Decimal or Library of Congress classifications.

Author codes can be generated automatically and manually overridden when necessary.

[!WARNING]
Permanent accession numbers should identify an item, not a spreadsheet row. Verify identifiers before sorting, deleting, or migrating records.

📍 Locations

Location Type uses controlled values such as:

Shelf · Magazine Rack · Journal Rack · Archive · Cabinet · Box · Other

The actual Location Identifier remains user-defined:

A2 · SHELF-03 · MR1 · ARCH-01 · BOX-07

This keeps the system flexible as the physical library changes.

🖨️ Spine Labels

The web application provides three label styles:

Standard Card

Library Block

Minimal Spine

The print workflow targets A4 sheets with multiple labels per page.

Recommended settings:

Setting

Value

Paper

A4

Orientation

Portrait

Scale

100% / Actual Size

Browser headers

Off

Fit to page

Off

🌐 Web Application

The browser application supports:

🔎 Search and filtering

📥 Excel / CSV import

🌐 Published Google Sheets

✍️ Manual local entry

📷 ISBN scanning

📖 Google Books metadata lookup

🖨️ Label preview and printing

📱 PWA installation

💾 Browser caching where supported

Google Sheets integration is read-only from the application.

🔒 Privacy

The local workflow does not require an account or application server. Local spreadsheet parsing happens in the browser.

Optional network-connected features may contact:

Google Sheets

Google Books

external CDN resources

[!CAUTION]
Never publish private borrower information, private notes, unpublished material, or other sensitive records in a public Google Sheet or GitHub repository.

🛠️ Technology

Frontend: HTML5, CSS, Vanilla JavaScript

Spreadsheet parsing: SheetJS

Workbook generation: Python + openpyxl

Metadata lookup: Google Books API

Hosting: GitHub Pages

Offline support: Service Worker / browser caching

License: MIT

No database server or build pipeline is required for the basic application.

🐍 Generate the Workbook

pip install openpyxl
python build_library.py

For a blank template:

python build_library.py --blank

🗺️ Roadmap

The project is evolving toward a more complete personal library and digital-archive environment.

Current development areas

🏷️ Stronger Dublin Core implementation

💾 Improved digital-resource ingestion

🗂️ Richer archive relationships

🔗 Physical/digital object relationships

🔎 Bibliographic API enrichment

🆔 Stronger identifier management

✅ Improved data-quality validation

📤 Better CSV/JSON export

📜 TEI interoperability

📊 Improved collection analytics

🖨️ More flexible label layouts

The goal is not to reproduce a large institutional ILS. It is to create a portable, understandable, standards-aware library and archive system for serious personal collections.

🤝 Contributing

Bug reports, documentation improvements, and code contributions are welcome.

When reporting an issue, include:

What you were doing

Browser/spreadsheet application

Workbook version

Expected behaviour

Actual behaviour

Please do not upload private library data to public issues.

📜 License

Released under the MIT License.

See LICENSE for the complete terms.

☕ Support

If the project is useful:

Buy Me a Coffee: https://buymeacoffee.com/masudshaon

GitHub Issues: https://github.com/masudranamondal1-lgtm/library-spine-studio/issues

👤 Author

<p align="center">
  <strong>Masud Rana Mondal</strong><br>
  Writer · Researcher · Digital Humanities · Publishing · Library & Archive Technology
</p>

<p align="center">
  <a href="https://github.com/masudranamondal1-lgtm">
    <img src="https://img.shields.io/badge/GitHub-Masud%20Rana-181717?style=for-the-badge&logo=github" alt="GitHub">
  </a>
</p>
