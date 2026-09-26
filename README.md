📚 Personal Library Manager & Spine Studio

<p align="center">
  <strong>A lightweight library catalogue, digital-archive and spine-label toolkit.</strong>
</p>

<p align="center">
  <a href="https://masudranamondal1-lgtm.github.io/library-spine-studio/"><img src="https://img.shields.io/badge/🌐_Live_App-Open-2F6F6D?style=for-the-badge" alt="Live App"></a>
  <a href="https://github.com/masudranamondal1-lgtm/library-spine-studio"><img src="https://img.shields.io/github/stars/masudranamondal1-lgtm/library-spine-studio?style=for-the-badge" alt="GitHub Stars"></a>
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-7A5C3E?style=for-the-badge" alt="MIT License"></a>
</p>

<p align="center">
  📚 Physical Collection · 💾 Digital Archive · 🏷️ Dublin Core · 🖨️ A4 Labels · 🔒 Local-First
</p>

<h2>🧭 Overview</h2>

Personal Library Manager & Spine Studio combines a structured Excel catalogue with a browser-based interface for searching, browsing, metadata lookup, and printing spine labels.

[!IMPORTANT]
Excel is the structured data layer. The web application is the discovery and printing layer.

It is designed for personal libraries, researchers, collectors, writers, and small libraries without requiring a database server or subscription.

<h3>✨ Core Features</h3>

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

<h2>🚀 Quick Start</h2>

<h3>1. Prepare the workbook</h3>

Download Personal Library Template.xlsx, open it in Excel or another compatible spreadsheet application, and enter your collection in the Catalog sheet.

Typical fields include:

Title · Creator · Type · Genre · Language · Volume · Publisher · ISBN · Location · Condition

Save the workbook after editing so spreadsheet formulas are recalculated.

<h3>2. Open Spine Studio</h3>

🌐 https://masudranamondal1-lgtm.github.io/library-spine-studio/

Choose Local File and load your .xlsx or .csv.

<h3>3. Browse and print</h3>

Search or filter your collection, select a label style, preview the page, and print to A4 or save as PDF.

[!TIP]
Always print a test page on ordinary paper before using adhesive label stock.

<h2>🗂️ Workbook Architecture</h2>

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

<h2>📚 Collection Model</h2>

<h3>Physical</h3>

Books, magazines, journals, periodicals, newsletters, printed reference material and other publications.

<h3>Digital</h3>

PDF, EPUB, MOBI, DJVU, TXT, DOCX, HTML, XML, TEI XML, Markdown, images, audio, video, datasets and other digital resources.

<h3>Archive</h3>

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

<h2>🏷️ Metadata & Interoperability</h2>

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

<h2>🔢 Call Numbers</h2>

The project uses a readable, project-specific structure:

TYPE-GENRE-AUTHOR-ITEM-VOLUME

Example:

FIC-LIT-TAG-001-Vol1

These are not Dewey Decimal or Library of Congress classifications.

Author codes can be generated automatically and manually overridden when necessary.

[!WARNING]
Permanent accession numbers should identify an item, not a spreadsheet row. Verify identifiers before sorting, deleting, or migrating records.

<h2>📍 Locations</h2>

Location Type uses controlled values such as:

Shelf · Magazine Rack · Journal Rack · Archive · Cabinet · Box · Other

The actual Location Identifier remains user-defined:

A2 · SHELF-03 · MR1 · ARCH-01 · BOX-07

This keeps the system flexible as the physical library changes.

<h2>🖨️ Spine Labels</h2>

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

<h2>🌐 Web Application</h2>

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

<h2>🔒 Privacy</h2>

The local workflow does not require an account or application server. Local spreadsheet parsing happens in the browser.

Optional network-connected features may contact:

Google Sheets

Google Books

external CDN resources

[!CAUTION]
Never publish private borrower information, private notes, unpublished material, or other sensitive records in a public Google Sheet or GitHub repository.

<h2>🛠️ Technology</h2>

Frontend: HTML5, CSS, Vanilla JavaScript

Spreadsheet parsing: SheetJS

Workbook generation: Python + openpyxl

Metadata lookup: Google Books API

Hosting: GitHub Pages

Offline support: Service Worker / browser caching

License: MIT

No database server or build pipeline is required for the basic application.

<h2>🐍 Generate the Workbook</h2>

pip install openpyxl
python build_library.py

For a blank template:

python build_library.py --blank

<h2>🗺️ Roadmap</h2>

The project is evolving toward a more complete personal library and digital-archive environment.

<h3>Current development</h3>

areas

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

<h2>🤝 Contributing</h2>

Bug reports, documentation improvements, and code contributions are welcome.

When reporting an issue, include:

What you were doing

Browser/spreadsheet application

Workbook version

Expected behaviour

Actual behaviour

Please do not upload private library data to public issues.

<h2>📜 License</h2>

Released under the MIT License.

See LICENSE for the complete terms.

<h2>☕ Support</h2>

If the project is useful:

Buy Me a Coffee: https://buymeacoffee.com/masudshaon

GitHub Issues: https://github.com/masudranamondal1-lgtm/library-spine-studio/issues

<h2>👤 Author</h2>

<h3 align="center">Masud Rana Mondal</h3>
<p align="center">Writer · Researcher · Digital Humanities · Publishing · Library & Archive Technology</p>

<p align="center">
  <a href="https://github.com/masudranamondal1-lgtm">
    <img src="https://img.shields.io/badge/GitHub-Masud%20Rana-181717?style=for-the-badge&logo=github" alt="GitHub">
  </a>
</p>
