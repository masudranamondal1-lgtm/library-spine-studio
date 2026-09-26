Personal Library Manager & Spine Studio

A browser-based personal library cataloguing, discovery, and spine-label printing toolkit, backed by a structured Excel workbook.

Catalog books and other resources, generate consistent identifiers and call numbers, browse your collection in the browser, and prepare A4 spine labels without running a database server.

Live application: https://masudranamondal1-lgtm.github.io/library-spine-studio/

Overview

Personal Library Manager & Spine Studio is designed for personal libraries, researchers, collectors, writers, small libraries, and anyone who has accumulated enough books to discover that human memory is, regrettably, not a reliable cataloguing system.

The project separates the work into two complementary layers:

Layer

Purpose

Excel workbook

Structured cataloguing, identifiers, classification, locations, metadata, circulation, archive records, and data-quality management

Web application

Search, browse, filter, import, ISBN-assisted lookup, spine-label preview, and A4 printing

The workbook is the structured data layer. The browser application is the presentation and printing layer.

The application can work with local .xlsx and .csv files. It can also read a supported published Google Sheet when the user explicitly connects one.

Project Status

The project is under active development.

The repository currently contains the web application, a blank Excel template, the workbook generator, PWA assets, and project documentation. The workbook architecture is being expanded from a simple book catalogue into a broader Personal Library Manager & Digital Archive model.

Core Design Principles

Structured data first

The Excel workbook is not merely a list of books. It is intended to be a structured collection-management layer.

Dublin Core first

Dublin Core is the primary interoperability metadata framework. The workbook preserves bibliographic and archival information in a form that can later be exported or transformed into other standards and formats.

Physical and digital collections together

The project supports books, magazines, journals, periodicals, manuscripts, photographs, newspaper clippings, maps, posters, pamphlets, ephemera, PDFs, ebooks, scanned documents, audio, video, images, datasets, and other digital resources.

A physical object and its digital surrogate can be represented as related records rather than being forced into the same storage model.

Human-readable identifiers

The project uses project-specific identifiers and call numbers. It does not claim to implement Dewey Decimal Classification or Library of Congress Classification.

Semi-controlled physical locations

Location Type uses controlled values such as Shelf, Magazine Rack, Journal Rack, Archive, Cabinet, Box, and Other. The actual Location Identifier remains user-defined, for example A2, MR1, ARCH-01, or BOX-07.

Local-first operation

Local Excel/CSV catalogues can be loaded directly into the browser. No dedicated application server is required for the basic workflow.

How It Works

        Excel Workbook
             |
             | .xlsx / .csv
             v
      +------------------+
      |   Spine Studio   |
      | Search / Browse  |
      | Filter / ISBN    |
      | Labels / Print   |
      +--------+---------+
               |
               v
          A4 Spine Labels

A supported published Google Sheet can also be used as a read-only source. The application does not automatically write changes back to the original Excel workbook or Google Sheet.

Quick Start

1. Download the workbook

Download Personal Library Template.xlsx from the repository.

2. Open the workbook

Use Microsoft Excel, LibreOffice Calc, or another compatible spreadsheet application.

3. Enter your collection

Use the Catalog sheet for physical items. Typical fields include Title, Creator, Type, Genre, Language, Volume, Publisher, Publication Year, ISBN/ISSN, Location, Condition, Acquisition information, Rights, Notes, and Description.

4. Save the workbook

Save after editing so spreadsheet formulas are recalculated and cached values are available to the browser application.

5. Open Spine Studio

Visit https://masudranamondal1-lgtm.github.io/library-spine-studio/

6. Import the workbook

Use Local File and select the .xlsx file. Search, filter, sort, inspect records, and prepare labels.

7. Print

Choose a label style, preview it, and print to A4 or save the output as PDF. Test alignment on ordinary paper before using adhesive stock.

Workbook Architecture

The current workbook is designed as a multi-sheet collection-management system.

Sheet

Purpose

README

Workbook documentation

Settings

Configurable library and metadata parameters

Lookup_Codes

Controlled vocabularies and dropdown values

Catalog

Physical collection and bibliographic records

Digital_Catalog

Digital resources and digital-resource metadata

Archive

Archival objects, including physical and digital material

Circulation

Loans, borrowers, due dates, and returns

Members

Borrower/member records

Spine_Labels

Derived label data

Dublin_Core

Dublin Core mapping and metadata/export structure

Data_Quality

Missing-data and integrity checks

Dashboard

Collection-level summary information

JSON_Export

Structured export-oriented data

The exact workbook version in the repository remains authoritative for the fields and formulas available in that release.

Collection Types

Physical Collection

The physical catalog can represent books, magazines, journals, periodicals, newsletters, annuals, special issues, and other printed resources. Records can include identifiers, call numbers, bibliographic metadata, physical location, condition, acquisition information, and circulation status.

Digital Collection

The digital catalog is intended for PDF, EPUB, MOBI, DJVU, TXT, DOCX, HTML, XML, TEI XML, Markdown, JPG, PNG, TIFF, WebP, MP3, WAV, FLAC, M4A, MP4, MKV, MOV, WebM, datasets, spreadsheets, presentations, websites, software/code, and other digital resources.

Digital records can include format, MIME type, file extension, file size, storage path, URI, OCR status, searchability, rights, access level, checksum, preservation status, reading/usage status, and relationships to physical or archival records.

Archive

The archive is not restricted to digitised material. It can contain physical archival objects such as newspaper clippings, manuscripts, photographs, posters, maps, pamphlets, brochures, personal papers, ephemera, rare printed materials, and historical books, as well as their digital surrogates.

Archival records can include provenance, acquisition, condition, physical/digital status, archive location, digitisation status, OCR status, rights, access level, and related-resource identifiers.

Metadata

Dublin Core

Dublin Core is the project's primary metadata framework. The implementation is intended to support core elements such as:

Title

Creator

Contributor

Publisher

Date

Subject

Description

Type

Format

Identifier

Source

Language

Relation

Coverage

Rights

Project-specific cataloguing fields remain available where collection management requires information beyond the core Dublin Core vocabulary.

Interoperability

The workbook is designed with future interoperability in mind, including CSV, JSON, Dublin Core, TEI-related workflows, bibliographic APIs, external identifiers, and future digital-archive platforms.

TEI is treated as an interoperability and scholarly-text pathway rather than as a replacement for Dublin Core.

Call Numbers

The project uses a readable, project-specific structure:

TYPE-GENRE-AUTHOR-ITEM-VOLUME

Example:

FIC-LIT-TAG-001-Vol1

These call numbers are project-specific. They are not official Dewey Decimal or Library of Congress classifications, and author abbreviations should not be confused with standard Cutter numbers.

Author Codes and Identifiers

The workbook supports automatic author-code generation and manual override. For example, Rabindranath Tagore may produce TAG.

The project distinguishes between item identifiers, accession numbers, call numbers, ISBN/ISSN, DOI, URI, and other external identifiers.

A permanent accession number should identify the item rather than the current spreadsheet row. Do not assume that a row-derived formula remains stable after sorting, inserting, deleting, or migrating records.

Locations

The location model is intentionally semi-controlled.

Location Type: Shelf, Magazine Rack, Journal Rack, Archive, Cabinet, Box, Other.

Location Identifier: user-defined values such as A1, S3, MR1, JR2, ARCH-01, or BOX-04.

This keeps the system practical when the physical arrangement of a personal collection changes.

Circulation

The workbook includes a basic lending model covering Loan ID, Item ID, Member ID, borrower, checkout date, due date, return date, status, and notes. The default loan period is configurable.

This is intended for personal and small-scale use, not as a replacement for a full institutional ILS.

Spine Label Printing

The web application provides printable label layouts, including:

Standard Card

Library Block

Minimal Spine

The application targets A4 printing with multiple labels per page.

Before printing on adhesive stock:

Print a test page on ordinary A4 paper.

Use 100% / Actual Size.

Disable browser headers and footers.

Check paper size and orientation.

Compare the printout against the actual label sheet.

Adjust the layout if necessary.

Do not use Fit to Page to conceal a dimensional mismatch.

Web Application Features

Catalog search

Filtering and sorting

Local Excel import

CSV import

Read-only Google Sheets source support

Manual local entry

ISBN scanning on supported mobile browsers

Google Books metadata lookup where available

Call-number and location display

Label preview

A4 print layout

PWA installation support

Browser-side caching where supported

ISBN Scanner

On supported mobile browsers, the application can use the browser's barcode capabilities to detect ISBNs. When available, the ISBN can be used to request bibliographic metadata from Google Books.

Camera permission, browser support, HTTPS/secure context, and a readable barcode are required. External metadata should always be checked before it becomes part of the authoritative catalog.

Data Sources

Source

Behaviour

Local Excel

Import an .xlsx file in the browser

Local CSV

Import a .csv file in the browser

Published Google Sheet

Read-only external catalog source

Manual entry

Local working records in the application

Cached catalog

Previously available data where browser storage permits

Google Sheets integration is read-only from the application's perspective. Changes made in the application are not automatically written back to the source sheet.

Do not publish private borrower information or sensitive collection data to a public sheet.

Privacy

The basic local workflow does not require an account, database server, or subscription. Local spreadsheet parsing occurs in the browser.

Optional network-connected features may communicate with third parties, including Google Sheets, Google Books, and external CDN resources used by the application. Users should therefore distinguish between local cataloguing and optional connected features.

Technology

Frontend: HTML5, CSS, Vanilla JavaScript

Spreadsheet parsing: SheetJS

Metadata lookup: Google Books API

Hosting: GitHub Pages

Offline/PWA: Web App Manifest and Service Worker

Workbook generation: Python and openpyxl

The application has no dedicated backend or database server.

Python Workbook Generator

To regenerate the workbook programmatically:

pip install openpyxl
python build_library.py

To generate a blank template:

python build_library.py --blank

The generator keeps workbook construction reproducible rather than requiring manual creation of every worksheet.

Development

Clone the repository:

git clone https://github.com/masudranamondal1-lgtm/library-spine-studio.git
cd library-spine-studio

Main project files:

library-spine-studio/
├── .github/
├── index.html
├── Personal Library Template.xlsx
├── build_library.py
├── manifest.json
├── service-worker.js
├── donate.html
├── qr-code.png
├── LICENSE
└── README.md

Open index.html for basic testing. Use a local HTTP server when testing service-worker/PWA behaviour.

Repository Safety

The public repository should contain software and blank templates, not private personal catalogues.

Do not commit:

personal library workbooks

borrower/member information

private notes

private archival metadata

unpublished manuscripts

private file paths

credentials or API keys

If a private workbook has previously been committed, deleting the current copy does not necessarily remove it from Git history. A history cleanup may be required.

Roadmap

The project is evolving toward a more complete personal library and digital-archive environment. Development areas include:

stronger Dublin Core implementation

improved digital-resource ingestion

richer archive relationships

physical/digital object relationships

bibliographic API enrichment

stronger identifier management

improved data-quality validation

structured CSV/JSON export

TEI interoperability workflows

metadata reconciliation

tighter catalog-to-web-app integration

more flexible label layouts

richer dashboard and collection analytics

The goal is not to reproduce a large institutional library-management system. It is to create a portable, understandable, standards-aware personal library and archive system that can grow with a serious research collection.

Contributing

Bug reports, documentation improvements, and code contributions are welcome.

When reporting a problem, include:

What you were trying to do

Browser or spreadsheet application used

Workbook version

Expected result

Actual result

Screenshot or minimal example when useful

Do not upload private library data to public issues.

License

Released under the MIT License. See LICENSE for the complete terms.

Credits

The project uses or is compatible with:

SheetJS for browser-side spreadsheet parsing

Google Books API for optional ISBN metadata lookup

GitHub Pages for static hosting

Python / openpyxl for workbook generation

Support

If the project is useful, support options are available through the application.

Buy Me a Coffee: https://buymeacoffee.com/masudshaon

GitHub Issues: https://github.com/masudranamondal1-lgtm/library-spine-studio/issues

Author

Masud Rana Mondal

Writer · Researcher · Digital Humanities · Publishing · Library & Archive Technology

GitHub: https://github.com/masudranamondal1-lgtm
