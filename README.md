# 📚 Personal Library Manager & Spine Studio

**Catalog your books. Generate consistent call numbers. Print library labels directly from your browser.**

[**🌐 Launch the live app**](https://masudranamondal1-lgtm.github.io/library-spine-studio/) · [**📥 Download the Excel workbook**](./Personal%20Library%20Catalogue%20%26%20Accession%20Number%20Generator%20(Enhanched).xlsx) · [**🐛 Report an issue**](https://github.com/masudranamondal1-lgtm/library-spine-studio/issues)

A lightweight library-management toolkit for collectors, researchers, book lovers, schools, and small libraries. Use a structured Excel workbook to manage your catalog and generate call numbers, then use the browser-based Spine Studio to find books and prepare printable labels. No account or dedicated server required.

> **Two tools, one workflow:** Excel manages your catalog and classification; the web app handles discovery, import, label design, and printing.

## Contents

- [Why this project exists](#why-this-project-exists)
- [How it works](#how-it-works)
- [Excel catalog and accession-number generator](#excel-catalog-and-accession-number-generator)
- [Web application](#web-application)
- [Quick start](#quick-start)
- [Call-number format](#call-number-format)
- [Printing spine labels](#printing-spine-labels)
- [Data sources and synchronization](#data-sources-and-synchronization)
- [Mobile ISBN scanner](#mobile-isbn-scanner)
- [Technology and privacy](#technology-and-privacy)
- [Troubleshooting](#troubleshooting)
- [License](#license)
- [Support](#support)

## Why this project exists

A small library needs reliable identifiers, searchable records, and legible shelf labels. Full library-management systems can demand infrastructure and maintenance that a personal collection does not need; an unstructured spreadsheet, meanwhile, leaves too much room for inconsistent classifications and duplicate identifiers.

**Personal Library Manager & Spine Studio** separates those jobs. The workbook applies a consistent classification and call-number format, while the web app makes the resulting catalog easier to browse and turn into printable labels. It is designed for a manageable workflow without requiring a database server or paid subscription.

## How it works

```text
Enter or update books in the Excel workbook
                    │
                    ▼
Generate item numbers and formatted call numbers
                    │
          ┌─────────┴─────────┐
          ▼                   ▼
   Import Excel/CSV     Publish to Google Sheets
          │                   │
          └─────────┬─────────┘
                    ▼
         Open the web application
                    │
                    ▼
      Search, filter, and select books
                    │
                    ▼
       Preview and print A4 labels
```

The public Google Sheets connection is **read-only**. Changes made in the workbook or published sheet can be imported into the app; manual entries in the app are not automatically written back to Google Sheets.

## Excel catalog and accession-number generator

**[📥 Download the Excel workbook](./Personal%20Library%20Catalogue%20%26%20Accession%20Number%20Generator%20(Enhanched).xlsx)**

The workbook is the system's primary catalog and the source of its call-number conventions. Enter the book's bibliographic and classification information; its formulas generate the associated identifiers according to the configured rules.

### What the workbook provides

- **Consistent item numbering:** three-digit, zero-padded item numbers such as `001` and `019`.
- **Author codes:** an abbreviated author-surname component used in call numbers.
- **Formatted call numbers:** a predictable combination of type, genre, author code, item number, and optional volume.
- **Controlled vocabularies:** dropdown lists help keep classification, language, and status values consistent.
- **Physical and digital catalogs:** separate records for printed books and electronic publications.

> **Catalog integrity:** Automatic formulas reduce transcription errors, but they are not a substitute for checking uniqueness and maintaining an accession register. Before sorting, deleting rows, or importing records from another collection, verify that existing identifiers remain stable. A permanent accession number should not be reassigned merely because a row moves.

### Workbook sheets

| Worksheet | Purpose |
|---|---|
| `Catalog` | Physical books, bibliographic details, classification, item numbers, call numbers, and shelf locations. |
| `Digital_Catalog` | PDFs, EPUBs, audiobooks, digital formats, OCR status, and storage links. |
| `Lookup_Codes` | Controlled values for classification and data-entry dropdowns. |

The downloadable workbook may include additional dashboard or label-layout sheets; consult the version you download for its complete worksheet list.

### Use the workbook

1. [Download the template](./Personal%20Library%20Catalogue%20%26%20Accession%20Number%20Generator%20(Enhanched).xlsx).
2. Open it in Microsoft Excel, or another spreadsheet application that supports its formulas and validation rules.
3. Enter each book's title, author, type, genre, and other available details.
4. Check the generated item number and call number; assign a shelf location.
5. Save the `.xlsx` file for local import, or upload it to Google Sheets for cloud-based reading.

**Compatibility note:** Spreadsheet formulas and data validation can behave differently in LibreOffice and Google Sheets. Verify generated identifiers after converting the workbook.

## Web application

**[🌐 Open Spine Studio](https://masudranamondal1-lgtm.github.io/library-spine-studio/)**

The browser application reads catalog data from a published Google Sheet, a local Excel/CSV file, or manual entry. It provides a searchable catalog and a dedicated label-preview and print interface.

| Feature | Description |
|---|---|
| **Search and filters** | Find books by title, author, call number, type, genre, language, status, or shelf, as supported by your catalog. |
| **Three label styles** | Choose Standard Card, Library Block, or Minimal Spine. |
| **A4 print layout** | Arrange up to 30 labels on each sheet, with optional spare positions. |
| **Google Sheets refresh** | Check the published catalog approximately every 12 seconds while connected. |
| **Local import** | Read supported `.xlsx` and `.csv` files in the browser. |
| **Manual entry** | Add books to the current local working catalog without modifying the cloud source. |
| **ISBN lookup** | On supported mobile browsers, scan an ISBN and retrieve available metadata through Google Books. |
| **Offline access** | Use cached catalog data and previously cached application resources where available. |

### Label templates

- **Standard Card:** call number, item number, title, author, volume, and optional shelf badge.
- **Library Block:** stacked call-number presentation, metadata, and an optional decorative barcode.
- **Minimal Spine:** a compact design for narrower book spines.

A decorative barcode is **not** a guaranteed scannable inventory barcode. Use a validated barcode standard and test it with a scanner before relying on it for circulation or inventory.

## Quick start

1. **Prepare your catalog.** [Download the workbook](./Personal%20Library%20Catalogue%20%26%20Accession%20Number%20Generator%20(Enhanched).xlsx), enter your books, and check the generated call numbers.
2. **Open the app.** Visit [Personal Library Manager & Spine Studio](https://masudranamondal1-lgtm.github.io/library-spine-studio/).
3. **Choose a data source.** Import your `.xlsx`/`.csv` file, connect a published Google Sheet, or enter books manually.
4. **Find your books.** Search and filter the catalog, then choose a label template and appearance settings.
5. **Preview and print.** Start with a plain-paper test sheet, check the alignment against your actual adhesive stock, and print the labels.

### Install on your device

| Platform | Installation |
|---|---|
| Android, Chrome | Open the site, then choose **Install app** or **Add to Home screen** from the browser menu, if offered. |
| iPhone/iPad, Safari | Open the site, tap **Share**, then **Add to Home Screen**. |
| Desktop, Chrome/Edge | Use the install icon in the address bar, if available. |

Installation and full offline operation depend on browser support and successful caching of the application and its external resources. A home-screen shortcut alone does not guarantee offline functionality.

## Call-number format

The workbook uses an adapted classification pattern:

```text
TYPE-GENRE-AUTHOR-ITEM-VOLUME
```

For example, `FIC-LIT-TAG-001-Vol1` identifies a fiction/literature item using the author code `TAG`, item number `001`, and volume `1`. For a single-volume work, the volume suffix may be omitted, as in `NF-HIS-HAR-019`.

| Component | Example | Meaning |
|---|---|---|
| Type | `NF` | Non-fiction |
| Genre | `HIS` | History |
| Author code | `HAR` | Code derived from Harari |
| Item number | `019` | Catalog item sequence |
| Volume | `Vol1` | Optional volume designation |

These are **project-specific call numbers**, not official Library of Congress or Dewey Decimal classifications. An abbreviated surname is also not necessarily a standard Cutter number.

## Printing spine labels

The print layout targets **A4 portrait** with **3 columns × 10 rows**, or up to **30 labels per sheet**. The intended label size is **70 × 25.4 mm**.

### Recommended browser settings

| Setting | Value |
|---|---|
| Paper | A4 |
| Orientation | Portrait |
| Scale | 100% / Actual size |
| Margins | None, if the application supplies its own print margins |
| Background graphics | On, for colored accents and guides |
| Browser headers and footers | Off |
| Fit to page | Off |

**Important physical-fit check:** Three labels measuring 70 mm across already occupy the entire 210 mm width of A4 paper. They cannot also have horizontal gaps or left/right margins. Similarly, ten 25.4 mm labels plus nine 2.2 mm gaps require 273.8 mm before top and bottom margins. The dimensions of the actual sticker stock and the application's print CSS must agree. Do not assume every product advertised as “30-up” uses the same geometry.

**Before using adhesive sheets:**

1. Print a complete test page on plain A4 paper at 100% scale.
2. Measure the first and last labels and compare the printout with your sticker stock.
3. Hold the test print behind the sticker sheet against a light source to check alignment.
4. Adjust the print layout or select matching stock if the edges drift. Do not use *Fit to page* to disguise a dimensional mismatch.

A printer's non-printable margins may prevent true edge-to-edge output even when the browser is set to zero margins.

## Data sources and synchronization

| Source | Update behavior | Persistence |
|---|---|---|
| **Published Google Sheets** | Read-only polling, approximately every 12 seconds | Browser cache after successful fetch |
| **Local Excel / CSV** | Import when selected or dropped; re-import after external changes | Browser cache, subject to browser storage limits |
| **Manual entry** | Immediate changes to the local working catalog | Session/in-memory unless explicitly saved or exported |
| **Cached catalog** | Last successfully cached records, when available | Browser storage |

### Connect Google Sheets

1. Upload the workbook to Google Drive and open it as a Google Sheet.
2. Publish the relevant catalog worksheet to the web, or configure a supported public view link.
3. Copy the published URL into the app's Google Sheets source field.
4. Connect and confirm the displayed book count and last-sync status.

**Privacy warning:** Publishing a sheet can make its contents accessible to anyone with the link, or more widely depending on the publishing settings. Do not publish borrowers' personal details, private notes, or other sensitive records. Consider a separate, public-safe catalog sheet.

Cloud synchronization is **Sheets → App only**. Local changes do not write back to Google Sheets. If cloud data is refreshed, review any unsaved local changes before replacing the working catalog.

## Mobile ISBN scanner

On supported mobile browsers, the scanner uses the browser's native `BarcodeDetector` API to read an ISBN from a book's barcode. When a code is detected, the app can request matching bibliographic metadata from the [Google Books API](https://developers.google.com/books).

Camera scanning requires browser support, camera permission, and a secure context such as HTTPS. Availability varies by browser and device; use manual ISBN entry when scanning is unavailable. Metadata returned by Google Books may be incomplete or incorrect, so review it before adding a record.

## Technology and privacy

- **Frontend:** semantic HTML5, CSS custom properties, and vanilla JavaScript (ES2020+).
- **Spreadsheet import:** [SheetJS](https://sheetjs.com/) for client-side workbook parsing.
- **Cloud catalog:** published Google Sheets CSV endpoints.
- **ISBN metadata:** [Google Books API](https://developers.google.com/books).
- **Local persistence:** browser storage for cached data and preferences.
- **Offline support:** service-worker caching where supported and successfully installed.
- **Hosting:** static files on GitHub Pages, with no application backend or build step.

The app does not require an account or its own database server. **Local spreadsheet parsing happens in the browser**, but connecting to Google Sheets, using Google Books, or loading CDN resources necessarily makes requests to those third-party services. Offline use is limited to resources already available locally or cached.

## Troubleshooting

<details>
<summary><strong>My Google Sheet will not load.</strong></summary>

Check that the relevant worksheet is published or accessible through a supported public link. Confirm the URL and worksheet name, and test the published CSV endpoint directly. Browser/network restrictions can also prevent access.

</details>

<details>
<summary><strong>My labels are misaligned.</strong></summary>

Use A4 paper, 100% scale, no browser headers/footers, and a plain-paper test. Verify the actual label width, height, gaps, and sheet margins. Different 30-up products may have incompatible dimensions.

</details>

<details>
<summary><strong>Bengali text appears as boxes.</strong></summary>

Use UTF-8 CSV files and ensure the device has a Bengali-capable font, such as [Noto Sans Bengali](https://fonts.google.com/noto/specimen/Noto+Sans+Bengali). Check the browser's print preview before printing.

</details>

<details>
<summary><strong>The app still shows an older version.</strong></summary>

Reload the page. If the app uses a service worker, an older cached version may remain active until the updated worker installs. Developers should update the service-worker cache version when deploying changes; users may need to close and reopen the app or clear its site data.

</details>

<details>
<summary><strong>My call numbers are blank or duplicated.</strong></summary>

Check the required workbook fields and formulas, particularly the title, classification values, author, and item-number sequence. Verify that formula cells were not overwritten and that existing identifiers have not been regenerated after sorting or deleting records.

</details>

<details>
<summary><strong>ISBN scanning is unavailable.</strong></summary>

Check camera permissions, HTTPS, and browser support for `BarcodeDetector`. Use manual ISBN entry if the device does not support camera detection.

</details>

## License

Released under the **MIT License**. See [LICENSE](./LICENSE) for the complete terms.

## Support

If this project helps you organize your books, consider sharing it with another reader or supporting its development through the **📖 Buy me a book** button in the app or [Buy Me a Coffee](https://buymeacoffee.com/masudshaon).

Found a bug or have an idea? [Open a GitHub issue](https://github.com/masudranamondal1-lgtm/library-spine-studio/issues).
