# 📚 Personal Library Manager & Spine Label Studio

A lightweight, production-ready personal library cataloging system and print-calibrated book spine label studio.

Built to run completely client-side in any web browser, operate offline, or live-sync catalog updates directly with Google Sheets and Microsoft Excel.

---

## 📑 Table of Contents

* [Overview](#-overview)
* [Key Features](#-key-features)
* [Architecture & Tech Stack](#-architecture--tech-stack)
* [Cataloging & Call Number Syntax](#-cataloging--call-number-syntax)
* [Repository Structure](#-repository-structure)
* [Quick Start Guide](#-quick-start-guide)

  * [1. Running Locally](#1-running-locally)
  * [2. Deploying to GitHub Pages](#2-deploying-to-github-pages)
  * [3. Installing as a PWA / Mobile App](#3-installing-as-a-pwa--mobile-app)
* [Data Sources & Synchronization](#-data-sources--synchronization)

  * [Option A: Live Google Sheets Auto-Sync](#option-a-live-google-sheets-auto-sync)
  * [Option B: Offline Excel Drag-and-Drop](#option-b-offline-excel-drag-and-drop)
  * [Option C: Manual / Quick Entry](#option-c-manual--quick-entry)
* [Printing Calibration & Specifications](#-printing-calibration--specifications)
* [Excel Workbook Schema](#-excel-workbook-schema)
* [Troubleshooting & FAQ](#-troubleshooting--faq)
* [License](#-license)

---

## 📖 Overview

Organizing a collection of physical books and digital publications typically forces a choice between bloated, subscription-heavy library enterprise software or unformatted, static spreadsheets.

**Personal Library Manager & Spine Label Studio** bridges that gap.

It combines a structured, multi-tier classification spreadsheet with automated Cutter author codes and sequence generation, together with an interactive HTML5/JavaScript label studio calibrated to standard A4 adhesive sticker sheets.

Any modification made in your catalog can be reflected immediately on printable spine labels.

---

## 🌟 Key Features

* **Multi-Tier Classification Engine:** Automatically formats and validates call numbers across five hierarchical metadata tiers: Type, Genre, Author Surname, Sequential ID, and Volume Designation.

* **Live Background Polling:** Monitors your connected Google Sheet via Google Visualization (GViz) CSV endpoints every 10 seconds, refreshing the preview grid whenever items are added, updated, or removed.

* **Client-Side Excel Engine:** Parses local `.xlsx` files directly in memory via SheetJS with zero external server dependencies, ensuring complete data privacy.

* **Calibrated A4 Sticker Layout (30-Up):** Built to exact dimensions (`70 mm × 25.4 mm` across 3 columns × 10 rows), preventing page creep and vertical drift across multiple pages.

* **Three Interchangeable Label Formats:**

  1. **Standard Detailed Card:** Accession ID, full call number, title, volume, author, and shelf badge.
  2. **Library Call-Block + Barcode:** Split monospace classification tag with simulated barcode tracking.
  3. **Minimal Spine Strip:** Compact horizontal layout for thin paperbacks and periodicals.

* **Visual Categorization:** Dynamically color-codes spine edges based on classification types such as Fiction, Non-Fiction, Poetry, Drama, Reference, and Graphic Novels, or switches to clean monochrome.

* **Automated Spare Filling:** Pads incomplete sheets with blank cutting frames so partially populated runs do not waste adhesive sheets.

---

## 🛠️ Architecture & Tech Stack

```text
┌────────────────────────────────────────────────────────┐
│                    DATA LAYER                          │
│  • Google Sheets (Cloud REST / GViz CSV API)           │
│  • Excel (.xlsx) parsed via SheetJS (XLSX CDN)         │
│  • SQLite3 / CSV local fallback storage                │
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
┌────────────────────────────────────────────────────────┐
│              CORE ENGINE (index.html)                  │
│  • Dynamic Cutter extraction algorithm                 │
│  • State synchronization & interval auto-polling       │
│  • A4 pagination & print stylesheet engine             │
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
┌────────────────────────────────────────────────────────┐
│                   OUTPUT LAYER                         │
│  • Desktop / Mobile UI Viewport                        │
│  • Native Print / PDF Engine (A4 30-Up Sticker Grid)   │
└────────────────────────────────────────────────────────┘
```

### Frontend

* Vanilla HTML5
* Modern CSS Grid / Flexbox
* ES6+ JavaScript
* Zero framework bloat

### Excel Parsing

[SheetJS (`xlsx.full.min.js`)](https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js) loaded via CDN.

### Spreadsheet Engine

Pre-built `.xlsx` template containing:

* Cross-tab data validation
* Dynamic accession formulas
* Lending tracking
* Analytics dashboard

---

## 🏷️ Cataloging & Call Number Syntax

Call numbers follow an adapted library classification standard:

```text
[Type]-[Genre]-[Author]-[Item ID]-[Volume]
```

Example:

```text
FIC - LIT - TAG - 001 - Vol1
 │     │     │     │     │
 │     │     │     │     └── Volume designation
 │     │     │     │         (omitted for single-volume books)
 │     │     │     └──────── 3-digit zero-padded sequential accession ID
 │     │     └────────────── 3-letter Cutter code derived from author's surname
 │     └──────────────────── 3-letter genre classification code
 └────────────────────────── 2–3 letter format / type code
```

### Classification Codes Reference

| Type Code | Meaning                | Genre Code | Meaning              |
| --------- | ---------------------- | ---------- | -------------------- |
| `FIC`     | Fiction                | `LIT`      | Literature           |
| `NF`      | Non-Fiction            | `HIS`      | History              |
| `POE`     | Poetry                 | `PHI`      | Philosophy           |
| `DRA`     | Drama                  | `SCI`      | Science & Technology |
| `REF`     | Reference              | `POL`      | Politics             |
| `GRA`     | Graphic Novel          | `BIO`      | Biography / Memoir   |
| `DIG-`    | Digital Catalog Prefix | `ESS`      | Essays               |
| —         | —                      | `THR`      | Mystery / Thriller   |
| —         | —                      | `REL`      | Religion & Mythology |
| —         | —                      | `ART`      | Art & Cinema         |

### Example Transformation

**Book:** *Sapiens: A Brief History of Humankind* by Yuval Noah Harari

* **Type:** Non-Fiction (`NF`)
* **Genre:** History (`HIS`)
* **Author Surname:** Harari → `HAR`
* **Item Number:** `019`
* **Generated Call Number:** `NF-HIS-HAR-019`

---

## 📁 Repository Structure

```text
├── index.html
│   # Main Application & Spine Studio
│   # formerly connected_spine_label_generator.html
│
├── Personal_Library_Catalog_Enhanced.xlsx
│   # Pre-configured Excel Catalog with Dashboard & Spine tabs
│
├── LibraryManager_App.py
│   # Standalone Python/Tkinter Desktop GUI application
│
├── Run_Library_App.bat
│   # 1-Click Windows batch launcher for standalone app
│
├── Library_Catalog.csv
│   # Seeded physical library collection schema
│
├── Library_Digital_Catalog.csv
│   # E-book and PDF digital library schema
│
├── Library_Dashboard.csv
│   # Analytics layout data backup
│
├── Library_Spine_Labels.csv
│   # Raw label export backup
│
├── Library_Lookup_Codes.csv
│   # Controlled master dropdown tables
│
└── README.md
    # Comprehensive documentation
```

---

## 🚀 Quick Start Guide

### 1. Running Locally

Clone the repository or extract the project files:

```bash
git clone https://github.com/YOUR-USERNAME/library-spine-studio.git
cd library-spine-studio
```

Then open `index.html` in:

* Google Chrome
* Microsoft Edge
* Firefox
* Safari

No web server or build step (`npm` / `node`) is required.

---

### 2. Deploying to GitHub Pages

To make the application globally accessible on mobile and desktop:

#### Step 1: Push the repository to GitHub

```bash
git add .
git commit -m "Initial commit of Library Studio"
git push origin main
```

#### Step 2: Enable GitHub Pages

In your GitHub repository:

1. Navigate to **Settings → Pages**.
2. Under **Build and deployment → Branch**, select `main` (or `master`).
3. Select the directory `/ (root)`.
4. Click **Save**.

Within a short period, the application will be available at:

```text
https://YOUR-USERNAME.github.io/library-spine-studio/
```

---

### 3. Installing as a PWA / Mobile App

#### Android

Open the live deployment URL in Chrome:

1. Tap the top-right menu (`⋮`).
2. Select **Add to Home screen** or **Install App**.

#### iOS

Open the URL in Safari:

1. Tap the **Share** button.
2. Select **Add to Home Screen**.

#### PC / Mac

In Chrome or Edge:

1. Open the live application.
2. Click the **Install** icon on the right side of the address bar.
3. The application will open in a dedicated desktop window.

---

## 🔄 Data Sources & Synchronization

The application supports three primary data-entry and synchronization methods.

### Option A: Live Google Sheets Auto-Sync

1. Upload `Personal_Library_Catalog_Enhanced.xlsx` to Google Drive.
2. Open the file as a Google Sheet.
3. Set the spreadsheet access to **"Anyone with the link can view"**, or publish it via:
   **File → Share → Publish to web**.
4. Copy your spreadsheet URL or Sheet ID from the address bar.

Example:

```text
https://docs.google.com/spreadsheets/d/YOUR-SHEET-ID/edit
```

5. Paste the URL into the **☁️ Live Cloud Sync** field inside the application.
6. Click **Fetch Live Books**.
7. Keep **Live Auto-Refresh** enabled.

The studio will poll the Google Sheet every 10 seconds and automatically re-render the printable labels whenever books are added, modified, or removed.

---

### Option B: Offline Excel Drag-and-Drop

1. Maintain your collection inside `Personal_Library_Catalog_Enhanced.xlsx`.
2. Open the application.
3. Switch to the **📁 Upload Excel** tab.
4. Drag and drop your `.xlsx` file onto the upload area.
5. SheetJS parses the `Catalog` worksheet locally.
6. The application generates the label grids instantly.

No spreadsheet data needs to be uploaded to an external server.

---

### Option C: Manual / Quick Entry

1. Switch to the **✏️ Manual Entry** tab.
2. Select the appropriate **Type** and **Genre**.
3. Enter the **Title** and **Author**.
4. Specify a target shelf.
5. Click **+ Add Book**.

The system will:

* Calculate the next sequential ID.
* Generate the Cutter author code.
* Format the call number.
* Push the new label into the print queue.

---

## 🖨️ Printing Calibration & Specifications

The printable layout is engineered for standard **30-Up A4 Adhesive Sticker Sheets**, comparable to Avery 7160 / 5160 metric variants.

### Physical Grid Dimensions

| Specification          | Value                                         |
| ---------------------- | --------------------------------------------- |
| **Paper Format**       | ISO A4 (`210 mm × 297 mm`), Portrait          |
| **Grid Layout**        | 3 columns × 10 rows                           |
| **Labels per Sheet**   | 30                                            |
| **Sticker Dimensions** | `70.0 mm × 25.4 mm`                           |
| **Column Gap**         | `3.5 mm`                                      |
| **Row Gap**            | `2.2 mm`                                      |
| **Page Margins**       | Top `10 mm`, Bottom `8 mm`, Left/Right `9 mm` |

### Sheet Layout

```text
               ◄────── 210mm (A4 Width) ──────►
         ┌──────────────────────────────────────────┐ ▲
         │        SHEET BANNER / PAGE INDEX         │ │ 10mm Top Margin
         ├────────────┬─────────────┬───────────────┤ ▼
         │ Label 1    │ Label 2     │ Label 3       │ ▲
         │ 70 × 25mm  │ 70 × 25mm   │ 70 × 25mm     │ │
         ├────────────┼─────────────┼───────────────┤ │ 297mm
         │ Label 4    │ Label 5     │ Label 6       │ │ (A4 Height)
         │            │             │               │ │
         │   ...      │    ...      │    ...        │ │
         ├────────────┼─────────────┼───────────────┤ │
         │ Label 28   │ Label 29    │ Label 30      │ ▼
         └────────────┴─────────────┴───────────────┘ ▲ 8mm Bottom Margin
```

### Exact Browser Print Settings

1. Press `Ctrl + P` on Windows/Linux or `Cmd + P` on macOS.
2. Set **Destination** to your physical printer or **Save as PDF**.
3. Set **Paper Size** to **A4**.
4. Set **Margins** to **None** or **Custom: 0 mm**.
5. **Enable "Background graphics".**

   * This renders the colored category bars.
   * This renders cut guides.
   * This renders shelf badges.
6. Set **Scale** to **100% / Default**.
7. Do **not** use **Fit to page**, as this can distort sticker alignment.

---

## 📊 Excel Workbook Schema

The included spreadsheet, `Personal_Library_Catalog_Enhanced.xlsx`, contains five specialized worksheets:

| Worksheet             | Purpose                                                                                                                                                                                         |
| --------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **`Dashboard`**       | High-level collection analytics including total physical volumes, digital titles, combined counts, reading progress breakdown (`Read`, `Reading`, `Unread`, `Loan`), and language distribution. |
| **`Catalog`**         | Master inventory for physical books. Contains auto-numbering sequence formulas, unified Cutter call number calculations, shelf allocation, borrower tracking, and loan dates.                   |
| **`Digital_Catalog`** | Inventory for PDFs, EPUBs, and audiobooks. Tracks OCR searchable status, digital format tags, and cloud storage links.                                                                          |
| **`Spine_Labels`**    | In-sheet formula-driven print layout referencing physical catalog items.                                                                                                                        |
| **`Lookup_Codes`**    | Centralized master tables feeding data-validation dropdowns across all catalog tabs.                                                                                                            |

---

## ❓ Troubleshooting & FAQ

### Why are labels misaligned with my sticker sheet?

Verify that your browser's print dialog has:

* **Margins:** `None`
* **Scale:** `100%`
* **Paper Size:** `A4`

Enabling printer margins can shift the CSS grid downwards and cause alignment drift on lower rows.

---

### Why does the Google Sheets sync status indicate an error?

Ensure that your Google Sheet sharing permissions are set to:

> **Anyone with the link can view**

If your sheet is restricted to private access, Google's public CSV API may block the request.

---

### How do I add custom genres or categories?

Open:

```text
Personal_Library_Catalog_Enhanced.xlsx
```

Navigate to the:

```text
Lookup_Codes
```

worksheet and add your entries under the appropriate column.

Then update the corresponding dropdown menus in `index.html`, particularly:

```html
#qType
#qGenre
```

---

## 📄 License

This project is licensed under the **MIT License**.

You are free to:

* Use the software
* Modify the software
* Distribute the software
* Commercialize the software

without restriction, subject to the terms of the MIT License.
