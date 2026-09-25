# 📚 Personal Library Manager & Spine Label Studio

A lightweight, client-side personal library cataloging system and print-calibrated book spine label studio. Catalog your books, filter them, and print A4 30-up spine labels — all in the browser, with optional Google Sheets sync.

🌐 **Live Demo:** [https://masudranamondal1-lgtm.github.io/library-spine-studio/](https://masudranamondal1-lgtm.github.io/library-spine-studio/)
📊 **Excel Template:** [Personal Library Catalogue & Accession Number Generator (Enhanced).xlsx](./Personal%20Library%20Catalogue%20%26%20Accession%20Number%20Generator%20(Enhanced).xlsx)

---

## ✨ Features

- **Three data sources** — Google Sheets (live-synced), local Excel/CSV import, or manual entry
- **Three label templates** — Standard Card, Library Block, and Minimal Spine
- **Calibrated A4 30-up grid** — 70 mm × 25.4 mm labels, 3 columns × 10 rows, ready for standard sticker sheets
- **Automatic call numbers** — `TYPE-GENRE-AUTHOR-ID-VOLUME` pattern (e.g. `FIC-LIT-TAG-001-Vol1`)
- **Search, filter, and sort** — by title, author, call number, type, genre, language, status, or shelf
- **Light / Dark / Auto theme** — with `prefers-color-scheme` support and persistence
- **100% client-side** — no backend, no tracking, no accounts; your data never leaves your browser
- **Works offline** — cached catalog stays available when the network is down

---

## 🚀 Quick Start

### Use the live app

1. Open [https://masudranamondal1-lgtm.github.io/library-spine-studio/](https://masudranamondal1-lgtm.github.io/library-spine-studio/)
2. Choose a data source:
   - **Google Sheets** — paste a published or standard Google Sheets URL
   - **Local File** — drag in an `.xlsx`, `.xls`, or `.csv` file
   - **Manual** — add books one at a time
3. Pick a label template and color mode
4. Click **Print** and use the browser print settings below

### Run locally

Download `index.html` and open it in any modern browser. Everything works offline — Excel parsing, filtering, label rendering, and printing.

### Deploy to GitHub Pages

Fork the repo → **Settings → Pages** → set Source to `main` / `root`. Your app is live at `https://<username>.github.io/<repo>/`.

---

## 📊 Excel Workbook

The workbook has three worksheets:

| Worksheet | Purpose |
|-----------|---------|
| **Catalog** | Physical books. Auto-generates item numbers and call numbers. |
| **Digital_Catalog** | PDFs, EPUBs, audiobooks — tracks format, OCR status, storage links. |
| **Lookup_Codes** | Controlled vocabularies for Type, Genre, Language, Status, Format. |

**Call number pattern:** `TYPE-GENRE-AUTHOR-ID-VOLUME` — for example `NF-HIS-HAR-019`.

---

## 🖨 Print Settings

For accurate A4 30-up sticker alignment:

```
Paper size:            A4
Layout:                Portrait
Scale:                 100%
Margins:               None
Background graphics:   ON
Headers and footers:   OFF
Fit to page:           OFF   ⚠️
```

> Do not use "Fit to page" — it silently shrinks the layout and misaligns labels.

Each sheet holds exactly 30 labels at **70 mm × 25.4 mm**, with the following margins: 10 mm top, 8 mm bottom, 9 mm left, 9 mm right.

---

## 🔄 Data Sync

| Source | Sync | Persistence |
|--------|------|-------------|
| Google Sheets (published or `/edit`) | Polls every 12 s | `localStorage` cache |
| Local `.xlsx` / `.xls` / `.csv` | One-time import | `localStorage` cache |
| Manual entry | Session-only | In-memory |
| Cached data | Offline fallback | — |

Google Sheets sync is **one-directional** (Sheets → App). The app never writes back to your spreadsheet; manual additions stay local.

---

## 🛠 Tech Stack

- Semantic HTML5, vanilla CSS with custom-property tokens, vanilla JavaScript (ES2020+)
- [SheetJS](https://sheetjs.com/) for client-side Excel parsing
- `localStorage` for cache, preferences, and theme
- No build step, no framework, no dependencies beyond the single CDN script

---

## 🐛 Troubleshooting

**Google Sheet won't load.** Make sure the sheet is published to the web or shared as "Anyone with the link can view".

**Labels are misaligned.** Print on plain paper first, verify all print settings above, and confirm the paper size is genuinely A4 (not Letter).

**Bengali text shows as boxes.** Install Noto Sans Bengali locally, or save CSV files as UTF-8.

**Manual additions disappeared.** Manual entries live only in the current browser's `localStorage` — add them to the workbook and re-sync to persist across devices.

---

## 🙏 Support

If this project saved you time, you can buy the developer a book via the **📖 Buy me a book** button inside the app.

---

## 📄 License

MIT License. See [LICENSE](./LICENSE) for details.
