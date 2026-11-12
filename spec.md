# Spec: Wint Master/Taxation Report — Word (.docx) Export Feature

**Branch:** `feature/excel-to-pdf-converter`  
**Status:** Planning — do not implement until confirmed  
**Prepared:** 2026-11-12  

---

## 1. Overview

Add a **"Download Word"** option alongside the existing Excel and PDF download buttons. The `.docx` output must be a faithful, fully-formatted replica of the accepted reference document, including:

- Wint letterhead on **every page**
- Exact table structure for each report section
- Correctly rendered metadata, acquisition-cost note, and legal disclaimer
- Signed-off closing block ("For and on behalf of Wint Securities…")

---

## 2. Letterhead Reference

![Letterhead Screenshot](./public/wint-logo.png)

> **Reference:** Letterhead uses Wint logo (left) + company details right-aligned (Poppins font) + a full-width green/black decorative stripe below.

### 2a. Letterhead Anatomy (from `header1.xml`)

The header is a **borderless 2-column table** that spans the full page width (`21962 dxa` ≈ 15.3 cm × full A4):

| Column | Width (dxa) | Content |
|--------|-------------|---------|
| Left | 4228 | Wint logo (`image3.png` → extracted as `wint-logo.png`, 276×178 px) |
| Right | 17734 | Company text, **right-aligned**, font: **Poppins** |

**Right-column text (3 lines, all right-aligned):**
1. `WINT SECURITIES PRIVATE LIMITED (Formerly known as Fourdegreewater Services Private Limited)` — **Bold**, Poppins, ~9pt
2. `CIN: U66120KA2022PTC165599` — Normal, Poppins, ~9pt
3. `Website: https://www.wintwealth.com` — Normal, Poppins, ~9pt

**Below the header table:** A full-width thin decorative bar image (`image2.png` → `wint-header-bar.png`, 516×3 px — a green/black gradient stripe) that runs edge-to-edge.

This two-element header appears on **every page** of the document.

---

## 3. Document Sections (from `document.xml`)

The reference `.docx` contains **four distinct report sections**, each identified by a bold centred heading row:

| # | Section Title | Table Columns | Notes |
|---|--------------|---------------|-------|
| 1 | CONSOLIDATED INVESTMENT REPORT WINT WEALTH | 11 | Per-FY investment summaries; repeats for each FY |
| 2 | CONSOLIDATED REPAYMENT REPORT WINT WEALTH | 9 | Past cashflow per FY |
| 3 | Upcoming Cashflow Statement | 5 | Month-wise forward payments |
| 4 | Bonds Holding Statement (Wint Wealth) | 17 | Current portfolio snapshot |

Each section ends with:
- A **"Date of report generation: DD/MM/YYYY"** row
- An **\*Acquisition Cost** note paragraph (plain text, 10pt)
- A **Disclaimer** paragraph (plain text, 10pt, italic or regular)

The document closes with a signing block:
> *For and on behalf of Wint Securities Private Limited (formerly known as Fourdegreewater Services Private Limited)*
> Name: Rohit Kumar Jha
> Designation: Company Secretary and Compliance Officer

---

## 4. Typography & Styling (from `document.xml` runs)

| Element | Font | Size | Style | Alignment |
|---------|------|------|-------|-----------|
| Section heading row | Helvetica | 10pt (20 half-pt) | Bold | Centre |
| Column headers | Helvetica-Bold / Helvetica | 10pt | Bold | Centre or Left |
| Data cells | Helvetica | 10pt | Normal | Left / Right for numbers |
| Total row | Helvetica | 10pt | Normal | Left |
| Acquisition Cost label | Helvetica-Bold | 10pt | Bold | Left |
| Acquisition Cost text | Helvetica | 10pt | Normal | Justified |
| Disclaimer label | Helvetica-Bold | 10pt | Bold | Left |
| Disclaimer text | Helvetica | 10pt | Normal | Justified |
| Signing block | Helvetica | 10pt | Normal | Left |
| Header company name | Poppins | ~9pt | Bold | Right |
| Header CIN / Website | Poppins | ~9pt | Normal | Right |

**Table borders:** `single`, sz=4, color=`auto` (black) on all sides + inner grid lines.  
**Row height:** `399 dxa` for data rows, `1693 dxa` for header logo row.

---

## 5. Library Choice

| Library | Why |
|---------|-----|
| **`docx`** (npm: `docx`) | Pure TypeScript/Node OOXML generator; no native deps; Vercel-compatible; supports headers with images, tables, runs, paragraphs, fonts, borders | ✅ **Recommended** |
| `officegen` | Older, less maintained, no TypeScript types | ❌ |
| `docxtemplater` | Template-based only; can't generate from scratch | ❌ |

---

## 6. Implementation Plan

### Phase 1 — New dependency
```
npm install docx
```

### Phase 2 — New conversion library (`lib/excel-to-docx.ts`)

Exports:
```typescript
export function convertExcelToDocx(
  data: ArrayBuffer | Uint8Array,
  options?: ExcelToDocxOptions
): Promise<Uint8Array>
```

Internally uses the existing `parseSheetSections()` logic from `lib/excel-to-pdf.ts` (to be extracted into a shared `lib/sheet-parser.ts`).

**Steps per sheet:**
1. Parse metadata, section title, table headers, body rows, disclaimers (reuse parser)
2. Build `docx.Document` with:
   - `sections[].headers` containing the letterhead (`Header` object)
   - Per-section: heading `Paragraph`, metadata key-value `Table` (2-col), data `Table`, disclaimer `Paragraph`s
3. Pack and return as `Uint8Array` via `docx.Packer.toBuffer()`

**Letterhead construction (every page):**
```
Header → Table (2 cols, no border)
  Col A (narrow): ImageRun(wint-logo.png, 50mm × 26mm)
  Col B (wide):   3× Paragraphs right-aligned, Poppins
                   - "WINT SECURITIES PRIVATE LIMITED..." (Bold)
                   - "CIN: U66120KA2022PTC165599"
                   - "Website: https://www.wintwealth.com"
After table: ImageRun(wint-header-bar.png, full-width, 2mm tall)
```

### Phase 3 — New API route (`app/api/convert-docx/route.ts`)

```
GET /api/convert-docx?url=<encodedUrl>&filename=<name>
```

Same pattern as `/api/convert-pdf`:
1. Validate URL via `parseDownloadUrl()`
2. Fetch upstream `.xlsx`
3. Call `convertExcelToDocx(buffer)`
4. Stream back with `Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document` and `Content-Disposition: attachment; filename="..."`

### Phase 4 — UI changes

#### `lib/api.ts` — Add helpers
```typescript
export function docxDownloadHref(url: string, filename: string): string
export function triggerDocxDownload(url: string, filename: string): void
```

#### `components/ResultPanel.tsx` — Add Word button
When a report is ready, show **three** download buttons:
```
[ Excel (.xlsx) ]   [ PDF (.pdf) ]   [ Word (.docx) ]
```

#### `components/HistoryList.tsx` — Add Word re-download link
Each successful history entry gets a third "Word" link: `Excel · PDF · Word`

#### `components/ExcelToPdfConverter.tsx` — Add Word export option
Add a "Download as Word" button alongside "Convert & Download PDF" in the standalone converter.

### Phase 5 — Assets
Two static images committed to `public/`:
- `public/wint-logo.png` — Wint logo (276×178 px, extracted from reference `.docx`)
- `public/wint-header-bar.png` — Decorative colour bar (516×3 px, extracted from reference `.docx`)

Both are already in `public/` from the extraction step above.

---

## 7. File Change Summary

| File | Action | Description |
|------|--------|-------------|
| `package.json` | Modify | Add `docx` dependency |
| `lib/sheet-parser.ts` | **New** | Shared parsing logic extracted from `excel-to-pdf.ts` |
| `lib/excel-to-docx.ts` | **New** | Core `.docx` generation with letterhead + tables |
| `lib/excel-to-pdf.ts` | Modify | Import `parseSheetSections` from new `sheet-parser.ts` |
| `lib/api.ts` | Modify | Add `docxDownloadHref()` + `triggerDocxDownload()` |
| `app/api/convert-docx/route.ts` | **New** | API route: fetch → convert → stream `.docx` |
| `components/ResultPanel.tsx` | Modify | Add "Word (.docx)" button |
| `components/HistoryList.tsx` | Modify | Add "Word" re-download link |
| `components/ExcelToPdfConverter.tsx` | Modify | Add Word export button |
| `public/wint-logo.png` | **New** | Wint logo (extracted from reference `.docx`) |
| `public/wint-header-bar.png` | **New** | Decorative header bar (extracted from reference `.docx`) |

---

## 8. Open Questions

> [!IMPORTANT]
> Please confirm the following before implementation begins:

1. **Signing block** — Should the closing block ("For and on behalf of Wint Securities…", Name: Rohit Kumar Jha, Designation: Company Secretary…) be included in every generated `.docx`? Or is it only for some report types?

2. **Report scope** — Should the Word export cover **all four sections** (Investment Summary + Repayment + Upcoming Cashflow + Holding Statement) in a single `.docx` file, or only the **Master / Taxation report** that the user selected on the form?

3. **Logo source** — The logo was extracted from the reference `.docx` (small, 276×178 px). Should we use the higher-resolution `wint-logo-hd.png` already in the repo root instead? (It is significantly larger but the same graphic.)

4. **Fonts** — The reference document uses **Poppins** (header) and **Helvetica** (body). `docx` library embeds font *names* but not font *files*, so the user's machine must have these fonts installed to render them correctly. Is that acceptable, or should we fall back to `Arial` as a universal substitute?

5. **Standalone converter** — Should the Word export button also appear in the **Excel-to-PDF Converter** tool (for uploading arbitrary `.xlsx` files), or only on the **User Report Downloader** where the upstream report is fetched?
