import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import autoTable, { UserOptions } from "jspdf-autotable";

export interface ExcelToPdfOptions {
  orientation?: "portrait" | "landscape" | "auto";
  title?: string;
  subtitle?: string;
  selectedSheets?: string[]; // If omitted or empty, all sheets are included
  fontSize?: number;
  themeColor?: string; // e.g. '#1c8a4b'
}

export interface SheetPreviewData {
  name: string;
  totalRows: number;
  totalCols: number;
  headers: string[];
  previewRows: (string | number | boolean | null)[][];
}

export interface WorkbookInfo {
  sheetNames: string[];
  sheets: SheetPreviewData[];
}

interface ParsedSheetContent {
  bannerTitle?: string;
  sectionTitle?: string;
  metadata: Array<{ label: string; value: string }>;
  headers: string[];
  bodyRows: string[][];
  disclaimers: string[];
  columnAlignments: Array<"left" | "center" | "right">;
}

/**
 * Checks if a string looks like a disclaimer, legal note, or terms block.
 */
function isDisclaimerOrNote(text: string): boolean {
  const trimmed = text.trim();
  const lower = trimmed.toLowerCase();
  if (lower.startsWith("disclaimer") || lower.startsWith("*acquisition cost") || lower.startsWith("acquisition cost")) {
    return true;
  }
  if (lower.startsWith("note:") || lower.startsWith("notes:") || lower.startsWith("important note")) {
    return true;
  }
  if (
    lower.includes("the document and details set out thereunder") ||
    lower.includes("before filing your income tax return") ||
    lower.includes("we have classified") ||
    lower.includes("all information provided herein") ||
    lower.includes("tax laws and interpretations") ||
    lower.includes("wint wealth disclaims") ||
    lower.includes("please consult a qualified tax") ||
    lower.includes("demat linked bank account") ||
    lower.includes("informational purposes only") ||
    lower.includes("cannot guarantee 100% accuracy")
  ) {
    return true;
  }
  // Any single cell with very long text (> 110 characters) is almost certainly a paragraph/note
  if (trimmed.length > 110 && !trimmed.includes("\t")) {
    return true;
  }
  return false;
}

/**
 * Checks if a string looks like a user/report metadata key (e.g. Name, Phone, Email, Date).
 */
function isMetadataKey(text: string): boolean {
  const lower = text.trim().toLowerCase().replace(/:$/, "");
  const keys = [
    "name",
    "phone number",
    "phone",
    "email",
    "user id",
    "user",
    "pan",
    "pan number",
    "date of report generation",
    "report generation date",
    "report date",
    "date of generation",
    "generated on",
    "client id",
    "demat account",
  ];
  return keys.includes(lower);
}

/**
 * Detects whether a column primarily contains numbers, currency, dates, or text.
 */
function detectColumnAlignment(header: string, values: string[]): "left" | "center" | "right" {
  const hLower = header.toLowerCase();

  // Explicit header cues
  if (
    hLower.includes("date") ||
    hLower === "isin" ||
    hLower.includes("pan") ||
    hLower === "fy" ||
    hLower.includes("year") ||
    hLower.includes("month") ||
    hLower.includes("frequency") ||
    hLower.includes("status") ||
    hLower.includes("type")
  ) {
    return "center";
  }

  if (
    hLower.includes("amount") ||
    hLower.includes("units") ||
    hLower.includes("value") ||
    hLower.includes("cost") ||
    hLower.includes("xirr") ||
    hLower.includes("ytm") ||
    hLower.includes("tds") ||
    hLower.includes("interest") ||
    hLower.includes("principal") ||
    hLower.includes("repaid") ||
    hLower.includes("sold") ||
    hLower.includes("total") ||
    hLower.includes("rate") ||
    hLower.includes("price")
  ) {
    return "right";
  }

  // Value inspection
  let numericCount = 0;
  let dateCount = 0;
  let nonEmptyCount = 0;

  for (const v of values) {
    const val = v.trim();
    if (!val || val === "-") continue;
    nonEmptyCount++;

    // Date check (e.g. 17/10/2025, 2025-04-15, 13-04-2027)
    if (/^\d{1,4}[-/.]\d{1,2}[-/.]\d{2,4}$/.test(val) || /^\d{4}-\d{4}$/.test(val)) {
      dateCount++;
      continue;
    }

    // Number / Currency / Percent check
    const cleanNum = val.replace(/^[₹$€£]\s?/, "").replace(/%$/, "").replace(/,/g, "").trim();
    if (!isNaN(Number(cleanNum))) {
      numericCount++;
    }
  }

  if (nonEmptyCount > 0) {
    if (numericCount / nonEmptyCount > 0.6) return "right";
    if (dateCount / nonEmptyCount > 0.6) return "center";
  }

  return "left";
}

/**
 * Intelligently separates an Excel worksheet into structured sections:
 * - Banner titles (e.g. "CONSOLIDATED INVESTMENT REPORT WINT WEALTH")
 * - User info / metadata (e.g. Name: Vedant, Phone: 8637663482, Date: 10/11/2026)
 * - Section titles (e.g. "INVESTMENT SUMMARY FY 25-26")
 * - The actual data table (headers + body rows)
 * - Disclaimers & notes (separated from the table to prevent column squishing & page bloat)
 */
function parseSheetSections(rawData: unknown[][]): ParsedSheetContent | null {
  const cleanRows = rawData
    .map((row) =>
      (Array.isArray(row) ? row : []).map((cell) =>
        cell !== null && cell !== undefined ? String(cell).trim() : ""
      )
    )
    .filter((row) => row.some((val) => val !== ""));

  if (cleanRows.length === 0) return null;

  const metadata: Array<{ label: string; value: string }> = [];
  const disclaimers: string[] = [];
  let bannerTitle: string | undefined = undefined;
  let sectionTitle: string | undefined = undefined;

  const candidateTableRows: string[][] = [];

  for (let i = 0; i < cleanRows.length; i++) {
    const row = cleanRows[i];
    const nonBlank = row.filter((c) => c !== "");
    const firstCell = nonBlank[0] || "";
    const firstCellLower = firstCell.toLowerCase();

    // 1. Disclaimer / Note detection
    if (isDisclaimerOrNote(firstCell) || (nonBlank.length <= 2 && row.join(" ").length > 110)) {
      const fullText = nonBlank.join(" ").trim();
      if (fullText && !disclaimers.includes(fullText)) {
        disclaimers.push(fullText);
      }
      continue;
    }

    // 2. Metadata detection (Key-value pairs like Name, Phone, Email, Date of report generation)
    if (nonBlank.length >= 2 && nonBlank.length <= 4 && isMetadataKey(nonBlank[0])) {
      const label = nonBlank[0].replace(/:$/, "").trim();
      const value = nonBlank[1]?.trim() || "";
      if (label && value && !metadata.some((m) => m.label.toLowerCase() === label.toLowerCase())) {
        metadata.push({ label, value });
      }
      continue;
    }

    // 3. Document / Banner title (Single prominent cell before table)
    if (
      candidateTableRows.length === 0 &&
      nonBlank.length === 1 &&
      (firstCellLower.includes("report") || firstCellLower.includes("wint") || firstCellLower.includes("consolidated"))
    ) {
      bannerTitle = firstCell;
      continue;
    }

    // 4. Section subheadings (e.g. "INVESTMENT SUMMARY (...) FY 25-26", "Bonds Holding Statement")
    if (
      nonBlank.length === 1 &&
      (firstCellLower.includes("summary") || firstCellLower.includes("statement") || firstCellLower.includes("holding")) &&
      candidateTableRows.length <= 1
    ) {
      sectionTitle = firstCell;
      continue;
    }

    candidateTableRows.push(row);
  }

  if (candidateTableRows.length === 0) {
    return {
      bannerTitle,
      sectionTitle,
      metadata,
      headers: ["Information"],
      bodyRows: cleanRows.map((r) => [r.join(" ")]),
      disclaimers,
      columnAlignments: ["left"],
    };
  }

  // Find table header row: the row with the maximum number of populated cells
  let headerIdx = 0;
  let maxCols = 0;
  const searchLimit = Math.min(candidateTableRows.length, 6);
  for (let i = 0; i < searchLimit; i++) {
    const count = candidateTableRows[i].filter((c) => c !== "").length;
    if (count > maxCols) {
      maxCols = count;
      headerIdx = i;
    }
  }

  const rawHeaders = candidateTableRows[headerIdx] || [];
  // Trim empty trailing columns
  let lastNonEmpty = rawHeaders.length - 1;
  while (lastNonEmpty >= 0 && rawHeaders[lastNonEmpty] === "") {
    lastNonEmpty--;
  }
  const colCount = Math.max(1, lastNonEmpty + 1);

  const headers = rawHeaders.slice(0, colCount).map((h, i) => h || `Col ${i + 1}`);

  // Extract body rows and ensure equal column widths
  const rawBodyRows = candidateTableRows.slice(headerIdx + 1);
  const bodyRows: string[][] = [];

  for (const r of rawBodyRows) {
    const rowSlice = r.slice(0, colCount);
    while (rowSlice.length < colCount) {
      rowSlice.push("");
    }
    // Filter out rows that are completely empty or duplicate disclaimers
    if (rowSlice.some((c) => c !== "") && !isDisclaimerOrNote(rowSlice[0])) {
      bodyRows.push(rowSlice);
    }
  }

  // Calculate column alignments
  const columnAlignments: Array<"left" | "center" | "right"> = headers.map((h, colIdx) => {
    const colValues = bodyRows.map((r) => r[colIdx] || "");
    return detectColumnAlignment(h, colValues);
  });

  return {
    bannerTitle,
    sectionTitle,
    metadata,
    headers,
    bodyRows,
    disclaimers,
    columnAlignments,
  };
}

/**
 * Parses an Excel file buffer and extracts workbook metadata and preview rows.
 */
export function inspectExcelWorkbook(data: ArrayBuffer | Uint8Array): WorkbookInfo {
  const workbook = XLSX.read(data, { type: "array" });
  const sheetNames = workbook.SheetNames;

  const sheets: SheetPreviewData[] = sheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json<(string | number | boolean | null)[]>(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
    });

    const parsed = parseSheetSections(rawData);
    if (!parsed) {
      return {
        name: sheetName,
        totalRows: 0,
        totalCols: 0,
        headers: [],
        previewRows: [],
      };
    }

    return {
      name: sheetName,
      totalRows: parsed.bodyRows.length,
      totalCols: parsed.headers.length,
      headers: parsed.headers,
      previewRows: parsed.bodyRows.slice(0, 10),
    };
  });

  return {
    sheetNames,
    sheets,
  };
}

/**
 * Converts an Excel buffer (or ArrayBuffer) to a beautifully aligned PDF document.
 * Returns the PDF as a Uint8Array.
 */
export function convertExcelToPdf(
  data: ArrayBuffer | Uint8Array,
  options: ExcelToPdfOptions = {}
): { pdfBytes: Uint8Array; pageCount: number } {
  const workbook = XLSX.read(data, { type: "array", cellDates: true });
  const allSheetNames = workbook.SheetNames;
  const sheetsToProcess =
    options.selectedSheets && options.selectedSheets.length > 0
      ? options.selectedSheets.filter((s) => allSheetNames.includes(s))
      : allSheetNames;

  if (sheetsToProcess.length === 0) {
    throw new Error("No valid sheets found to convert.");
  }

  // Pre-parse all selected sheets to determine global orientation
  const parsedSheets: Array<{ sheetName: string; parsed: ParsedSheetContent }> = [];
  let maxColsAcrossSheets = 0;

  for (const sheetName of sheetsToProcess) {
    const sheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
    });
    const parsed = parseSheetSections(rawData);
    if (parsed) {
      parsedSheets.push({ sheetName, parsed });
      if (parsed.headers.length > maxColsAcrossSheets) {
        maxColsAcrossSheets = parsed.headers.length;
      }
    }
  }

  if (parsedSheets.length === 0) {
    throw new Error("No readable data found in the selected Excel sheets.");
  }

  // Determine orientation: if auto, use landscape for sheets with > 6 columns
  let globalOrientation: "portrait" | "landscape" = "portrait";
  if (options.orientation === "landscape") {
    globalOrientation = "landscape";
  } else if (options.orientation === "portrait") {
    globalOrientation = "portrait";
  } else {
    globalOrientation = maxColsAcrossSheets > 6 ? "landscape" : "portrait";
  }

  const doc = new jsPDF({
    orientation: globalOrientation,
    unit: "mm",
    format: "a4",
  });

  const headerColor = options.themeColor || "#1c8a4b"; // Wint primary accent
  let isFirstPage = true;

  for (const { sheetName, parsed } of parsedSheets) {
    const colCount = parsed.headers.length;

    // Per-sheet orientation if auto
    const sheetOrientation =
      options.orientation === "auto" ? (colCount > 6 ? "landscape" : "portrait") : globalOrientation;

    if (!isFirstPage) {
      doc.addPage("a4", sheetOrientation);
    }
    isFirstPage = false;

    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const margin = colCount > 12 ? 8 : 12;
    const printableWidth = pageWidth - margin * 2;

    let currentY = 12;

    // 1. Brand Header / Banner
    doc.setFont("helvetica", "bold");
    doc.setFontSize(13);
    doc.setTextColor(22, 21, 19);
    const mainTitle = parsed.bannerTitle || options.title || "Wint Wealth Report";
    doc.text(mainTitle, margin, currentY);

    // Sheet Name Badge on top right
    const badgeText = `Sheet: ${sheetName}`;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.5);
    doc.setTextColor(28, 138, 75); // accent green
    const badgeWidth = doc.getStringUnitWidth(badgeText) * 8.5 * 0.352778;
    doc.text(badgeText, pageWidth - margin - badgeWidth, currentY);

    currentY += 5;

    // Subtitle / Date stamp
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(120, 120, 120);
    const sub = options.subtitle || `Generated on ${new Date().toLocaleDateString(undefined, { dateStyle: "medium" })}`;
    doc.text(sub, margin, currentY);

    currentY += 5;

    // 2. Metadata Box (User details: Name, Phone, Email, Date of report generation)
    if (parsed.metadata.length > 0) {
      const meta = parsed.metadata;
      const boxHeight = meta.length > 4 ? 13 : 8;
      doc.setFillColor(250, 246, 240); // Wint cream
      doc.setDrawColor(228, 222, 212);
      doc.roundedRect(margin, currentY, printableWidth, boxHeight, 1.2, 1.2, "FD");

      const itemsPerRow = meta.length > 4 ? Math.ceil(meta.length / 2) : meta.length;
      const colStep = printableWidth / itemsPerRow;

      meta.forEach((item, idx) => {
        const rowIdx = meta.length > 4 && idx >= itemsPerRow ? 1 : 0;
        const colIdx = meta.length > 4 && idx >= itemsPerRow ? idx - itemsPerRow : idx;
        const xPos = margin + 3.5 + colIdx * colStep;
        const yPos = currentY + (rowIdx === 0 ? 5.2 : 10.5);

        doc.setFont("helvetica", "bold");
        doc.setFontSize(7.2);
        doc.setTextColor(60, 60, 60);
        const labelText = item.label + ": ";
        doc.text(labelText, xPos, yPos);

        const labelWidth = doc.getStringUnitWidth(labelText) * 7.2 * 0.352778;
        doc.setFont("helvetica", "normal");
        doc.setTextColor(22, 21, 19);
        doc.text(item.value, xPos + labelWidth, yPos);
      });

      currentY += boxHeight + 4;
    }

    // 3. Section Title (e.g. "INVESTMENT SUMMARY FY 25-26")
    if (parsed.sectionTitle) {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(8.5);
      doc.setTextColor(28, 138, 75);
      doc.text(parsed.sectionTitle, margin, currentY);
      currentY += 4;
    }

    // 4. Data Table Configuration
    // Dynamic typography scaling based on column count to prevent vertical squishing
    let tableFontSize = 7.5;
    let cellPadding = 2.0;

    if (colCount <= 6) {
      tableFontSize = 8.5;
      cellPadding = 2.4;
    } else if (colCount <= 9) {
      tableFontSize = 7.5;
      cellPadding = 2.0;
    } else if (colCount <= 13) {
      tableFontSize = 6.6;
      cellPadding = 1.6;
    } else {
      // 14+ columns (e.g. Upcoming Cashflow or Holding Statement)
      tableFontSize = 5.8;
      cellPadding = 1.3;
    }

    const columnStyles: Record<number, { halign: "left" | "center" | "right" }> = {};
    parsed.columnAlignments.forEach((align, idx) => {
      columnStyles[idx] = { halign: align };
    });

    const tableOptions: UserOptions = {
      head: [parsed.headers],
      body: parsed.bodyRows,
      startY: currentY,
      theme: "grid",
      styles: {
        fontSize: tableFontSize,
        cellPadding,
        textColor: [22, 21, 19],
        lineColor: [225, 225, 225],
        lineWidth: 0.1,
        overflow: "linebreak",
        valign: "middle",
      },
      headStyles: {
        fillColor: [28, 138, 75], // #1c8a4b
        textColor: [255, 255, 255],
        fontStyle: "bold",
        fontSize: tableFontSize + 0.3,
        halign: "center",
      },
      alternateRowStyles: {
        fillColor: [250, 246, 240], // cream #faf6f0
      },
      columnStyles,
      margin: { top: 12, right: margin, bottom: 16, left: margin },
      didParseCell: (data) => {
        // Highlight Total / Grand Total rows
        const rawRow = parsed.bodyRows[data.row.index];
        const isTotalRow = rawRow && (rawRow[0]?.toLowerCase().includes("total") || rawRow[1]?.toLowerCase().includes("total"));
        if (isTotalRow) {
          data.cell.styles.fontStyle = "bold";
          data.cell.styles.fillColor = [243, 240, 233];
        }
      },
      didDrawPage: (data) => {
        // Footer: Page number & generation tag
        const totalPages = doc.getNumberOfPages();
        const currentPage = data.pageNumber;

        doc.setFont("helvetica", "normal");
        doc.setFontSize(7.5);
        doc.setTextColor(140, 140, 140);

        // Left footer
        doc.text("Generated by Wint Reports Converter", margin, pageHeight - 6);

        // Right footer
        const pageStr = `Page ${currentPage} of ${totalPages}`;
        const strWidth = doc.getStringUnitWidth(pageStr) * 7.5 * 0.352778;
        doc.text(pageStr, pageWidth - margin - strWidth, pageHeight - 6);
      },
    };

    autoTable(doc, tableOptions);

    // 5. Notes & Disclaimers Section
    // Rendered cleanly below the table with compact, elegant typography (6.2pt font)
    if (parsed.disclaimers.length > 0) {
      let discY = (doc as any).lastAutoTable.finalY + 5;
      const bottomLimit = pageHeight - 16;

      // Estimate required height for disclaimers
      doc.setFont("helvetica", "normal");
      doc.setFontSize(6.2);
      let estimatedHeight = 5; // title height
      for (const disc of parsed.disclaimers) {
        const lines = doc.splitTextToSize(disc, printableWidth);
        estimatedHeight += lines.length * 2.8 + 2;
      }

      // If disclaimers don't fit on the current page, move cleanly to a new page
      if (discY + estimatedHeight > bottomLimit && discY > 70) {
        doc.addPage("a4", sheetOrientation);
        discY = 14;
      }

      // Disclaimers section heading
      doc.setFont("helvetica", "bold");
      doc.setFontSize(6.8);
      doc.setTextColor(80, 80, 80);
      doc.text("IMPORTANT NOTES & DISCLAIMER", margin, discY);
      discY += 3.2;

      // Disclaimers text
      doc.setFont("helvetica", "normal");
      doc.setFontSize(6.2);
      doc.setTextColor(105, 105, 105);

      for (const disc of parsed.disclaimers) {
        const lines = doc.splitTextToSize(disc, printableWidth);
        doc.text(lines, margin, discY);
        discY += lines.length * 2.7 + 2;
      }
    }
  }

  // Update total page numbers across all pages
  const totalPages = doc.getNumberOfPages();
  for (let i = 1; i <= totalPages; i++) {
    doc.setPage(i);
    const pWidth = doc.internal.pageSize.getWidth();
    const pHeight = doc.internal.pageSize.getHeight();
    const margin = pWidth > 250 ? 12 : 10;

    // Clean footer background wipe
    doc.setFillColor(255, 255, 255);
    doc.rect(pWidth - margin - 35, pHeight - 9, 35, 6, "F");

    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(140, 140, 140);
    const pageStr = `Page ${i} of ${totalPages}`;
    const strWidth = doc.getStringUnitWidth(pageStr) * 7.5 * 0.352778;
    doc.text(pageStr, pWidth - margin - strWidth, pHeight - 6);
  }

  const pdfOutput = doc.output("arraybuffer");
  return {
    pdfBytes: new Uint8Array(pdfOutput),
    pageCount: totalPages,
  };
}

/**
 * Creates a sample Investment / Taxation Excel workbook for demonstration/testing.
 * Formatted with realistic Wint Wealth structure including user metadata and disclaimers.
 */
export function generateSampleExcelWorkbook(): Uint8Array {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Master Report
  const masterData = [
    ["", "", "", "", "", "CONSOLIDATED INVESTMENT REPORT WINT WEALTH"],
    ["Name", "Vedant Gopalakrishnan"],
    ["Phone Number", "8637663482"],
    ["Email", "vedant.g@wintwealth.com"],
    ["", "", "", "", "", "INVESTMENT SUMMARY (1/04/2025 - 31/03/2026) FY 25-26"],
    ["Bond Name", "ISIN", "No. Of Units", "Invested Amount", "Face Value", "Acquisition Cost*", "Date of Investment", "Maturity Date", "XIRR", "Frequency of Interest Payment", "Frequency Of Principal Repayment"],
    ["IIFL Samasta", "INE413U07277", 1, 1065.69, 1000, 65.69, "17/10/2025", "21/12/2026", "11.50%", "Annually", "At Maturity"],
    ["Capri Global", "INE180C07221", 1, 997.55, 1000, "-", "01/11/2025", "13/04/2027", "9.50%", "Annually", "At Maturity"],
    ["Total:", "", 2, 2063.24, 2000, 65.69, "", "", "", "", ""],
    ["Date of report generation:", "10/11/2026"],
    ["*Acquisition Cost: When you purchase a bond, the price you pay might be a little higher than the face value. This extra amount – cost of acquisition – includes accrued interest, which is the interest for days you were not invested in the bond (since it was issued). It may also include Wint Wealth spread which is the margin that we earn on the transaction. Note: This cost of acquisition should be deducted as an expense when you file your ITR."],
    ["Disclaimer: The document and details set out thereunder are for informational purposes only being prepared and shared for customer convenience. The data is based on transactions undertaken through Wint Wealth Platform only. There may be certain inaccuracies in the data presented in the document due to (i) transactions or dealings done by the client outside of Wint Wealth; or (ii) any other error in sharing of data by the issuers / trustees to Wint Wealth. You are advised to verify the information in this document basis the above from your demat linked bank account."]
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(masterData);
  XLSX.utils.book_append_sheet(wb, ws1, "Investment Summary");

  // Sheet 2: Repayment Summary
  const repayData = [
    ["", "", "", "", "", "CONSOLIDATED REPAYMENT REPORT WINT WEALTH"],
    ["Name", "Vedant Gopalakrishnan"],
    ["Phone Number", "8637663482"],
    ["Email", "vedant.g@wintwealth.com"],
    ["", "", "", "", "", "REPAYMENT SUMMARY (1/04/2026 - 10/11/2026) FY 26-27"],
    ["Date", "Name Of Bond", "ISIN", "No. Of Units", "Amount in Bank", "Principal Repaid", "Interest Paid (Before TDS)", "Interest Paid (After TDS)", "TDS Deducted"],
    ["13/10/2026", "Capri Global", "INE180C07221", 1, 80.1, 0, 89, 80.1, 8.9],
    ["Total:", "", "", 1, 80.1, 0, 89, 80.1, 8.9],
    ["Date of report generation:", "10/11/2026"],
    ["Disclaimer: The document and details set out thereunder are for informational purposes only being prepared and shared for customer convenience. The data is based on transactions undertaken through Wint Wealth Platform only. You are advised to verify the information from your demat linked bank account."]
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(repayData);
  XLSX.utils.book_append_sheet(wb, ws2, "Repayment Summary");

  // Sheet 3: Taxation Details
  const taxData = [
    ["", "", "", "", "", "TAXATION SCHEDULE & TDS STATEMENT"],
    ["Name", "Vedant Gopalakrishnan"],
    ["PAN Number", "ABCDE1234F"],
    ["Financial Year", "2025-26"],
    ["Payout Date", "Asset Description", "Gross Interest (INR)", "TDS Deducted (10%)", "Net Credited (INR)", "Bank Account"],
    ["15/04/2025", "Muthoot Fincorp Senior NCD", 2465.75, 246.58, 2219.17, "HDFC Bank ****4102"],
    ["20/05/2025", "InCred Financial Services", 1275.34, 127.53, 1147.81, "HDFC Bank ****4102"],
    ["15/06/2025", "Spandana Sphoorty Financial", 2082.19, 208.22, 1873.97, "HDFC Bank ****4102"],
    ["Total Annual", "", 5823.28, 582.33, 5240.95, ""],
    ["Date of report generation:", "10/11/2026"],
    ["Disclaimer: Tax information is for reference and convenience only. Please consult a qualified chartered accountant to validate before filing your income tax return."]
  ];
  const ws3 = XLSX.utils.aoa_to_sheet(taxData);
  XLSX.utils.book_append_sheet(wb, ws3, "Taxation Schedule");

  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return new Uint8Array(buf);
}
