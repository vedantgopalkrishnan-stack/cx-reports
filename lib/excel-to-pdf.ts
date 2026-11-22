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

import {
  parseSheetSections,
  isDisclaimerOrNote,
  isMetadataKey,
  detectColumnAlignment,
  type ParsedSheetContent,
} from "./sheet-parser";

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

  // Sheet 1: Investment Summary
  const masterData = [
    ["CONSOLIDATED INVESTMENT REPORT WINT WEALTH"],
    ["Name", "Anil Vesvikar"],
    ["Phone Number", "9833074106"],
    ["Email", "anilvdlh@gmail.com"],
    [],
    ["INVESTMENT SUMMARY (1/04/2023 - 31/03/2024) FY 23-24"],
    ["Bond Name", "ISIN", "No. Of Units", "Invested Amount", "Face Value", "Acquisition Cost*", "Date of Investment", "Maturity Date", "XIRR", "Frequency of Interest Payment", "Frequency Of Principal Repayment"],
    ["Navi Finserv", "INE342T07411", 3, "₹2,995.26", "₹3,000.00", "-", "24/11/2023", "18/01/2025", "10.50%", "Monthly", "At Maturity"],
    ["", "Total:", 3, "₹2,995.26", "₹3,000.00", "-", "", "", "", "", ""],
    [],
    ["INVESTMENT SUMMARY (1/04/2026 - 11/11/2026) FY 26-27"],
    ["Bond Name", "ISIN", "No. Of Units", "Invested Amount", "Face Value", "Acquisition Cost*", "Date of Investment", "Maturity Date", "XIRR", "Frequency of Interest Payment", "Frequency Of Principal Repayment"],
    ["Muthoot MCred", "INE101Q07BU7", 102, "₹10,00,727.38", "₹10,20,000.00", "-", "13/04/2026", "13/01/2028", "11%", "Monthly", "At Maturity"],
    ["", "Total:", 102, "₹10,00,727.38", "₹10,20,000.00", "-", "", "", "", "", ""],
    [],
    ["Date of report generation:", "11/11/2026"],
    ["*Acquisition Cost: When you purchase a bond, the price you pay might be a little higher than the face value. This extra amount – cost of acquisition – includes accrued interest, which is the interest for days you were not invested in the bond (since it was issued). It may also include Wint Wealth's spread which is the margin that we earn on the transaction. This does not affect the XIRR mentioned for the bond, as the XIRR is calculated after taking into account this cost. Note: This cost of acquisition should be deducted as an expense when you file your ITR. Read more about it here."],
    ["Disclaimer: The document and details set out thereunder are for informational purposes only being prepared and shared for customer convenience. The data is based on transactions undertaken through Wint Wealth Platform only. There may be certain inaccuracies in the data presented in the document due to (i) transactions or dealings done by the client outside of Wint Wealth; or (ii) any other error in sharing of data by the issuers / trustees to Wint Wealth. You are advised to verify the information in this document basis the above from your demat linked bank account."]
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(masterData);
  XLSX.utils.book_append_sheet(wb, ws1, "Investment Summary");

  // Sheet 2: Repayment & Sell Report
  const repayData = [
    ["CONSOLIDATED REPAYMENT REPORT WINT WEALTH"],
    ["Name", "Anil Vesvikar"],
    ["Phone Number", "9833074106"],
    ["Email", "anilvdlh@gmail.com"],
    [],
    ["REPAYMENT SUMMARY (1/04/2023 - 31/03/2024) FY 23-24"],
    ["Date", "Name Of Bond", "ISIN", "No. Of Units", "Amount in Bank", "Principal Repaid", "Interest Paid (Before TDS Deduction)", "Interest Paid (After TDS Deduction)", "TDS Deducted"],
    ["18/12/2023", "Navi Finserv", "INE342T07411", 3, "₹21.58", "₹0.00", "₹23.98", "₹21.58", "₹2.40"],
    ["18/01/2024", "Navi Finserv", "INE342T07411", 3, "₹22.28", "₹0.00", "₹24.78", "₹22.28", "₹2.50"],
    ["16/02/2024", "Navi Finserv", "INE342T07411", 3, "₹22.30", "₹0.00", "₹24.78", "₹22.30", "₹2.48"],
    ["18/03/2024", "Navi Finserv", "INE342T07411", 3, "₹21.00", "₹0.00", "₹23.32", "₹21.00", "₹2.32"],
    ["", "", "Total:", 12, "₹87.16", "₹0.00", "₹96.86", "₹87.16", "₹9.70"],
    [],
    ["REPAYMENT SUMMARY (1/04/2024 - 31/03/2025) FY 24-25"],
    ["Date", "Name Of Bond", "ISIN", "No. Of Units", "Amount in Bank", "Principal Repaid", "Interest Paid (Before TDS Deduction)", "Interest Paid (After TDS Deduction)", "TDS Deducted"],
    ["18/04/2024", "Navi Finserv", "INE342T07411", 3, "₹22.30", "₹0.00", "₹24.78", "₹22.30", "₹2.48"],
    ["17/05/2024", "Navi Finserv", "INE342T07411", 3, "₹21.57", "₹0.00", "₹23.97", "₹21.57", "₹2.40"],
    ["18/06/2024", "Navi Finserv", "INE342T07411", 3, "₹22.30", "₹0.00", "₹24.78", "₹22.30", "₹2.48"],
    ["18/07/2024", "Navi Finserv", "INE342T07411", 3, "₹21.57", "₹0.00", "₹23.97", "₹21.57", "₹2.40"],
    ["16/08/2024", "Navi Finserv", "INE342T07411", 3, "₹22.36", "₹0.00", "₹24.84", "₹22.36", "₹2.48"],
    ["17/09/2024", "Navi Finserv", "INE342T07411", 3, "₹22.36", "₹0.00", "₹24.84", "₹22.36", "₹2.48"],
    ["18/10/2024", "Navi Finserv", "INE342T07411", 3, "₹21.63", "₹0.00", "₹24.03", "₹21.63", "₹2.40"],
    ["18/11/2024", "Navi Finserv", "INE342T07411", 3, "₹22.36", "₹0.00", "₹24.84", "₹22.36", "₹2.48"],
    ["18/12/2024", "Navi Finserv", "INE342T07411", 3, "₹21.63", "₹0.00", "₹24.03", "₹21.63", "₹2.40"],
    ["17/01/2025", "Navi Finserv", "INE342T07411", 3, "₹3,022.36", "₹3,000.00", "₹24.84", "₹22.36", "₹2.48"],
    ["", "", "Total:", 30, "₹3,220.44", "₹3,000.00", "₹244.92", "₹220.44", "₹24.48"],
    [],
    ["REPAYMENT SUMMARY (1/04/2026 - 11/11/2026) FY 26-27"],
    ["Date", "Name Of Bond", "ISIN", "No. Of Units", "Amount in Bank", "Principal Repaid", "Interest Paid (Before TDS Deduction)", "Interest Paid (After TDS Deduction)", "TDS Deducted"],
    ["09/05/2026", "Muthoot MCred", "INE101Q07BU7", 102, "₹6,979.55", "₹0.00", "₹7,754.75", "₹6,979.55", "₹775.20"],
    ["10/06/2026", "Muthoot MCred", "INE101Q07BU7", 102, "₹7,211.81", "₹0.00", "₹8,013.53", "₹7,211.81", "₹801.72"],
    ["10/07/2026", "Muthoot MCred", "INE101Q07BU7", 102, "₹6,979.55", "₹0.00", "₹7,754.75", "₹6,979.55", "₹775.20"],
    ["10/08/2026", "Muthoot MCred", "INE101Q07BU7", 102, "₹7,211.81", "₹0.00", "₹8,013.53", "₹7,211.81", "₹801.72"],
    ["10/09/2026", "Muthoot MCred", "INE101Q07BU7", 102, "₹7,211.81", "₹0.00", "₹8,013.53", "₹7,211.81", "₹801.72"],
    ["09/10/2026", "Muthoot MCred", "INE101Q07BU7", 102, "₹6,979.55", "₹0.00", "₹7,754.75", "₹6,979.55", "₹775.20"],
    ["09/11/2026", "Muthoot MCred", "INE101Q07BU7", 102, "₹7,211.81", "₹0.00", "₹8,013.53", "₹7,211.81", "₹801.72"],
    ["", "", "Total:", 714, "₹49,785.89", "₹0.00", "₹55,318.37", "₹49,785.89", "₹5,532.48"],
    [],
    ["Date of report generation:", "11/11/2026"],
    ["Disclaimer: The document and details set out thereunder are for informational purposes only being prepared and shared for customer convenience. The data is based on transactions undertaken through Wint Wealth Platform only. There may be certain inaccuracies in the data presented in the document due to (i) transactions or dealings done by the client outside of Wint Wealth; or (ii) any other error in sharing of data by the issuers / trustees to Wint Wealth. You are advised to verify the information in this document basis the above from your demat linked bank account."],
    [],
    ["CONSOLIDATED SELL REPORT WINT WEALTH"],
    ["Name", "Anil Vesvikar"],
    ["Phone Number", "9833074106"],
    ["Email", "anilvdlh@gmail.com"],
    [],
    ["Date of report generation:", "11/11/2026"],
    ["Disclaimer: The document and details set out thereunder are for informational purposes only being prepared and shared for customer convenience. The data is based on transactions undertaken through Wint Wealth Platform only. There may be certain inaccuracies in the data presented in the document due to (i) transactions or dealings done by the client outside of Wint Wealth; or (ii) any other error in sharing of data by the issuers / trustees to Wint Wealth. You are advised to verify the information in this document basis the above from your demat linked bank account."]
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(repayData);
  XLSX.utils.book_append_sheet(wb, ws2, "Repayment Summary");

  // Sheet 3: Upcoming Cashflow
  const cashflowData = [
    ["Wint Wealth"],
    ["Name", "Anil Vesvikar"],
    ["Phone Number", "9833074106"],
    ["Email", "anilvdlh@gmail.com"],
    [],
    ["Upcoming Cashflow Statement"],
    ["Name Of Bond", "ISIN", "Date", "Amount in Bank", "Principal Repaid", "Interest Paid (After TDS deduction)", "Interest Paid (Before TDS deduction)", "TDS", "Sell Value", "Transactiontype"],
    ["Muthoot MCred", "INE101Q07BU7", "10-12-2026", "₹6,979.55", "", "₹6,979.55", "₹7,755.06", "₹775.51", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "10-01-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "10-02-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "10-03-2027", "₹6,514.13", "", "₹6,514.13", "₹7,237.92", "₹723.79", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "10-04-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "10-05-2027", "₹6,979.55", "", "₹6,979.55", "₹7,755.06", "₹775.51", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "10-06-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "10-07-2027", "₹6,979.55", "", "₹6,979.55", "₹7,755.06", "₹775.51", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "10-08-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "10-09-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "10-10-2027", "₹6,979.55", "", "₹6,979.55", "₹7,755.06", "₹775.51", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "10-11-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "10-12-2027", "₹6,979.55", "", "₹6,979.55", "₹7,755.06", "₹775.51", "", "INTEREST"],
    ["Muthoot MCred", "INE101Q07BU7", "13-01-2028", "₹10,27,909.49", "₹10,20,000.00", "₹7,909.49", "₹8,788.32", "₹878.83", "", "PRINCIPAL AND INTEREST"],
    [],
    ["Financial Year Summary (Upcoming Cashflow)"],
    ["Financial Year", "Amount in Bank", "Principal", "Interest (After TDS Deduction)", "Interest (Before TDS Deduction)", "TDS", "Sell Value"],
    ["2026-2027", "₹27,917.30", "", "₹27,917.30", "₹31,019.22", "₹3,101.92", ""],
    ["2027-2028", "₹10,91,886.74", "₹10,20,000.00", "₹71,886.74", "₹79,874.16", "₹7,987.42", ""],
    [],
    ["Monthly Summary (Upcoming Cashflow)"],
    ["Month", "Amount in Bank", "Principal", "Interest (After TDS Deduction)", "Interest (Before TDS Deduction)", "TDS Deducted", "Sell Value"],
    ["December-2026", "₹6,979.55", "", "₹6,979.55", "₹7,755.06", "₹775.51", ""],
    ["January-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", ""],
    ["February-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", ""],
    ["March-2027", "₹6,514.13", "", "₹6,514.13", "₹7,237.92", "₹723.79", ""],
    ["April-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", ""],
    ["May-2027", "₹6,979.55", "", "₹6,979.55", "₹7,755.06", "₹775.51", ""],
    ["June-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", ""],
    ["July-2027", "₹6,979.55", "", "₹6,979.55", "₹7,755.06", "₹775.51", ""],
    ["August-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", ""],
    ["September-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", ""],
    ["October-2027", "₹6,979.55", "", "₹6,979.55", "₹7,755.06", "₹775.51", ""],
    ["November-2027", "₹7,211.81", "", "₹7,211.81", "₹8,013.12", "₹801.31", ""],
    ["December-2027", "₹6,979.55", "", "₹6,979.55", "₹7,755.06", "₹775.51", ""],
    ["January-2028", "₹10,27,909.49", "₹10,20,000.00", "₹7,909.49", "₹8,788.32", "₹878.83", ""]
  ];
  const ws3 = XLSX.utils.aoa_to_sheet(cashflowData);
  XLSX.utils.book_append_sheet(wb, ws3, "Upcoming Cashflow");

  // Sheet 4: Bonds Holding Statement
  const holdingData = [
    ["Wint Wealth"],
    ["Name", "Anil Vesvikar"],
    ["Phone Number", "9833074106"],
    ["Email", "anilvdlh@gmail.com"],
    [],
    ["Bonds Holding Statement"],
    ["Name Of Bond", "ISIN", "Maturity Date", "Units", "YTM", "Current Value", "Upcoming Sell Value", "Upcoming interest", "Upcoming Principal", "Total Invested", "Total Sold", "Principal Repaid till Date", "Interest Paid (Before TDS Deduction) till Date", "Interest Paid (After TDS Deduction) till Date", "TDS Deducted till Date", "Principal Repayment Type", "Interest Repayment Type"],
    ["Muthoot MCred", "INE101Q07BU7", "13-01-2028", 102, 11.0, "₹10,06,442.27", "", "₹1,10,893.38", "₹10,20,000.00", "₹10,00,727.38", "", "", "₹55,318.37", "₹49,785.89", "₹5,532.48", "At Maturity", "Monthly"],
    ["Total", "", "", "", 11.0, "₹10,06,442.27", "", "₹1,10,893.38", "₹10,20,000.00", "₹10,00,727.38", "", "", "₹55,318.37", "₹49,785.89", "₹5,532.48", "", ""],
    [],
    ["Date of report generation:", "11/11/2026"],
    ["Disclaimer:"],
    ["The information provided in this Consolidated Tax Statement is for your reference and convenience only. It is not intended to serve as tax advice, and Wint Wealth is not a tax advisor. Please review the following disclaimers carefully:"],
    ["Before filing your Income Tax Return (ITR), we strongly recommend reconciling this statement with form 26 AS and your demat linked bank account as we cannot guarantee 100% accuracy."],
    ["We have classified and accounted for income, interest, principal repayments, and capital gains based on standard tax treatment for listed and taxable bonds in India."],
    ["All information provided herein is in relation to transactions made on the Wint Wealth platform alone and does not make reference to other trades made in one's demat account."],
    ["Tax laws and interpretations may vary based on your individual circumstances, and the actual tax liability may differ from the estimates provided here."],
    ["Please consult a qualified tax advisor or chartered accountant to validate the same."],
    ["Wint Wealth disclaims any liability arising from tax misfilings, discrepancies, or interpretations based on this statement."]
  ];
  const ws4 = XLSX.utils.aoa_to_sheet(holdingData);
  XLSX.utils.book_append_sheet(wb, ws4, "Holding Statement");

  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return new Uint8Array(buf);
}
