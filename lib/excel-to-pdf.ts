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

/**
 * Parses an Excel file buffer or array buffer and extracts workbook metadata and preview rows.
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

    const headers = (rawData[0] || []).map((h) => (h !== null && h !== undefined ? String(h).trim() : ""));
    const previewRows = rawData.slice(1, 11); // first 10 rows for preview
    const totalCols = Math.max(...rawData.map((r) => r.length), headers.length);

    return {
      name: sheetName,
      totalRows: Math.max(0, rawData.length - 1),
      totalCols,
      headers,
      previewRows,
    };
  });

  return {
    sheetNames,
    sheets,
  };
}

/**
 * Converts an Excel buffer (or ArrayBuffer) to a PDF document.
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

  // Determine overall orientation
  let initialOrientation: "portrait" | "landscape" = "portrait";
  if (options.orientation === "landscape") {
    initialOrientation = "landscape";
  } else if (options.orientation === "portrait") {
    initialOrientation = "portrait";
  } else {
    // Auto: inspect max columns across chosen sheets
    let maxCols = 0;
    for (const sheetName of sheetsToProcess) {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, { header: 1, blankrows: false });
      for (const row of rows) {
        if (Array.isArray(row) && row.length > maxCols) {
          maxCols = row.length;
        }
      }
    }
    initialOrientation = maxCols > 6 ? "landscape" : "portrait";
  }

  const doc = new jsPDF({
    orientation: initialOrientation,
    unit: "mm",
    format: "a4",
  });

  const headerColor = options.themeColor || "#1c8a4b"; // Wint primary accent
  let isFirstSheet = true;

  for (const sheetName of sheetsToProcess) {
    const sheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
    });

    if (rawData.length === 0) continue;

    // Filter out completely empty rows
    const cleanedRows = rawData
      .map((row) =>
        (Array.isArray(row) ? row : []).map((cell) =>
          cell !== null && cell !== undefined ? String(cell).trim() : ""
        )
      )
      .filter((row) => row.some((val) => val !== ""));

    if (cleanedRows.length === 0) continue;

    // If not first sheet, add a new page
    if (!isFirstSheet) {
      const sheetOrientation =
        options.orientation === "auto"
          ? (cleanedRows[0]?.length || 0) > 6
            ? "landscape"
            : "portrait"
          : initialOrientation;
      doc.addPage("a4", sheetOrientation);
    }
    isFirstSheet = false;

    // Page dimensions
    const pageWidth = doc.internal.pageSize.getWidth();
    let currentY = 14;

    // Title / Header section on sheet page
    if (options.title) {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(14);
      doc.setTextColor(22, 21, 19); // ink
      doc.text(options.title, 14, currentY);
      currentY += 6;
    }

    if (options.subtitle) {
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.setTextColor(100, 100, 100);
      doc.text(options.subtitle, 14, currentY);
      currentY += 5;
    }

    // Sheet Name Badge / Subtitle
    if (sheetsToProcess.length > 1 || !options.title) {
      doc.setFont("helvetica", "bold");
      doc.setFontSize(11);
      doc.setTextColor(28, 138, 75); // accent green
      doc.text(`Sheet: ${sheetName}`, 14, currentY);
      currentY += 6;
    }

    const headers = cleanedRows[0] || [];
    const bodyRows = cleanedRows.slice(1);

    // AutoTable options
    const tableOptions: UserOptions = {
      head: [headers],
      body: bodyRows,
      startY: currentY,
      theme: "grid",
      styles: {
        fontSize: options.fontSize || 8,
        cellPadding: 2.5,
        textColor: [22, 21, 19],
        lineColor: [220, 220, 220],
        lineWidth: 0.1,
        overflow: "linebreak",
        valign: "middle",
      },
      headStyles: {
        fillColor: [28, 138, 75], // #1c8a4b
        textColor: [255, 255, 255],
        fontStyle: "bold",
        fontSize: (options.fontSize || 8) + 0.5,
      },
      alternateRowStyles: {
        fillColor: [250, 246, 240], // cream #faf6f0
      },
      margin: { top: 15, right: 12, bottom: 18, left: 12 },
      didDrawPage: (data) => {
        // Footer: Page number & timestamp
        const pageCount = doc.getNumberOfPages();
        const currentPage = data.pageNumber;
        const pageHeight = doc.internal.pageSize.getHeight();

        doc.setFont("helvetica", "normal");
        doc.setFontSize(8);
        doc.setTextColor(140, 140, 140);

        // Left footer
        doc.text("Generated by Wint Reports Converter", 12, pageHeight - 8);

        // Right footer
        const pageStr = `Page ${currentPage} of ${pageCount}`;
        const strWidth = doc.getStringUnitWidth(pageStr) * 8 * 0.352778;
        doc.text(pageStr, pageWidth - 12 - strWidth, pageHeight - 8);
      },
    };

    autoTable(doc, tableOptions);
  }

  const pdfOutput = doc.output("arraybuffer");
  return {
    pdfBytes: new Uint8Array(pdfOutput),
    pageCount: doc.getNumberOfPages(),
  };
}

/**
 * Creates a sample Investment / Taxation Excel workbook for demonstration/testing.
 */
export function generateSampleExcelWorkbook(): Uint8Array {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Master Report
  const masterData = [
    ["User ID", "941458", "Report Generated", new Date().toISOString().slice(0, 10)],
    ["", "", "", ""],
    ["Asset Name", "ISIN", "Units", "Invested Amount (INR)", "Current Value (INR)", "XIRR (%)", "Status"],
    ["Muthoot Fincorp Senior NCD", "INE549K07890", 100, 100000, 104750, "9.85%", "Active"],
    ["InCred Financial Services", "INE945W07234", 50, 50000, 52100, "10.20%", "Active"],
    ["Spandana Sphoorty Financial", "INE572J07412", 75, 75000, 78600, "11.10%", "Active"],
    ["Navi Finserv Secured Bond", "INE342T07123", 25, 25000, 25850, "9.75%", "Active"],
    ["Edelweiss Financial Services", "INE532F07BY2", 150, 150000, 158200, "10.05%", "Active"],
    ["Total Portfolio", "", 400, 400000, 419500, "10.15%", "Verified"],
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(masterData);
  XLSX.utils.book_append_sheet(wb, ws1, "Portfolio Summary");

  // Sheet 2: Taxation Details
  const taxData = [
    ["FY 2025-26 Interest Payouts & TDS Schedule", "", "", "", "", ""],
    ["", "", "", "", "", ""],
    ["Payout Date", "Asset Description", "Gross Interest (INR)", "TDS Deducted (10%)", "Net Credited (INR)", "Bank Account"],
    ["2025-04-15", "Muthoot Fincorp Senior NCD", 2465.75, 246.58, 2219.17, "HDFC Bank ****4102"],
    ["2025-05-20", "InCred Financial Services", 1275.34, 127.53, 1147.81, "HDFC Bank ****4102"],
    ["2025-06-15", "Spandana Sphoorty Financial", 2082.19, 208.22, 1873.97, "HDFC Bank ****4102"],
    ["2025-07-15", "Muthoot Fincorp Senior NCD", 2465.75, 246.58, 2219.17, "HDFC Bank ****4102"],
    ["2025-08-20", "InCred Financial Services", 1275.34, 127.53, 1147.81, "HDFC Bank ****4102"],
    ["Total Annual", "", 9564.37, 956.44, 8607.93, ""],
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(taxData);
  XLSX.utils.book_append_sheet(wb, ws2, "Taxation Schedule");

  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return new Uint8Array(buf);
}
