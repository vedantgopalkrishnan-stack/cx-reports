import * as XLSX from "xlsx";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  Header,
  Footer,
  ImageRun,
  WidthType,
  PageOrientation,
  AlignmentType,
  BorderStyle,
  PageBreak,
} from "docx";
import { WINT_LOGO_BASE64, WINT_BAR_BASE64 } from "./assets-base64";

export interface ExcelToDocxOptions {
  orientation?: "portrait" | "landscape" | "auto";
  title?: string;
  selectedSheets?: string[];
  themeColor?: string;
}

type AlignmentTypeValue = (typeof AlignmentType)[keyof typeof AlignmentType];

function decodeBase64ToUint8Array(base64: string): Uint8Array {
  if (typeof Buffer !== "undefined" && typeof Buffer.from === "function") {
    const buf = Buffer.from(base64, "base64");
    return new Uint8Array(buf.buffer, buf.byteOffset, buf.length);
  }
  const binaryString = atob(base64);
  const bytes = new Uint8Array(binaryString.length);
  for (let i = 0; i < binaryString.length; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes;
}

function getLogoAsset(): Uint8Array {
  return decodeBase64ToUint8Array(WINT_LOGO_BASE64);
}

function getBarAsset(): Uint8Array {
  return decodeBase64ToUint8Array(WINT_BAR_BASE64);
}

const TABLE_BORDER_STYLE = {
  style: BorderStyle.SINGLE,
  size: 4,
  color: "000000",
};

const NO_BORDER_STYLE = {
  style: BorderStyle.NONE,
};

// Known exact column grids matching Master_Report_(1).docx (in dxa)
const KNOWN_GRIDS: Record<string, number[]> = {
  investment: [1413, 2009, 1160, 1560, 1795, 1620, 1217, 1280, 895, 2407, 3067, 2962],
  acquisition_cost: [5082, 16459],
  repayment: [1271, 2009, 1651, 1018, 1276, 1417, 2126, 1985, 1701],
  sell: [2263, 2127, 8363, 1701],
  cashflow: [1520, 2009, 1120, 1620, 1640, 3460, 3640, 939, 1020, 2640],
  "cashflow-fy": [1731, 1952, 1635, 3615, 3823, 1148, 1217],
  "cashflow-monthly": [1806, 1858, 1556, 3442, 3640, 1660, 1158],
  holding: [1265, 1943, 965, 698, 623, 1449, 1158, 1341, 1449, 1449, 676, 1030, 1223, 1223, 1094, 1255, 1255],
};

function createCell(
  text: string,
  options: {
    columnSpan?: number;
    bold?: boolean;
    align?: AlignmentTypeValue;
    size?: number; // in half-points: 20 = 10pt
    font?: string;
    color?: string;
    widthDxa?: number;
  } = {}
): TableCell {
  const {
    columnSpan,
    bold = false,
    align = AlignmentType.LEFT,
    size = 20,
    font = "Helvetica",
    color = "000000",
    widthDxa,
  } = options;

  return new TableCell({
    columnSpan,
    width: widthDxa ? { size: widthDxa, type: WidthType.DXA } : undefined,
    borders: {
      top: TABLE_BORDER_STYLE,
      bottom: TABLE_BORDER_STYLE,
      left: TABLE_BORDER_STYLE,
      right: TABLE_BORDER_STYLE,
    },
    children: [
      new Paragraph({
        alignment: align,
        spacing: { after: 0, line: 240, lineRule: "auto" },
        children: [
          new TextRun({
            text,
            bold,
            size,
            font,
            color,
          }),
        ],
      }),
    ],
  });
}

function detectColumnAlignment(header: string, values: unknown[]): AlignmentTypeValue {
  const hLower = String(header).toLowerCase();

  if (
    hLower.includes("date") ||
    hLower === "isin" ||
    hLower.includes("pan") ||
    hLower === "fy" ||
    hLower.includes("year") ||
    hLower.includes("month") ||
    hLower.includes("frequency") ||
    hLower.includes("status") ||
    hLower.includes("type") ||
    hLower === "transactiontype"
  ) {
    return AlignmentType.CENTER;
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
    hLower.includes("price") ||
    hLower.includes("invested")
  ) {
    return AlignmentType.RIGHT;
  }

  let numericCount = 0;
  let dateCount = 0;
  let nonEmptyCount = 0;

  for (const v of values) {
    const val = String(v ?? "").trim();
    if (!val || val === "-") continue;
    nonEmptyCount++;

    if (/^\d{1,4}[-/.]\d{1,2}[-/.]\d{2,4}$/.test(val) || /^\d{4}-\d{4}$/.test(val)) {
      dateCount++;
      continue;
    }

    const cleanNum = val.replace(/^[₹$€£]\s?/, "").replace(/%$/, "").replace(/,/g, "").trim();
    if (!isNaN(Number(cleanNum))) {
      numericCount++;
    }
  }

  if (nonEmptyCount > 0) {
    if (numericCount / nonEmptyCount > 0.6) return AlignmentType.RIGHT;
    if (dateCount / nonEmptyCount > 0.6) return AlignmentType.CENTER;
  }

  return AlignmentType.LEFT;
}

/**
 * Creates the standard Wint header matching Master_Report_(1).docx:
 * - Table width 21962 dxa with invisible borders
 * - Left column (4228 dxa): Wint logo (146pt x 77pt)
 * - Right column (17734 dxa): Right-aligned company details in Poppins
 * - Below table: Full-width decorative bar (1087pt x 5.25pt)
 */
function createWintHeader(logoArr: Uint8Array, barArr: Uint8Array): Header {
  return new Header({
    children: [
      new Table({
        width: { size: 21962, type: WidthType.DXA },
        columnWidths: [4228, 17734],
        borders: {
          top: NO_BORDER_STYLE,
          bottom: NO_BORDER_STYLE,
          left: NO_BORDER_STYLE,
          right: NO_BORDER_STYLE,
          insideHorizontal: NO_BORDER_STYLE,
          insideVertical: NO_BORDER_STYLE,
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 4228, type: WidthType.DXA },
                borders: {
                  top: NO_BORDER_STYLE,
                  bottom: NO_BORDER_STYLE,
                  left: NO_BORDER_STYLE,
                  right: NO_BORDER_STYLE,
                },
                children: [
                  new Paragraph({
                    children: [
                      new ImageRun({
                        data: logoArr,
                        type: "png",
                        transformation: { width: 146, height: 77 },
                      }),
                    ],
                  }),
                ],
              }),
              new TableCell({
                width: { size: 17734, type: WidthType.DXA },
                borders: {
                  top: NO_BORDER_STYLE,
                  bottom: NO_BORDER_STYLE,
                  left: NO_BORDER_STYLE,
                  right: NO_BORDER_STYLE,
                },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    spacing: { after: 20 },
                    children: [
                      new TextRun({
                        text: "WINT SECURITIES PRIVATE LIMITED (Formerly known as Fourdegreewater Services Private Limited)",
                        bold: true,
                        size: 20,
                        font: "Poppins",
                        color: "000000",
                      }),
                    ],
                  }),
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    spacing: { after: 20 },
                    children: [
                      new TextRun({
                        text: "CIN: U66120KA2022PTC165599",
                        size: 20,
                        font: "Poppins",
                        color: "000000",
                      }),
                    ],
                  }),
                  new Paragraph({
                    alignment: AlignmentType.RIGHT,
                    children: [
                      new TextRun({
                        text: "Website: https://www.wintwealth.com",
                        size: 20,
                        font: "Poppins",
                        color: "000000",
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
      new Paragraph({
        spacing: { before: 40, after: 100 },
        children: [
          new ImageRun({
            data: barArr,
            type: "png",
            transformation: { width: 1087, height: 5.25 },
          }),
        ],
      }),
    ],
  });
}

/**
 * Creates the standard Wint footer matching Master_Report_(1).docx:
 * - Top decorative bar (1085pt x 3.6pt)
 * - 3-column table (23804 dxa):
 *   - Col 1 (13031 dxa): Registered address in Bangalore
 *   - Col 2 (2588 dxa): Email (compliance@wintwealth.com)
 *   - Col 3 (8185 dxa): Contact (+91-8861217554)
 * - Thin vertical dividers (#EFEFEF)
 */
function createWintFooter(barArr: Uint8Array): Footer {
  return new Footer({
    children: [
      new Paragraph({
        spacing: { before: 20, after: 40 },
        children: [
          new ImageRun({
            data: barArr,
            type: "png",
            transformation: { width: 1085, height: 3.6 },
          }),
        ],
      }),
      new Table({
        width: { size: 23804, type: WidthType.DXA },
        columnWidths: [13031, 2588, 8185],
        borders: {
          top: NO_BORDER_STYLE,
          bottom: NO_BORDER_STYLE,
          left: NO_BORDER_STYLE,
          right: NO_BORDER_STYLE,
          insideHorizontal: NO_BORDER_STYLE,
          insideVertical: { style: BorderStyle.SINGLE, size: 8, color: "EFEFEF" },
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({
                width: { size: 13031, type: WidthType.DXA },
                margins: { top: 100, bottom: 100, left: 100, right: 100 },
                borders: {
                  top: NO_BORDER_STYLE,
                  bottom: NO_BORDER_STYLE,
                  left: NO_BORDER_STYLE,
                  right: { style: BorderStyle.SINGLE, size: 8, color: "EFEFEF" },
                },
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: "Reg. Address: The Pavilion, #175, Dollars Colony, Near Rainbow Hospital, Bannerghatta Main Road, Bannerghatta Road, Bangalore South, Bangalore - 560076, Karnataka, India.",
                        size: 16,
                        font: "Poppins",
                        color: "000000",
                      }),
                    ],
                  }),
                ],
              }),
              new TableCell({
                width: { size: 2588, type: WidthType.DXA },
                margins: { top: 100, bottom: 100, left: 100, right: 100 },
                borders: {
                  top: NO_BORDER_STYLE,
                  bottom: NO_BORDER_STYLE,
                  left: { style: BorderStyle.SINGLE, size: 8, color: "EFEFEF" },
                  right: { style: BorderStyle.SINGLE, size: 8, color: "EFEFEF" },
                },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: " Email: compliance@wintwealth.com",
                        size: 16,
                        font: "Poppins",
                        color: "000000",
                      }),
                    ],
                  }),
                ],
              }),
              new TableCell({
                width: { size: 8185, type: WidthType.DXA },
                margins: { top: 100, bottom: 100, left: 100, right: 100 },
                borders: {
                  top: NO_BORDER_STYLE,
                  bottom: NO_BORDER_STYLE,
                  left: { style: BorderStyle.SINGLE, size: 8, color: "EFEFEF" },
                  right: NO_BORDER_STYLE,
                },
                children: [
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: " Contact: ",
                        size: 16,
                        font: "Poppins",
                        color: "000000",
                      }),
                    ],
                  }),
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "+91-8861217554",
                        size: 16,
                        font: "Poppins",
                        color: "000000",
                      }),
                    ],
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function createSigningBlock(): Paragraph[] {
  return [
    new Paragraph({
      spacing: { before: 120, after: 20 },
      children: [
        new TextRun({
          text: "For and on behalf of Wint Securities Private Limited",
          bold: true,
          size: 20,
          font: "Helvetica",
          color: "000000",
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun({
          text: "(formerly known as Fourdegreewater Services Private Limited)",
          bold: true,
          size: 20,
          font: "Helvetica",
          color: "000000",
        }),
      ],
    }),
    new Paragraph({ text: "" }),
    new Paragraph({ text: "" }),
    new Paragraph({
      spacing: { after: 20 },
      children: [
        new TextRun({
          text: "Name: Rohit Kumar Jha",
          bold: true,
          size: 20,
          font: "Helvetica",
          color: "000000",
        }),
      ],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: "Designation: Company Secretary and Compliance Officer",
          bold: true,
          size: 20,
          font: "Helvetica",
          color: "000000",
        }),
      ],
    }),
  ];
}

interface RawSection {
  type: string;
  rows: string[][];
}

function segmentRowsIntoSections(cleanRows: string[][]): RawSection[] {
  const sections: RawSection[] = [];
  let curSec: RawSection | null = null;

  for (let i = 0; i < cleanRows.length; i++) {
    const row = cleanRows[i];
    const nonBlank = row.filter((c) => c !== "");
    const rowText = nonBlank.join(" ").trim().toLowerCase();

    // Check if this row initiates a new major report section
    let newSecType: string | null = null;

    if (rowText.includes("consolidated investment report")) {
      newSecType = "investment";
    } else if (rowText.includes("consolidated repayment report")) {
      newSecType = "repayment";
    } else if (rowText.includes("consolidated sell report")) {
      newSecType = "sell";
    } else if (rowText.includes("financial year summary (upcoming cashflow)")) {
      newSecType = "cashflow-fy";
    } else if (rowText.includes("monthly summary (upcoming cashflow)")) {
      newSecType = "cashflow-monthly";
    } else if (rowText === "wint wealth") {
      // Look ahead up to 8 rows to see whether this starts cashflow or holding
      const nextWindow = cleanRows.slice(i, i + 8).map((r) => r.join(" ").toLowerCase());
      if (nextWindow.some((t) => t.includes("bonds holding"))) {
        newSecType = "holding";
      } else if (nextWindow.some((t) => t.includes("upcoming cashflow"))) {
        newSecType = "cashflow";
      }
    } else if (rowText.includes("bonds holding statement") && (!curSec || curSec.type !== "holding")) {
      newSecType = "holding";
    } else if (rowText.includes("upcoming cashflow statement") && (!curSec || curSec.type !== "cashflow")) {
      newSecType = "cashflow";
    }

    if (newSecType) {
      if (curSec && curSec.rows.length > 0) {
        sections.push(curSec);
      }
      curSec = { type: newSecType, rows: [] };
    }

    if (!curSec) {
      curSec = { type: "general", rows: [] };
    }
    curSec.rows.push(row);
  }

  if (curSec && curSec.rows.length > 0) {
    sections.push(curSec);
  }

  return sections;
}

/**
 * Creates the secondary Table 2 on Page 1 for *Acquisition Cost and Disclaimer.
 */
function createAcquisitionCostTable(costText: string, disclaimerText: string): Table {
  const colGrid = KNOWN_GRIDS.acquisition_cost;
  const rows: TableRow[] = [];

  if (costText) {
    rows.push(
      new TableRow({
        cantSplit: true,
        children: [
          createCell("*Acquisition Cost:", {
            bold: true,
            size: 20,
            widthDxa: colGrid[0],
          }),
          createCell(costText, {
            size: 22,
            widthDxa: colGrid[1],
          }),
        ],
      })
    );
  }

  if (disclaimerText) {
    rows.push(
      new TableRow({
        cantSplit: true,
        children: [
          createCell(disclaimerText, {
            columnSpan: 2,
            size: 20,
            widthDxa: colGrid[0] + colGrid[1],
          }),
        ],
      })
    );
  }

  return new Table({
    width: { size: colGrid[0] + colGrid[1], type: WidthType.DXA },
    columnWidths: colGrid,
    borders: {
      top: TABLE_BORDER_STYLE,
      bottom: TABLE_BORDER_STYLE,
      left: TABLE_BORDER_STYLE,
      right: TABLE_BORDER_STYLE,
      insideHorizontal: TABLE_BORDER_STYLE,
      insideVertical: TABLE_BORDER_STYLE,
    },
    rows,
  });
}

function renderSectionToTable(sec: RawSection): {
  table: Table;
  acquisitionCostTable?: Table;
} {
  const rows = sec.rows;
  const colGrid = KNOWN_GRIDS[sec.type];

  // For Page 1 (Investment report), extract Acquisition Cost and Disclaimer into Table 2
  let acCostText = "";
  let acDisclaimerText = "";
  const tableSourceRows: string[][] = [];

  if (sec.type === "investment") {
    for (const r of rows) {
      const nonBlank = r.filter((c) => c !== "");
      const fullText = nonBlank.join(" ").trim();
      const firstLower = (nonBlank[0] || "").toLowerCase();

      if (firstLower.startsWith("*acquisition cost") || firstLower.startsWith("acquisition cost")) {
        const colonIdx = fullText.indexOf(":");
        acCostText = colonIdx !== -1 ? fullText.slice(colonIdx + 1).trim() : fullText;
      } else if (firstLower.startsWith("disclaimer") || firstLower.includes("the document and details")) {
        acDisclaimerText = fullText;
      } else {
        tableSourceRows.push(r);
      }
    }
  } else {
    tableSourceRows.push(...rows);
  }

  // Determine column count and widths
  let colCount = colGrid ? colGrid.length : 0;
  if (!colCount) {
    let maxCols = 0;
    for (const r of tableSourceRows) {
      const nonBlank = r.filter((c) => c !== "");
      if (nonBlank.length > maxCols) maxCols = nonBlank.length;
    }
    colCount = Math.max(maxCols, 2);
  }

  // Alignments
  let headerRow: string[] | null = null;
  for (const r of tableSourceRows) {
    const nonBlank = r.filter((c) => c !== "");
    const firstLower = (nonBlank[0] || "").toLowerCase();
    const isSpecial =
      firstLower.includes("report") ||
      firstLower.includes("summary") ||
      firstLower.includes("statement") ||
      firstLower.includes("name") ||
      firstLower.includes("phone") ||
      firstLower.includes("email") ||
      firstLower.includes("date") ||
      firstLower.includes("disclaimer") ||
      firstLower === "wint wealth";

    if (!isSpecial && nonBlank.length >= 3) {
      const isAllText = nonBlank.every((c) => isNaN(Number(String(c).replace(/^[₹$€£,%]\s?/g, ""))));
      if (isAllText && (!headerRow || nonBlank.length > headerRow.length)) {
        headerRow = r;
      }
    }
  }

  const headerStrings = (headerRow || []).slice(0, colCount).map((c) => String(c || ""));
  const alignments: AlignmentTypeValue[] = headerStrings.map((h, colIdx) => {
    const vals = tableSourceRows.map((r) => r[colIdx] || "");
    return detectColumnAlignment(h, vals);
  });

  const tableRows: TableRow[] = [];

  for (let rIdx = 0; rIdx < tableSourceRows.length; rIdx++) {
    const row = tableSourceRows[rIdx];
    const nonBlank = row.filter((c) => c !== "");
    const firstCell = nonBlank[0] || "";
    const firstLower = firstCell.toLowerCase();

    // 1. Empty spacer row
    if (nonBlank.length === 0) {
      const emptyCells: TableCell[] = [];
      for (let c = 0; c < colCount; c++) {
        emptyCells.push(createCell("", { widthDxa: colGrid ? colGrid[c] : undefined }));
      }
      tableRows.push(new TableRow({ cantSplit: true, children: emptyCells }));
      continue;
    }

    // 2. Banner or Section heading
    const isMainTitle =
      rIdx === 0 &&
      (firstLower.includes("report") ||
        firstLower.includes("statement") ||
        firstLower === "wint wealth");

    const isSubHeading =
      nonBlank.length === 1 &&
      (firstLower.includes("summary") ||
        firstLower.includes("statement") ||
        firstLower.includes("report") ||
        firstLower === "wint wealth");

    if (isMainTitle || isSubHeading) {
      const isLargeWint = firstLower === "wint wealth" && sec.type === "holding";
      const totalWidth = colGrid ? colGrid.reduce((a, b) => a + b, 0) : undefined;
      tableRows.push(
        new TableRow({
          cantSplit: true,
          children: [
            createCell(firstCell, {
              columnSpan: colCount,
              bold: true,
              align: AlignmentType.CENTER,
              size: isLargeWint ? 40 : 20,
              widthDxa: totalWidth,
            }),
          ],
        })
      );
      continue;
    }

    // 3. User metadata row (Name, Phone, Email)
    const isMeta =
      nonBlank.length >= 2 &&
      nonBlank.length <= 4 &&
      ["name", "phone number", "phone", "email", "pan"].includes(firstLower.replace(/:$/, ""));

    if (isMeta) {
      const label = firstCell.replace(/:$/, "").trim();
      const val = nonBlank[1] || "";
      const isHolding = sec.type === "holding";
      const cells: TableCell[] = [
        createCell(label, {
          bold: isHolding,
          size: 20,
          widthDxa: colGrid ? colGrid[0] : undefined,
        }),
        createCell(val, {
          size: 20,
          widthDxa: colGrid ? colGrid[1] : undefined,
        }),
      ];
      for (let c = 2; c < colCount; c++) {
        cells.push(createCell("", { size: 20, widthDxa: colGrid ? colGrid[c] : undefined }));
      }
      tableRows.push(new TableRow({ cantSplit: true, children: cells }));
      continue;
    }

    // 4. Date of report generation row
    if (firstLower.includes("date of report generation") || firstLower.includes("report generation date")) {
      const label = firstCell.replace(/:$/, "").trim() + ":";
      const val = nonBlank[1] || "";
      const isHolding = sec.type === "holding";
      const cells: TableCell[] = [
        createCell(label, {
          bold: isHolding,
          size: 20,
          widthDxa: colGrid ? colGrid[0] : undefined,
        }),
        createCell(val, {
          bold: isHolding,
          size: 20,
          widthDxa: colGrid ? colGrid[1] : undefined,
        }),
      ];
      for (let c = 2; c < colCount; c++) {
        cells.push(createCell("", { size: 20, widthDxa: colGrid ? colGrid[c] : undefined }));
      }
      tableRows.push(new TableRow({ cantSplit: true, children: cells }));
      continue;
    }

    // 5. Disclaimer row (for repayment, sell, cashflow, holding)
    if (
      firstLower.startsWith("disclaimer") ||
      firstLower.includes("the document and details") ||
      firstLower.includes("the information provided") ||
      firstLower.includes("before filing your income tax") ||
      firstLower.includes("we have classified") ||
      firstLower.includes("all information provided herein") ||
      firstLower.includes("tax laws and interpretations") ||
      firstLower.includes("please consult a qualified") ||
      firstLower.includes("wint wealth disclaims")
    ) {
      const fullText = nonBlank.join(" ").trim();
      const isDisclaimerTitle = firstLower === "disclaimer:" || firstLower === "disclaimer";
      const totalWidth = colGrid ? colGrid.reduce((a, b) => a + b, 0) : undefined;
      tableRows.push(
        new TableRow({
          cantSplit: true,
          children: [
            createCell(fullText, {
              columnSpan: colCount,
              bold: isDisclaimerTitle,
              size: 20,
              widthDxa: totalWidth,
            }),
          ],
        })
      );
      continue;
    }

    // 6. Header row
    if (row === headerRow) {
      const isHolding = sec.type === "holding";
      const cells: TableCell[] = [];
      for (let c = 0; c < colCount; c++) {
        cells.push(
          createCell(String(row[c] || ""), {
            bold: isHolding,
            size: 20,
            align: isHolding ? AlignmentType.CENTER : AlignmentType.LEFT,
            widthDxa: colGrid ? colGrid[c] : undefined,
          })
        );
      }
      tableRows.push(new TableRow({ tableHeader: true, cantSplit: true, children: cells }));
      continue;
    }

    // 7. Data row or Total row
    const isTotal = row.some((cell) => String(cell).toLowerCase().includes("total"));
    const isHolding = sec.type === "holding";
    const cells: TableCell[] = [];

    for (let c = 0; c < colCount; c++) {
      const val = String(row[c] !== undefined && row[c] !== null ? row[c] : "").trim();
      let align = alignments[c] || AlignmentType.LEFT;
      if (isHolding) {
        align = AlignmentType.CENTER;
      }
      cells.push(
        createCell(val, {
          bold: isTotal,
          size: 20,
          align,
          widthDxa: colGrid ? colGrid[c] : undefined,
        })
      );
    }
    tableRows.push(new TableRow({ cantSplit: true, children: cells }));
  }

  const totalTableWidth = colGrid ? colGrid.reduce((a, b) => a + b, 0) : undefined;

  const docxTable = new Table({
    width: totalTableWidth ? { size: totalTableWidth, type: WidthType.DXA } : { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: colGrid,
    borders: {
      top: TABLE_BORDER_STYLE,
      bottom: TABLE_BORDER_STYLE,
      left: TABLE_BORDER_STYLE,
      right: TABLE_BORDER_STYLE,
      insideHorizontal: TABLE_BORDER_STYLE,
      insideVertical: TABLE_BORDER_STYLE,
    },
    rows: tableRows,
  });

  let acquisitionCostTable: Table | undefined = undefined;
  if (acCostText || acDisclaimerText) {
    acquisitionCostTable = createAcquisitionCostTable(acCostText, acDisclaimerText);
  }

  return { table: docxTable, acquisitionCostTable };
}

/**
 * Converts an Excel buffer (ArrayBuffer or Uint8Array) into a fully styled Word (.docx) file
 * matching Master_Report_(1).docx with 100% precision.
 */
export async function convertExcelToDocx(
  data: ArrayBuffer | Uint8Array,
  options: ExcelToDocxOptions = {}
): Promise<Uint8Array> {
  const workbook = XLSX.read(data, { type: "array", cellDates: true });
  const allSheetNames = workbook.SheetNames;
  const sheetsToProcess =
    options.selectedSheets && options.selectedSheets.length > 0
      ? options.selectedSheets.filter((s) => allSheetNames.includes(s))
      : allSheetNames;

  if (sheetsToProcess.length === 0) {
    throw new Error("No readable sheets found in the Excel workbook.");
  }

  const logoArr = getLogoAsset();
  const barArr = getBarAsset();

  const header = createWintHeader(logoArr, barArr);
  const footer = createWintFooter(barArr);

  // Collect all raw rows across processed sheets
  const allSections: RawSection[] = [];

  for (let sIdx = 0; sIdx < sheetsToProcess.length; sIdx++) {
    const sheetName = sheetsToProcess[sIdx];
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;

    const rawData = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
    });

    const cleanRows = rawData.map((row) =>
      (Array.isArray(row) ? row : []).map((c) =>
        c !== null && c !== undefined ? String(c).trim() : ""
      )
    );

    const sheetSections = segmentRowsIntoSections(cleanRows);
    allSections.push(...sheetSections);
  }

  // Organize sections into exact pages matching Master_Report_(1).docx:
  // Page 1: Investment Report (Table 1) + Acquisition Cost & Disclaimer (Table 2)
  // Page 2: Repayment Report (Table 3) + Sell Report (Table 4)
  // Page 3: Upcoming Cashflow (Table 5) + FY Summary (Table 6) + Monthly Summary (Table 7)
  // Page 4: Bonds Holding Statement (Table 8) + Signing Block
  const bodyChildren: Array<Table | Paragraph> = [];

  // Group by page
  const page1Sections = allSections.filter((s) => s.type === "investment");
  const page2Sections = allSections.filter((s) => s.type === "repayment" || s.type === "sell");
  const page3Sections = allSections.filter(
    (s) => s.type === "cashflow" || s.type === "cashflow-fy" || s.type === "cashflow-monthly"
  );
  const page4Sections = allSections.filter((s) => s.type === "holding");
  const genericSections = allSections.filter(
    (s) =>
      s.type !== "investment" &&
      s.type !== "repayment" &&
      s.type !== "sell" &&
      s.type !== "cashflow" &&
      s.type !== "cashflow-fy" &&
      s.type !== "cashflow-monthly" &&
      s.type !== "holding"
  );

  let hasPageContent = false;

  // Helper to add a page of sections
  const addPageGroup = (secs: RawSection[], isLastPage = false) => {
    if (secs.length === 0) return;
    if (hasPageContent) {
      bodyChildren.push(new Paragraph({ children: [new PageBreak()] }));
    }

    for (let i = 0; i < secs.length; i++) {
      if (i > 0) {
        bodyChildren.push(new Paragraph({ text: "" }));
      }
      const { table, acquisitionCostTable } = renderSectionToTable(secs[i]);
      bodyChildren.push(table);
      if (acquisitionCostTable) {
        bodyChildren.push(new Paragraph({ text: "" }));
        bodyChildren.push(acquisitionCostTable);
      }
    }

    if (isLastPage) {
      bodyChildren.push(new Paragraph({ text: "" }));
      bodyChildren.push(...createSigningBlock());
    }

    hasPageContent = true;
  };

  // If standard Wint sections are present:
  if (
    page1Sections.length > 0 ||
    page2Sections.length > 0 ||
    page3Sections.length > 0 ||
    page4Sections.length > 0
  ) {
    if (page1Sections.length > 0) addPageGroup(page1Sections, false);
    if (page2Sections.length > 0) addPageGroup(page2Sections, false);
    if (page3Sections.length > 0) addPageGroup(page3Sections, false);
    if (page4Sections.length > 0) addPageGroup(page4Sections, true);
    if (genericSections.length > 0) addPageGroup(genericSections, false);
  } else {
    // Arbitrary workbook fallback
    for (let i = 0; i < allSections.length; i++) {
      if (i > 0) {
        bodyChildren.push(new Paragraph({ children: [new PageBreak()] }));
      }
      const { table } = renderSectionToTable(allSections[i]);
      bodyChildren.push(table);
    }
    bodyChildren.push(new Paragraph({ text: "" }));
    bodyChildren.push(...createSigningBlock());
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            font: "Helvetica",
            size: 20,
          },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: {
              width: 31678, // Landscape width matching Master_Report_(1).docx (code 7)
              height: 23814, // Landscape height
              orientation: PageOrientation.LANDSCAPE,
            },
            margin: {
              top: 1440,
              right: 2268,
              bottom: 1440,
              left: 1440,
              header: 0,
              footer: 0,
            },
          },
        },
        headers: {
          default: header,
        },
        footers: {
          default: footer,
        },
        children: bodyChildren,
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  return new Uint8Array(buffer);
}

/**
 * Universal helper to convert XLSX buffer to Word buffer.
 */
export async function convertExcelToDocxBuffer(
  excelData: ArrayBuffer | Uint8Array,
  options?: ExcelToDocxOptions
): Promise<Buffer> {
  const bytes = await convertExcelToDocx(excelData, options);
  return Buffer.from(bytes);
}
