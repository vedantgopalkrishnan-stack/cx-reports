import * as XLSX from "xlsx";

export interface ParsedSheetContent {
  bannerTitle?: string;
  sectionTitle?: string;
  metadata: Array<{ label: string; value: string }>;
  headers: string[];
  bodyRows: string[][];
  disclaimers: string[];
  columnAlignments: Array<"left" | "center" | "right">;
}

export interface ParsedDocxTable {
  sectionTitle?: string;
  headers: string[];
  bodyRows: string[][];
  alignments: Array<"left" | "center" | "right">;
}

export interface ParsedDocxSheet {
  sheetName: string;
  bannerTitle?: string;
  metadata: Array<{ label: string; value: string }>;
  dateOfGeneration?: string;
  tables: ParsedDocxTable[];
  disclaimers: string[];
  hasSigningBlock: boolean;
}

/**
 * Checks if a string looks like a disclaimer, legal note, or terms block.
 */
export function isDisclaimerOrNote(text: string): boolean {
  const trimmed = text.trim();
  const lower = trimmed.toLowerCase();
  if (
    lower.startsWith("disclaimer") ||
    lower.startsWith("*acquisition cost") ||
    lower.startsWith("acquisition cost")
  ) {
    return true;
  }
  if (
    lower.startsWith("note:") ||
    lower.startsWith("notes:") ||
    lower.startsWith("important note")
  ) {
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
export function isMetadataKey(text: string): boolean {
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
    "financial year",
  ];
  return keys.includes(lower);
}

/**
 * Detects whether a column primarily contains numbers, currency, dates, or text.
 */
export function detectColumnAlignment(header: string, values: string[]): "left" | "center" | "right" {
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
    hLower.includes("type") ||
    hLower === "transactiontype"
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
    hLower.includes("price") ||
    hLower.includes("invested")
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
 * Standard parser used by excel-to-pdf for single-table sheets.
 */
export function parseSheetSections(rawData: unknown[][]): ParsedSheetContent | null {
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
  let lastNonEmpty = rawHeaders.length - 1;
  while (lastNonEmpty >= 0 && rawHeaders[lastNonEmpty] === "") {
    lastNonEmpty--;
  }
  const colCount = Math.max(1, lastNonEmpty + 1);

  const headers = rawHeaders.slice(0, colCount).map((h, i) => h || `Col ${i + 1}`);

  const rawBodyRows = candidateTableRows.slice(headerIdx + 1);
  const bodyRows: string[][] = [];

  for (const r of rawBodyRows) {
    const rowSlice = r.slice(0, colCount);
    while (rowSlice.length < colCount) {
      rowSlice.push("");
    }
    if (rowSlice.some((c) => c !== "") && !isDisclaimerOrNote(rowSlice[0])) {
      bodyRows.push(rowSlice);
    }
  }

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
 * Enhanced parser for Word (.docx) export that accurately segments sheets into:
 * - Banner titles (e.g. "CONSOLIDATED INVESTMENT REPORT WINT WEALTH")
 * - User metadata (Name, Phone, Email, PAN)
 * - Multiple sub-tables if present (e.g. multiple FY summaries, cashflow tables)
 * - Date of generation
 * - Disclaimers & acquisition cost notes
 * - Signing block detection
 */
export function parseSheetForDocx(
  rawData: unknown[][],
  sheetName: string,
  isLastSheet = false
): ParsedDocxSheet | null {
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
  let dateOfGeneration: string | undefined = undefined;

  const candidateRows: string[][] = [];

  for (const row of cleanRows) {
    const nonBlank = row.filter((c) => c !== "");
    const firstCell = nonBlank[0] || "";
    const firstCellLower = firstCell.toLowerCase();

    // 1. Disclaimer / Notes
    if (isDisclaimerOrNote(firstCell) || (nonBlank.length <= 2 && row.join(" ").length > 110)) {
      const fullText = nonBlank.join(" ").trim();
      if (fullText && !disclaimers.includes(fullText)) {
        disclaimers.push(fullText);
      }
      continue;
    }

    // 2. Date of report generation row
    if (firstCellLower.includes("date of report generation") || firstCellLower.includes("report generation date")) {
      dateOfGeneration = nonBlank[1] || "";
      continue;
    }

    // 3. User metadata
    if (nonBlank.length >= 2 && nonBlank.length <= 4 && isMetadataKey(nonBlank[0])) {
      const label = nonBlank[0].replace(/:$/, "").trim();
      const value = nonBlank[1]?.trim() || "";
      if (label && value && !metadata.some((m) => m.label.toLowerCase() === label.toLowerCase())) {
        metadata.push({ label, value });
      }
      continue;
    }

    // 4. Banner title (before tables)
    if (
      candidateRows.length === 0 &&
      nonBlank.length === 1 &&
      (firstCellLower.includes("report") ||
        firstCellLower.includes("wint") ||
        firstCellLower.includes("consolidated") ||
        firstCellLower.includes("statement"))
    ) {
      bannerTitle = firstCell;
      continue;
    }

    candidateRows.push(row);
  }

  // Parse candidate rows into one or more sub-tables
  const tables: ParsedDocxTable[] = [];

  let currentSectionTitle: string | undefined = undefined;
  let currentHeaders: string[] | null = null;
  let currentBodyRows: string[][] = [];

  function flushCurrentTable() {
    if (!currentHeaders || currentHeaders.length === 0) {
      // If we only have body rows without clear headers, create fallback headers
      if (currentBodyRows.length > 0) {
        const colCount = Math.max(...currentBodyRows.map((r) => r.length), 1);
        const headers = Array.from({ length: colCount }, (_, i) => `Column ${i + 1}`);
        const alignments = headers.map(() => "left" as const);
        tables.push({
          sectionTitle: currentSectionTitle,
          headers,
          bodyRows: currentBodyRows,
          alignments,
        });
      }
    } else {
      const colCount = currentHeaders.length;
      const normalizedRows = currentBodyRows.map((r) => {
        const slice = r.slice(0, colCount);
        while (slice.length < colCount) slice.push("");
        return slice;
      });

      const alignments: Array<"left" | "center" | "right"> = currentHeaders.map((h, colIdx) => {
        const colValues = normalizedRows.map((r) => r[colIdx] || "");
        return detectColumnAlignment(h, colValues);
      });

      tables.push({
        sectionTitle: currentSectionTitle,
        headers: currentHeaders,
        bodyRows: normalizedRows,
        alignments,
      });
    }

    currentHeaders = null;
    currentBodyRows = [];
  }

  for (let i = 0; i < candidateRows.length; i++) {
    const row = candidateRows[i];
    const nonBlank = row.filter((c) => c !== "");
    const firstCell = nonBlank[0] || "";
    const firstCellLower = firstCell.toLowerCase();

    // Check if this row is a section title (single prominent label)
    const isSectionHeading =
      nonBlank.length === 1 &&
      (firstCellLower.includes("summary") ||
        firstCellLower.includes("statement") ||
        firstCellLower.includes("report") ||
        firstCellLower.includes("holding") ||
        firstCellLower.includes("fy") ||
        firstCell.toUpperCase() === firstCell);

    if (isSectionHeading) {
      if (currentHeaders !== null || currentBodyRows.length > 0) {
        flushCurrentTable();
      }
      currentSectionTitle = firstCell;
      continue;
    }

    // Check if this row is a header row:
    // Header rows have >= 3 non-empty cells, or non-numeric cells that don't look like data
    const hasEnoughCols = nonBlank.length >= 2;
    const isAllText = nonBlank.every((c) => isNaN(Number(c.replace(/^[₹$€£,%]\s?/g, ""))));
    const looksLikeHeader =
      hasEnoughCols &&
      isAllText &&
      nonBlank.some((c) => {
        const cl = c.toLowerCase();
        return (
          cl.includes("name") ||
          cl.includes("date") ||
          cl.includes("isin") ||
          cl.includes("units") ||
          cl.includes("amount") ||
          cl.includes("principal") ||
          cl.includes("interest") ||
          cl.includes("tds") ||
          cl.includes("value") ||
          cl.includes("rate") ||
          cl.includes("year") ||
          cl.includes("month") ||
          cl.includes("cost")
        );
      });

    if (looksLikeHeader && (currentHeaders === null || currentBodyRows.length > 0)) {
      if (currentHeaders !== null && currentBodyRows.length > 0) {
        flushCurrentTable();
      }
      // Trim empty trailing cells
      let lastNonEmpty = row.length - 1;
      while (lastNonEmpty >= 0 && row[lastNonEmpty] === "") {
        lastNonEmpty--;
      }
      const colCount = Math.max(1, lastNonEmpty + 1);
      currentHeaders = row.slice(0, colCount).map((h, idx) => h || `Col ${idx + 1}`);
      continue;
    }

    if (currentHeaders === null) {
      // If we haven't found a header yet, maybe this row is the header
      let lastNonEmpty = row.length - 1;
      while (lastNonEmpty >= 0 && row[lastNonEmpty] === "") {
        lastNonEmpty--;
      }
      const colCount = Math.max(1, lastNonEmpty + 1);
      currentHeaders = row.slice(0, colCount).map((h, idx) => h || `Col ${idx + 1}`);
      continue;
    }

    // Otherwise, this is a body row (or total row)
    currentBodyRows.push(row);
  }

  // Flush remaining table
  if (currentHeaders !== null || currentBodyRows.length > 0) {
    flushCurrentTable();
  }

  // If no tables could be extracted, build a simple single-column fallback
  if (tables.length === 0) {
    tables.push({
      headers: ["Information"],
      bodyRows: cleanRows.map((r) => [r.join(" ")]),
      alignments: ["left"],
    });
  }

  const sheetLower = sheetName.toLowerCase();
  const hasSigningBlock =
    isLastSheet ||
    sheetLower.includes("holding") ||
    sheetLower.includes("bond") ||
    sheetLower.includes("tax");

  return {
    sheetName,
    bannerTitle,
    metadata,
    dateOfGeneration,
    tables,
    disclaimers,
    hasSigningBlock,
  };
}

/**
 * Parses all selected sheets from an Excel workbook buffer for Word (.docx) generation.
 */
export function parseWorkbookForDocx(
  data: ArrayBuffer | Uint8Array,
  selectedSheets?: string[]
): ParsedDocxSheet[] {
  const workbook = XLSX.read(data, { type: "array", cellDates: true });
  const allSheetNames = workbook.SheetNames;
  const sheetsToProcess =
    selectedSheets && selectedSheets.length > 0
      ? selectedSheets.filter((s) => allSheetNames.includes(s))
      : allSheetNames;

  const result: ParsedDocxSheet[] = [];

  for (let i = 0; i < sheetsToProcess.length; i++) {
    const name = sheetsToProcess[i];
    const sheet = workbook.Sheets[name];
    if (!sheet) continue;

    const rawData = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
      blankrows: false,
    });

    const isLast = i === sheetsToProcess.length - 1;
    const parsed = parseSheetForDocx(rawData, name, isLast);
    if (parsed) {
      result.push(parsed);
    }
  }

  return result;
}
