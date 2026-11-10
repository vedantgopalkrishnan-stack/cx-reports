"use client";

import { useState, useRef, ChangeEvent, DragEvent } from "react";
import {
  convertExcelToPdf,
  inspectExcelWorkbook,
  generateSampleExcelWorkbook,
  WorkbookInfo,
  ExcelToPdfOptions,
} from "@/lib/excel-to-pdf";

interface Props {
  onToast: (message: string) => void;
}

export default function ExcelToPdfConverter({ onToast }: Props) {
  const [file, setFile] = useState<File | null>(null);
  const [fileData, setFileData] = useState<Uint8Array | null>(null);
  const [workbookInfo, setWorkbookInfo] = useState<WorkbookInfo | null>(null);
  const [activeSheetIndex, setActiveSheetIndex] = useState<number>(0);
  const [orientation, setOrientation] = useState<"auto" | "portrait" | "landscape">("auto");
  const [title, setTitle] = useState<string>("");
  const [isConverting, setIsConverting] = useState<boolean>(false);
  const [isDragging, setIsDragging] = useState<boolean>(false);
  const [pdfPreviewUrl, setPdfPreviewUrl] = useState<string | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  function processArrayBuffer(buf: ArrayBuffer, fileName: string) {
    try {
      const uint8 = new Uint8Array(buf);
      const info = inspectExcelWorkbook(uint8);
      setFileData(uint8);
      setWorkbookInfo(info);
      setActiveSheetIndex(0);
      setTitle(fileName.replace(/\.(xlsx|xls|csv)$/i, ""));
      setPdfPreviewUrl(null);
      onToast(`Loaded ${fileName} (${info.sheetNames.length} sheet${info.sheetNames.length > 1 ? "s" : ""})`);
    } catch (err) {
      onToast("Could not parse Excel file. Please ensure it is a valid .xlsx or .xls file.");
    }
  }

  function handleFileChange(e: ChangeEvent<HTMLInputElement>) {
    const selected = e.target.files?.[0];
    if (!selected) return;
    setFile(selected);
    const reader = new FileReader();
    reader.onload = (event) => {
      if (event.target?.result instanceof ArrayBuffer) {
        processArrayBuffer(event.target.result, selected.name);
      }
    };
    reader.readAsArrayBuffer(selected);
  }

  function handleDragOver(e: DragEvent<HTMLDivElement>) {
    e.preventDefault();
    setIsDragging(true);
  }

  function handleDragLeave(e: DragEvent<HTMLDivElement>) {
    e.preventDefault();
    setIsDragging(false);
  }

  function handleDrop(e: DragEvent<HTMLDivElement>) {
    e.preventDefault();
    setIsDragging(false);
    const droppedFile = e.dataTransfer.files?.[0];
    if (!droppedFile) return;

    if (!droppedFile.name.match(/\.(xlsx|xls|csv)$/i)) {
      onToast("Please drop an Excel file (.xlsx, .xls, .csv)");
      return;
    }

    setFile(droppedFile);
    const reader = new FileReader();
    reader.onload = (event) => {
      if (event.target?.result instanceof ArrayBuffer) {
        processArrayBuffer(event.target.result, droppedFile.name);
      }
    };
    reader.readAsArrayBuffer(droppedFile);
  }

  function handleLoadSample() {
    const sampleBytes = generateSampleExcelWorkbook();
    setFile(null);
    processArrayBuffer(sampleBytes.buffer as ArrayBuffer, "Wint_Investment_Report_Sample.xlsx");
  }

  function handleConvert(previewOnly = false) {
    if (!fileData) {
      onToast("Please upload an Excel file first.");
      return;
    }

    setIsConverting(true);
    onToast("Generating PDF…");

    // Allow UI to tick before heavy work
    setTimeout(() => {
      try {
        const options: ExcelToPdfOptions = {
          orientation,
          title: title.trim() || "Excel Report",
          subtitle: `Generated on ${new Date().toLocaleDateString(undefined, { dateStyle: "medium" })}`,
        };

        const { pdfBytes, pageCount } = convertExcelToPdf(fileData, options);
        // Create Blob
        const blob = new Blob([new Uint8Array(pdfBytes) as unknown as BlobPart], { type: "application/pdf" });
        const url = URL.createObjectURL(blob);
        setPdfPreviewUrl(url);

        const pdfName = (file?.name || title || "converted").replace(/\.(xlsx|xls|csv)$/i, "") + ".pdf";

        if (previewOnly) {
          window.open(url, "_blank");
          onToast(`PDF opened in new tab (${pageCount} page${pageCount > 1 ? "s" : ""})`);
        } else {
          const a = document.createElement("a");
          a.href = url;
          a.download = pdfName;
          document.body.appendChild(a);
          a.click();
          a.remove();
          onToast(`Downloaded ${pdfName} (${pageCount} page${pageCount > 1 ? "s" : ""})`);
        }
      } catch (err) {
        onToast("PDF conversion failed: " + (err instanceof Error ? err.message : String(err)));
      } finally {
        setIsConverting(false);
      }
    }, 100);
  }

  function handleReset() {
    setFile(null);
    setFileData(null);
    setWorkbookInfo(null);
    setPdfPreviewUrl(null);
    if (fileInputRef.current) fileInputRef.current.value = "";
  }

  const activeSheet = workbookInfo?.sheets[activeSheetIndex];

  return (
    <div className="card p-6 sm:p-8">
      {/* Upload area */}
      {!workbookInfo ? (
        <div>
          <div
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className={`flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed p-8 text-center transition-colors ${
              isDragging ? "border-accent bg-accent-light/30" : "border-ink/20 hover:border-ink/40 bg-white"
            }`}
          >
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileChange}
              className="hidden"
            />
            <div className="mb-3 flex h-14 w-14 items-center justify-center rounded-full bg-accent-light text-accent-dark">
              <svg className="h-7 w-7" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
              </svg>
            </div>
            <p className="mb-1 text-base font-semibold text-ink">
              Choose an Excel file or drag & drop here
            </p>
            <p className="text-xs text-ink/50">Supports .xlsx, .xls, and .csv files</p>
          </div>

          <div className="mt-4 flex items-center justify-center gap-2">
            <span className="text-xs text-ink/40">Don't have an Excel file on hand?</span>
            <button
              type="button"
              onClick={handleLoadSample}
              className="text-xs font-semibold text-accent-dark hover:underline"
            >
              Load Sample Investment Report →
            </button>
          </div>
        </div>
      ) : (
        /* File Loaded & Configuration View */
        <div className="space-y-6">
          {/* File bar */}
          <div className="flex flex-col gap-2 rounded-xl bg-ink/5 p-4 sm:flex-row sm:items-center sm:justify-between">
            <div className="flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-lg bg-accent text-white">
                <svg className="h-5 w-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
              </div>
              <div>
                <p className="text-sm font-bold text-ink">
                  {file?.name || "Sample_Wint_Report.xlsx"}
                </p>
                <p className="text-xs text-ink/60">
                  {workbookInfo.sheetNames.length} sheet{workbookInfo.sheetNames.length > 1 ? "s" : ""} detected
                </p>
              </div>
            </div>
            <button
              type="button"
              onClick={handleReset}
              className="text-xs font-semibold text-red-600 hover:text-red-700"
            >
              Choose different file
            </button>
          </div>

          {/* Options */}
          <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
            <div>
              <label className="mb-1.5 block text-xs font-semibold text-ink/70">
                PDF Document Title
              </label>
              <input
                type="text"
                value={title}
                onChange={(e) => setTitle(e.target.value)}
                placeholder="Title to appear on PDF"
                className="w-full rounded-xl border border-ink/15 bg-white px-3.5 py-2 text-sm outline-none focus:border-ink/40"
              />
            </div>
            <div>
              <label className="mb-1.5 block text-xs font-semibold text-ink/70">
                Page Orientation
              </label>
              <div className="flex gap-2">
                {(["auto", "portrait", "landscape"] as const).map((mode) => (
                  <button
                    key={mode}
                    type="button"
                    onClick={() => setOrientation(mode)}
                    className={`flex-1 rounded-xl py-2 text-xs font-semibold capitalize transition-colors ${
                      orientation === mode
                        ? "bg-ink text-white"
                        : "border border-ink/15 bg-white text-ink/70 hover:bg-ink/5"
                    }`}
                  >
                    {mode}
                  </button>
                ))}
              </div>
            </div>
          </div>

          {/* Sheet tabs & preview */}
          <div>
            <div className="mb-2 flex items-center justify-between">
              <span className="text-xs font-semibold uppercase tracking-wider text-ink/50">
                Spreadsheet Preview
              </span>
              <span className="text-xs text-ink/50">
                {activeSheet ? `${activeSheet.totalRows} rows · ${activeSheet.totalCols} columns` : ""}
              </span>
            </div>

            {/* Tabs */}
            {workbookInfo.sheetNames.length > 1 && (
              <div className="mb-3 flex flex-wrap gap-1.5 border-b border-ink/10 pb-2">
                {workbookInfo.sheetNames.map((name, idx) => (
                  <button
                    key={name}
                    type="button"
                    onClick={() => setActiveSheetIndex(idx)}
                    className={`rounded-lg px-3 py-1.5 text-xs font-semibold transition-colors ${
                      activeSheetIndex === idx
                        ? "bg-accent text-white"
                        : "bg-ink/5 text-ink/70 hover:bg-ink/10"
                    }`}
                  >
                    {name}
                  </button>
                ))}
              </div>
            )}

            {/* Table Preview */}
            {activeSheet && (
              <div className="max-h-60 overflow-x-auto overflow-y-auto rounded-xl border border-ink/10 bg-cream/30 text-xs">
                <table className="w-full border-collapse text-left">
                  {activeSheet.headers.length > 0 && (
                    <thead className="sticky top-0 bg-ink text-white">
                      <tr>
                        {activeSheet.headers.map((h, i) => (
                          <th key={i} className="whitespace-nowrap px-3 py-2 font-medium">
                            {h || `Column ${i + 1}`}
                          </th>
                        ))}
                      </tr>
                    </thead>
                  )}
                  <tbody>
                    {activeSheet.previewRows.map((row, rIdx) => (
                      <tr key={rIdx} className="border-b border-ink/5 hover:bg-ink/5">
                        {row.map((cell, cIdx) => (
                          <td key={cIdx} className="whitespace-nowrap px-3 py-2 text-ink/80">
                            {cell !== null && cell !== undefined ? String(cell) : ""}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            <p className="mt-1.5 text-right text-[11px] text-ink/40">
              Showing first 10 rows · Full sheet will be included in PDF
            </p>
          </div>

          {/* Action buttons */}
          <div className="flex flex-col gap-3 sm:flex-row">
            <button
              type="button"
              disabled={isConverting}
              onClick={() => handleConvert(false)}
              className="btn-accent flex-1"
            >
              <svg className="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
              </svg>
              {isConverting ? "Converting…" : "Convert & Download PDF"}
            </button>

            <button
              type="button"
              disabled={isConverting}
              onClick={() => handleConvert(true)}
              className="btn-ghost flex-1 sm:flex-initial"
            >
              <svg className="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
              </svg>
              Preview PDF
            </button>
          </div>

          {pdfPreviewUrl && (
            <div className="rounded-xl border border-accent-light bg-accent-light/30 p-3 text-xs text-accent-dark flex items-center justify-between">
              <span>PDF is ready! Click "Convert & Download" anytime or "Preview PDF" to view it.</span>
              <a
                href={pdfPreviewUrl}
                target="_blank"
                rel="noreferrer"
                className="font-bold underline ml-2"
              >
                Open preview tab →
              </a>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
