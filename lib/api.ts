import type { ReportType } from "./reports";

export type { ReportType };

export interface GenerateReportResponse {
  success: boolean;
  downloadUrl?: string;
  message: string;
  accountDeletionSuspected?: boolean;
}

export class ApiError extends Error {}

export async function generateReport(reportType: ReportType, userId: string): Promise<GenerateReportResponse> {
  let res: Response;
  try {
    res = await fetch("/api/reports/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ reportType, userId }),
    });
  } catch {
    throw new ApiError("Could not reach the server. Check your connection and try again.");
  }

  // The API always answers with JSON. Anything else means the request never
  // reached the route handler, so say that rather than throwing a parse error.
  const text = await res.text();
  try {
    return JSON.parse(text) as GenerateReportResponse;
  } catch {
    throw new ApiError(`Unexpected response from the server (HTTP ${res.status}).`);
  }
}

const REPORT_SLUG: Record<ReportType, string> = {
  master: "master-report",
  taxation: "taxation-report",
};

export function reportFilename(reportType: ReportType, userId: string, format: "xlsx" | "pdf" = "xlsx"): string {
  return `${REPORT_SLUG[reportType]}-${userId}-${new Date().toISOString().slice(0, 10)}.${format}`;
}

// Route the file through our own origin. A cross-origin link ignores the
// `download` attribute and just navigates, which would throw the user out of
// the app; same-origin plus an attachment header saves the file in place.
export function downloadHref(url: string, filename: string): string {
  return `/api/download?${new URLSearchParams({ url, filename }).toString()}`;
}

export function pdfDownloadHref(url: string, filename: string): string {
  const pdfName = filename.replace(/\.xlsx$/i, ".pdf");
  return `/api/convert-pdf?${new URLSearchParams({ url, filename: pdfName }).toString()}`;
}

export function triggerDownload(url: string, filename: string): void {
  const anchor = document.createElement("a");
  anchor.href = downloadHref(url, filename);
  anchor.download = filename;
  anchor.rel = "noopener";
  anchor.style.display = "none";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
}

export function triggerPdfDownload(url: string, filename: string): void {
  const pdfName = filename.replace(/\.xlsx$/i, ".pdf");
  const anchor = document.createElement("a");
  anchor.href = pdfDownloadHref(url, pdfName);
  anchor.download = pdfName;
  anchor.rel = "noopener";
  anchor.style.display = "none";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
}
