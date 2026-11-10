import { downloadHref, pdfDownloadHref, reportFilename } from "@/lib/api";
import { HistoryEntry } from "@/lib/history";

interface Props {
  entry: HistoryEntry | null;
  onToast: (message: string) => void;
}

export default function ResultPanel({ entry, onToast }: Props) {
  if (!entry) return null;

  if (entry.status === "success") {
    const filename = entry.filename ?? reportFilename(entry.reportType, entry.userId);
    const pdfFilename = filename.replace(/\.xlsx$/i, ".pdf");

    return (
      <div className="card mt-4 flex flex-col items-start gap-4 border-accent-light bg-accent-light/40 p-5 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <span className="badge mb-1">Ready</span>
          <p className="text-sm font-medium text-ink">{entry.message}</p>
        </div>
        {entry.downloadUrl && (
          <div className="flex flex-wrap items-center gap-2">
            <a
              href={downloadHref(entry.downloadUrl, filename)}
              download={filename}
              onClick={() => onToast("Excel download started")}
              className="btn-accent text-xs"
            >
              <svg className="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
              </svg>
              Excel (.xlsx)
            </a>
            <a
              href={pdfDownloadHref(entry.downloadUrl, filename)}
              download={pdfFilename}
              onClick={() => onToast("Converting to PDF…")}
              className="btn-primary text-xs"
            >
              <svg className="h-4 w-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 21h10a2 2 0 002-2V9.414a1 1 0 00-.293-.707l-5.414-5.414A1 1 0 0012.586 3H7a2 2 0 00-2 2v14a2 2 0 002 2z" />
              </svg>
              PDF (.pdf)
            </a>
          </div>
        )}
      </div>
    );
  }

  if (entry.status === "account_deletion") {
    return (
      <div className="card mt-4 border-amber-200 bg-amber-50 p-5">
        <p className="text-sm font-semibold text-amber-800">Heads up</p>
        <p className="text-sm text-amber-900">{entry.message}</p>
      </div>
    );
  }

  return (
    <div className="card mt-4 border-red-200 bg-red-50 p-5">
      <p className="text-sm font-semibold text-red-800">Request failed</p>
      <p className="text-sm text-red-900">{entry.message}</p>
    </div>
  );
}
