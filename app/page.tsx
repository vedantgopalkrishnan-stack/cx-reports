"use client";

import { useEffect, useState } from "react";
import Header from "@/components/Header";
import ReportForm from "@/components/ReportForm";
import ResultPanel from "@/components/ResultPanel";
import HistoryList from "@/components/HistoryList";
import ExcelToPdfConverter from "@/components/ExcelToPdfConverter";
import Toast from "@/components/Toast";
import { useToast } from "@/lib/useToast";
import { getHistory, HistoryEntry } from "@/lib/history";

export default function Home() {
  const [history, setHistory] = useState<HistoryEntry[]>([]);
  const [lastResult, setLastResult] = useState<HistoryEntry | null>(null);
  const [activeTab, setActiveTab] = useState<"converter" | "downloader">("converter");
  const { toastMessage, showToast } = useToast();

  // Read on mount rather than during render: localStorage does not exist while
  // the page is being server-rendered.
  useEffect(() => {
    setHistory(getHistory());
  }, []);

  function handleResult(entry: HistoryEntry) {
    setLastResult(entry);
    setHistory(getHistory());
  }

  return (
    <div className="min-h-screen bg-cream">
      <Header />
      <main className="mx-auto max-w-2xl px-6 py-12 sm:px-10">
        {/* Navigation Tabs */}
        <div className="mb-6 flex rounded-full border border-ink/10 bg-white/70 p-1 backdrop-blur-sm">
          <button
            type="button"
            onClick={() => setActiveTab("converter")}
            className={`flex-1 rounded-full py-2.5 text-xs font-bold transition-all sm:text-sm ${
              activeTab === "converter"
                ? "bg-ink text-white shadow-sm"
                : "text-ink/60 hover:text-ink"
            }`}
          >
            Excel to PDF Converter
            <span className="ml-2 rounded-full bg-accent px-2 py-0.5 text-[10px] text-white font-normal uppercase">
              New
            </span>
          </button>
          <button
            type="button"
            onClick={() => setActiveTab("downloader")}
            className={`flex-1 rounded-full py-2.5 text-xs font-bold transition-all sm:text-sm ${
              activeTab === "downloader"
                ? "bg-ink text-white shadow-sm"
                : "text-ink/60 hover:text-ink"
            }`}
          >
            User Report Downloader
          </button>
        </div>

        {activeTab === "converter" ? (
          <div>
            <span className="badge mb-4">Feature Preview</span>
            <h1 className="mb-2 text-4xl font-black tracking-tight">Excel to PDF Converter</h1>
            <p className="mb-8 text-sm text-ink/60">
              Upload any Excel file (.xlsx, .xls, .csv) or test with our sample report to instantly preview and convert it to a beautifully formatted PDF.
            </p>
            <ExcelToPdfConverter onToast={showToast} />
          </div>
        ) : (
          <div>
            <span className="badge mb-4">Report Library</span>
            <h1 className="mb-2 text-4xl font-black tracking-tight">Wint User Report Download</h1>
            <p className="mb-8 text-sm text-ink/60">
              Enter a user ID and pull their Master or Taxation report as an Excel file or converted PDF.
            </p>
            <ReportForm onResult={handleResult} onToast={showToast} />
            <ResultPanel entry={lastResult} onToast={showToast} />
            <HistoryList entries={history} onToast={showToast} />
          </div>
        )}
      </main>
      <Toast message={toastMessage} />
    </div>
  );
}
