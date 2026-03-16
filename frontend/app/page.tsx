"use client";

import { ChangeEvent, DragEvent, useRef, useState } from "react";

// ── DropZone component ────────────────────────────────────────────────────────

function DropZone({
  label,
  hint,
  accept,
  file,
  onFile,
  required,
}: {
  label: string;
  hint: string;
  accept: string;
  file: File | null;
  onFile: (f: File | null) => void;
  required?: boolean;
}) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);

  const handleDrop = (e: DragEvent) => {
    e.preventDefault();
    setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) onFile(f);
  };

  const handleChange = (e: ChangeEvent<HTMLInputElement>) => {
    onFile(e.target.files?.[0] ?? null);
  };

  return (
    <div
      className={[
        "relative border-2 border-dashed rounded-xl p-5 text-center cursor-pointer transition-all",
        dragging
          ? "border-blue-400 bg-blue-50 scale-[1.01]"
          : file
          ? "border-emerald-400 bg-emerald-50"
          : "border-gray-200 hover:border-blue-300 bg-gray-50 hover:bg-blue-50/40",
      ].join(" ")}
      onDragOver={(e) => {
        e.preventDefault();
        setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        className="hidden"
        onChange={handleChange}
      />

      {/* Icon */}
      <div className="flex justify-center mb-2">
        {file ? (
          <svg className="w-8 h-8 text-emerald-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
          </svg>
        ) : (
          <svg className="w-8 h-8 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5}
              d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" />
          </svg>
        )}
      </div>

      {/* Label */}
      <p className="text-sm font-semibold text-gray-700">
        {label}
        {required && <span className="text-red-400 ml-0.5">*</span>}
      </p>

      {/* File name or hint */}
      {file ? (
        <div className="mt-1 flex items-center justify-center gap-1.5">
          <span className="text-xs text-emerald-700 truncate max-w-[180px]">{file.name}</span>
          <button
            className="text-gray-400 hover:text-red-400 transition-colors"
            title="Remove file"
            onClick={(e) => {
              e.stopPropagation();
              onFile(null);
            }}
          >
            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
            </svg>
          </button>
        </div>
      ) : (
        <p className="mt-1 text-xs text-gray-400">{hint}</p>
      )}
    </div>
  );
}

// ── Spinner ───────────────────────────────────────────────────────────────────

function Spinner() {
  return (
    <svg className="animate-spin h-5 w-5" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
      <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8z" />
    </svg>
  );
}

// ── Main page ─────────────────────────────────────────────────────────────────

type Status = "idle" | "processing" | "done" | "error";

export default function Home() {
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  const [txFile, setTxFile] = useState<File | null>(null);
  const [debitSpread, setDebitSpread] = useState("0.35");
  const [creditSpread, setCreditSpread] = useState("-0.60");
  const [status, setStatus] = useState<Status>("idle");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [downloadName, setDownloadName] = useState("bank_interest_output.xlsx");
  const [errorMsg, setErrorMsg] = useState("");

  const canRun = pdfFile !== null && status !== "processing";

  const handleRun = async () => {
    if (!pdfFile) return;

    // Revoke previous blob URL if any
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);

    setStatus("processing");
    setDownloadUrl(null);
    setErrorMsg("");

    const form = new FormData();
    form.append("pdf_file", pdfFile);
    if (txFile) form.append("tx_file", txFile);
    form.append("debit_spread", debitSpread);
    form.append("credit_spread", creditSpread);

    try {
      const apiUrl = process.env.NEXT_PUBLIC_API_URL ?? "http://localhost:8000";
      const res = await fetch(`${apiUrl}/process`, { method: "POST", body: form });

      if (!res.ok) {
        let detail = `Server returned ${res.status}`;
        try {
          const json = await res.json();
          detail = json.detail ?? detail;
        } catch {
          // ignore JSON parse failure
        }
        throw new Error(detail);
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);

      // Extract filename from Content-Disposition header if present
      const disposition = res.headers.get("content-disposition") ?? "";
      const match = disposition.match(/filename[^;=\n]*=["']?([^"';\n]+)["']?/);
      setDownloadUrl(url);
      setDownloadName(match?.[1] ?? "bank_interest_output.xlsx");
      setStatus("done");
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : String(err));
      setStatus("error");
    }
  };

  return (
    <main className="min-h-screen bg-gradient-to-br from-slate-100 to-blue-50 flex items-center justify-center p-4">
      <div className="w-full max-w-xl">
        {/* Card */}
        <div className="bg-white rounded-2xl shadow-xl border border-gray-100 overflow-hidden">
          {/* Header */}
          <div className="bg-gradient-to-r from-blue-600 to-blue-500 px-8 py-6 text-white">
            <h1 className="text-xl font-bold tracking-tight">Bank Hapoalim — Interest Analyzer</h1>
            <p className="mt-1 text-sm text-blue-100">
              Upload your interest-charge PDF to generate the Excel reconciliation report.
            </p>
          </div>

          <div className="px-8 py-6 space-y-6">
            {/* File uploads */}
            <div className="space-y-3">
              <h2 className="text-xs font-semibold text-gray-500 uppercase tracking-wider">Files</h2>
              <DropZone
                label="Interest Charge PDF"
                hint="Drag & drop or click to browse (.pdf)"
                accept=".pdf"
                file={pdfFile}
                onFile={setPdfFile}
                required
              />
              <DropZone
                label="Transactions File"
                hint="Optional — enables Sheets 2 & 3 (.xlsx, .xls, .csv)"
                accept=".xlsx,.xls,.csv"
                file={txFile}
                onFile={setTxFile}
              />
            </div>

            {/* Spread inputs */}
            <div className="space-y-3">
              <h2 className="text-xs font-semibold text-gray-500 uppercase tracking-wider">Parameters</h2>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-1">
                    Debit Spread (%)
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    value={debitSpread}
                    onChange={(e) => setDebitSpread(e.target.value)}
                    className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400 focus:border-transparent"
                  />
                </div>
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-1">
                    Credit Spread (%)
                  </label>
                  <input
                    type="number"
                    step="0.01"
                    value={creditSpread}
                    onChange={(e) => setCreditSpread(e.target.value)}
                    className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-400 focus:border-transparent"
                  />
                </div>
              </div>
            </div>

            {/* Run button */}
            <button
              onClick={handleRun}
              disabled={!canRun}
              className="w-full py-3 rounded-xl font-semibold text-white text-sm transition-all
                         bg-blue-600 hover:bg-blue-700 active:scale-[0.98]
                         disabled:opacity-40 disabled:cursor-not-allowed disabled:active:scale-100
                         flex items-center justify-center gap-2"
            >
              {status === "processing" ? (
                <>
                  <Spinner />
                  Processing…
                </>
              ) : (
                <>
                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                      d="M14.752 11.168l-3.197-2.132A1 1 0 0010 9.87v4.263a1 1 0 001.555.832l3.197-2.132a1 1 0 000-1.664z" />
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                      d="M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                  Run
                </>
              )}
            </button>

            {/* Status messages */}
            {status === "processing" && (
              <div className="flex gap-3 p-4 bg-blue-50 rounded-xl text-blue-700 text-sm border border-blue-100">
                <Spinner />
                <span>
                  Parsing PDF and fetching Bank of Israel rates…
                  <br />
                  <span className="text-xs text-blue-500">
                    First request may take 30–60 s while the server warms up.
                  </span>
                </span>
              </div>
            )}

            {status === "done" && downloadUrl && (
              <div className="flex items-center justify-between p-4 bg-emerald-50 rounded-xl border border-emerald-100">
                <div className="flex items-center gap-2 text-emerald-700 text-sm">
                  <svg className="w-5 h-5 text-emerald-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                      d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                  </svg>
                  <span className="font-medium">Excel report ready</span>
                </div>
                <a
                  href={downloadUrl}
                  download={downloadName}
                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 active:scale-95
                             text-white text-sm font-medium rounded-lg transition-all
                             flex items-center gap-1.5"
                >
                  <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                      d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
                  </svg>
                  Download
                </a>
              </div>
            )}

            {status === "error" && (
              <div className="p-4 bg-red-50 rounded-xl border border-red-100 text-red-700 text-sm">
                <p className="font-semibold mb-0.5">Error</p>
                <p className="text-red-600">{errorMsg}</p>
              </div>
            )}
          </div>
        </div>

        <p className="mt-4 text-center text-xs text-gray-400">
          All processing happens server-side. Files are not stored.
        </p>
      </div>
    </main>
  );
}
