"use client";

import { useCallback, useState } from "react";
import Image from "next/image";

interface Stats {
  originalRows: number;
  cleanedRows: number;
  filteredRows: number;
  uomMatched: number;
  countryCodeRowsRemoved: number;
  stickerRows: number;
}

type Status = "idle" | "processing" | "done" | "error";

export default function FileUpload() {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [stats, setStats] = useState<Stats | null>(null);
  const [errorMsg, setErrorMsg] = useState("");
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [downloadName, setDownloadName] = useState("");
  const [dragging, setDragging] = useState(false);

  const handleFile = useCallback((f: File) => {
    setFile(f);
    setStatus("idle");
    setStats(null);
    setErrorMsg("");
    if (downloadUrl) URL.revokeObjectURL(downloadUrl);
    setDownloadUrl(null);
  }, [downloadUrl]);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      const dropped = e.dataTransfer.files[0];
      if (dropped) handleFile(dropped);
    },
    [handleFile]
  );

  const onInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selected = e.target.files?.[0];
    if (selected) handleFile(selected);
  };

  const processFile = async () => {
    if (!file) return;
    setStatus("processing");
    setErrorMsg("");

    try {
      const fd = new FormData();
      fd.append("file", file);

      const res = await fetch("/api/process", { method: "POST", body: fd });

      if (!res.ok) {
        const body = await res.json();
        throw new Error(body.error ?? "Unknown error");
      }

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);

      const rawStats = res.headers.get("X-Stats");
      if (rawStats) setStats(JSON.parse(rawStats));

      const disposition = res.headers.get("Content-Disposition") ?? "";
      const nameMatch = disposition.match(/filename="([^"]+)"/);
      setDownloadName(nameMatch?.[1] ?? "inventory-count-cleaned.xlsx");
      setDownloadUrl(url);
      setStatus("done");
    } catch (err) {
      setErrorMsg(err instanceof Error ? err.message : "Processing failed");
      setStatus("error");
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 px-6 py-4 flex items-center gap-4">
        <Image
          src="/at-logo.jpg"
          alt="Agri Technovation"
          width={160}
          height={48}
          className="object-contain h-12 w-auto"
          priority
        />
        <div className="h-8 w-px bg-gray-200" />
        <div>
          <h1 className="text-lg font-semibold text-gray-800">
            Inventory Count Processor
          </h1>
          <p className="text-sm text-gray-500">
            Sage Evolution export cleaner
          </p>
        </div>
      </header>

      {/* Main */}
      <main className="flex-1 flex items-start justify-center pt-12 px-4">
        <div className="w-full max-w-xl space-y-6">
          {/* Upload zone */}
          <div
            onDrop={onDrop}
            onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
            onDragLeave={() => setDragging(false)}
            className={`relative border-2 border-dashed rounded-xl p-10 text-center cursor-pointer transition-colors ${
              dragging
                ? "border-green-500 bg-green-50"
                : file
                ? "border-green-400 bg-green-50"
                : "border-gray-300 bg-white hover:border-gray-400"
            }`}
            onClick={() => document.getElementById("file-input")?.click()}
          >
            <input
              id="file-input"
              type="file"
              accept=".xls,.xlsx,.xlsm"
              className="hidden"
              onChange={onInputChange}
            />

            {file ? (
              <>
                <div className="text-4xl mb-3">📊</div>
                <p className="font-medium text-gray-800">{file.name}</p>
                <p className="text-sm text-gray-500 mt-1">
                  {(file.size / 1024).toFixed(0)} KB — click to change
                </p>
              </>
            ) : (
              <>
                <div className="text-4xl mb-3">📂</div>
                <p className="font-medium text-gray-700">
                  Drop your Sage Evolution export here
                </p>
                <p className="text-sm text-gray-400 mt-1">
                  or click to browse — .xls / .xlsx
                </p>
              </>
            )}
          </div>

          {/* Process button */}
          <button
            onClick={processFile}
            disabled={!file || status === "processing"}
            className="w-full py-3 rounded-xl font-semibold text-white transition-colors disabled:opacity-40 disabled:cursor-not-allowed bg-green-600 hover:bg-green-700 active:bg-green-800"
          >
            {status === "processing" ? "Processing…" : "Process File"}
          </button>

          {/* Error */}
          {status === "error" && (
            <div className="rounded-xl border border-red-200 bg-red-50 p-4 text-red-700 text-sm">
              {errorMsg}
            </div>
          )}

          {/* Success */}
          {status === "done" && stats && (
            <div className="rounded-xl border border-green-200 bg-white p-5 space-y-4">
              <h2 className="font-semibold text-gray-800">
                File processed successfully
              </h2>

              <div className="grid grid-cols-2 gap-3 text-sm">
                <StatCard label="Original rows" value={stats.originalRows} />
                <StatCard label="Cleaned rows" value={stats.cleanedRows} />
                <StatCard label="Rows filtered out" value={stats.filteredRows} color="amber" />
                <StatCard label="UOM extracted" value={stats.uomMatched} color="green" />
                <StatCard label="Country code rows removed" value={stats.countryCodeRowsRemoved} color="amber" />
                <StatCard label="Sticker rows" value={stats.stickerRows} color="amber" />
              </div>

              <a
                href={downloadUrl!}
                download={downloadName}
                className="flex items-center justify-center gap-2 w-full py-3 rounded-xl font-semibold text-white bg-green-600 hover:bg-green-700 active:bg-green-800 transition-colors"
              >
                ⬇ Download Processed File
              </a>
            </div>
          )}

          {/* What it does */}
          <details className="rounded-xl border border-gray-200 bg-white">
            <summary className="px-5 py-3 cursor-pointer text-sm font-medium text-gray-700 select-none">
              What does this app do?
            </summary>
            <ul className="px-5 pb-4 pt-1 text-sm text-gray-600 space-y-1 list-disc list-inside">
              <li>Sheet 1 — original data, untouched</li>
              <li>Sheet 2 — cleaned: noise rows removed, duplicate headers removed</li>
              <li>Sheet 2 — columns removed: System Qty, Actual Qty, Variance</li>
              <li>Sheet 2 — UOM extracted from item description into its own column</li>
              <li>Sheet 2 — rows with country code item codes removed (e.g. Aus_, NZ_, AUS0021_B)</li>
              <li>Sheet 3 — Cleaned No IBC: same as Cleaned minus 1000L / IBC rows</li>
              <li>Sheet 4 — Stickers: rows where description contains &quot;sticker&quot;</li>
              <li>Sheet 5 — every filtered-out row with the reason it was removed</li>
            </ul>
          </details>
        </div>
      </main>

      <footer className="text-center text-xs text-gray-400 py-4">
        Agri Technovation — Internal Tools
      </footer>
    </div>
  );
}

function StatCard({
  label,
  value,
  color = "gray",
}: {
  label: string;
  value: number;
  color?: "gray" | "green" | "amber";
}) {
  const colors = {
    gray: "bg-gray-50 text-gray-800",
    green: "bg-green-50 text-green-800",
    amber: "bg-amber-50 text-amber-800",
  };
  return (
    <div className={`rounded-lg p-3 ${colors[color]}`}>
      <div className="text-xl font-bold">{value.toLocaleString()}</div>
      <div className="text-xs mt-0.5 opacity-75">{label}</div>
    </div>
  );
}
