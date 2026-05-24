"use client";

import { useState } from "react";
import Navbar from "@/components/Navbar";
import { triggerScrape, fetchRecentJobs } from "./actions";

const POPULAR_MAKES = [
  "Volkswagen", "Toyota", "BMW", "Mercedes-Benz", "Ford",
  "Hyundai", "Kia", "Audi", "Nissan", "Mazda", "Honda",
  "Haval", "GWM", "Isuzu", "Renault", "Suzuki", "Land Rover",
];

type Job = Awaited<ReturnType<typeof fetchRecentJobs>>[number];
type ScrapeResult = { ok: boolean; stats?: { found: number; upserted: number; deactivated: number }; error?: string };

export default function AdminPage() {
  const [make, setMake] = useState("Volkswagen");
  const [model, setModel] = useState("Touareg");
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<ScrapeResult | null>(null);
  const [jobs, setJobs] = useState<Job[]>([]);
  const [jobsLoaded, setJobsLoaded] = useState(false);

  async function handleScrape() {
    setLoading(true);
    setResult(null);
    const res = await triggerScrape(make, model);
    setResult(res);
    setLoading(false);
    const updated = await fetchRecentJobs();
    setJobs(updated);
    setJobsLoaded(true);
  }

  async function handleLoadJobs() {
    const j = await fetchRecentJobs();
    setJobs(j);
    setJobsLoaded(true);
  }

  return (
    <div className="min-h-screen flex flex-col bg-[#09090f]">
      <Navbar />

      <div className="mx-auto w-full max-w-2xl px-4 py-10 space-y-8">
        <div>
          <h1 className="text-xl font-bold text-white">Scrape Manager</h1>
          <p className="text-sm text-zinc-500 mt-1">
            Trigger a scrape to populate the listing database for a specific make and model.
          </p>
        </div>

        {/* Scrape form */}
        <div className="rounded-xl border border-zinc-800 bg-[#111118] p-6 space-y-5">
          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <label className="text-xs font-medium text-zinc-400">Make</label>
              <input
                type="text"
                list="makes-list"
                value={make}
                onChange={(e) => setMake(e.target.value)}
                className="w-full rounded-lg border border-zinc-700 bg-zinc-900 px-3 py-2 text-sm text-white placeholder-zinc-600 focus:border-amber-500 focus:outline-none"
                placeholder="e.g. Volkswagen"
              />
              <datalist id="makes-list">
                {POPULAR_MAKES.map((m) => <option key={m} value={m} />)}
              </datalist>
            </div>
            <div className="space-y-1.5">
              <label className="text-xs font-medium text-zinc-400">Model</label>
              <input
                type="text"
                value={model}
                onChange={(e) => setModel(e.target.value)}
                className="w-full rounded-lg border border-zinc-700 bg-zinc-900 px-3 py-2 text-sm text-white placeholder-zinc-600 focus:border-amber-500 focus:outline-none"
                placeholder="e.g. Touareg"
              />
            </div>
          </div>

          <button
            onClick={handleScrape}
            disabled={loading || !make.trim()}
            className="w-full rounded-lg bg-amber-500 py-2.5 text-sm font-semibold text-zinc-900 hover:bg-amber-400 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
          >
            {loading ? "Scraping — this takes up to 30s…" : "Trigger Scrape"}
          </button>

          {result && (
            <div
              className={`rounded-lg border p-4 ${
                result.ok
                  ? "border-green-700/50 bg-green-900/20"
                  : "border-red-700/50 bg-red-900/20"
              }`}
            >
              {result.ok && result.stats ? (
                <div className="flex gap-8">
                  {[
                    { label: "Found", value: result.stats.found, color: "text-white" },
                    { label: "Upserted", value: result.stats.upserted, color: "text-white" },
                    { label: "Deactivated", value: result.stats.deactivated, color: "text-amber-400" },
                  ].map(({ label, value, color }) => (
                    <div key={label}>
                      <p className="text-xs text-zinc-400">{label}</p>
                      <p className={`text-2xl font-bold ${color}`}>{value}</p>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-sm text-red-400">{result.error}</p>
              )}
            </div>
          )}
        </div>

        {/* Job history */}
        <div className="rounded-xl border border-zinc-800 bg-[#111118] p-6 space-y-4">
          <div className="flex items-center justify-between">
            <h2 className="text-xs font-semibold uppercase tracking-widest text-zinc-500">
              Recent Jobs
            </h2>
            <button
              onClick={handleLoadJobs}
              className="text-xs text-zinc-500 hover:text-amber-400 transition-colors"
            >
              Refresh
            </button>
          </div>

          {!jobsLoaded ? (
            <button
              onClick={handleLoadJobs}
              className="text-sm text-zinc-500 hover:text-amber-400 transition-colors"
            >
              Load job history →
            </button>
          ) : jobs.length === 0 ? (
            <p className="text-sm text-zinc-600">No jobs yet.</p>
          ) : (
            <div className="space-y-2">
              {jobs.map((job) => (
                <div
                  key={job.id}
                  className="flex items-center justify-between rounded-lg border border-zinc-800 bg-zinc-900/50 px-4 py-3 text-sm"
                >
                  <div>
                    <span className="font-medium text-white">
                      {job.make} {job.model}
                    </span>
                    <span className="ml-2 text-xs text-zinc-500">
                      {new Date(job.started_at).toLocaleString()}
                    </span>
                  </div>
                  <div className="flex items-center gap-3">
                    {job.status === "completed" && job.listings_found !== null && (
                      <span className="text-xs text-zinc-400">{job.listings_found} found</span>
                    )}
                    {job.status === "failed" && job.error && (
                      <span className="text-xs text-red-400 max-w-[160px] truncate">{job.error}</span>
                    )}
                    <span
                      className={`rounded-full px-2 py-0.5 text-xs font-medium ${
                        job.status === "completed"
                          ? "bg-green-900/50 text-green-400"
                          : job.status === "failed"
                          ? "bg-red-900/50 text-red-400"
                          : "bg-amber-900/50 text-amber-400"
                      }`}
                    >
                      {job.status}
                    </span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
