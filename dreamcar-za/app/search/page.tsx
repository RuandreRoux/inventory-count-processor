"use client";

import { useSearchParams, useRouter } from "next/navigation";
import { useState, useEffect, useCallback, Suspense } from "react";
import type { Listing, SearchFilters, RankingWeights } from "@/lib/types";
import { DEFAULT_WEIGHTS } from "@/lib/types";
import Navbar from "@/components/Navbar";
import FilterPanel from "@/components/FilterPanel";
import WeightSliders from "@/components/WeightSliders";
import ListingGrid from "@/components/ListingGrid";

function buildQuery(filters: SearchFilters, weights: RankingWeights): string {
  const p = new URLSearchParams();
  if (filters.query) p.set("q", filters.query);
  if (filters.maxPrice) p.set("maxPrice", String(filters.maxPrice));
  if (filters.maxMileage) p.set("maxMileage", String(filters.maxMileage));
  if (filters.minYear) p.set("minYear", String(filters.minYear));
  if (filters.condition) p.set("condition", filters.condition);
  if (filters.serviceHistoryOnly) p.set("serviceHistoryOnly", "true");
  if (filters.transmission) p.set("transmission", filters.transmission);
  if (filters.fuel) p.set("fuel", filters.fuel);
  if (filters.province) p.set("province", filters.province);
  p.set(
    "weights",
    [weights.price, weights.mileage, weights.year, weights.condition, weights.serviceHistory].join(",")
  );
  return p.toString();
}

function SearchResults() {
  const searchParams = useSearchParams();
  const router = useRouter();
  const initialQ = searchParams.get("q") ?? "";

  const [filters, setFilters] = useState<SearchFilters>({ query: initialQ });
  const [weights, setWeights] = useState<RankingWeights>(DEFAULT_WEIGHTS);
  const [listings, setListings] = useState<Listing[]>([]);
  const [loading, setLoading] = useState(true);
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [dataSource, setDataSource] = useState<"live" | "disabled" | null>(null);

  const fetchListings = useCallback(async (f: SearchFilters, w: RankingWeights) => {
    setLoading(true);
    try {
      const res = await fetch(`/api/search?${buildQuery(f, w)}`);
      const source = res.headers.get("X-Data-Source");
      setDataSource(source === "live" ? "live" : "disabled");
      const data = await res.json();
      setListings(data);
    } catch {
      setListings([]);
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    fetchListings(filters, weights);
  }, [filters, weights, fetchListings]);

  const handleFilterChange = (partial: Partial<SearchFilters>) => {
    setFilters((prev) => ({ ...prev, ...partial }));
  };

  const handleNewSearch = (q: string) => {
    const newFilters = { ...filters, query: q };
    setFilters(newFilters);
    router.push(`/search?q=${encodeURIComponent(q)}`, { scroll: false });
  };

  const activeFilterCount = [
    filters.maxPrice, filters.maxMileage, filters.minYear,
    filters.condition, filters.serviceHistoryOnly, filters.transmission,
    filters.fuel, filters.province,
  ].filter(Boolean).length;

  return (
    <div className="min-h-screen flex flex-col">
      <Navbar />

      <div className="mx-auto w-full max-w-7xl flex flex-1 gap-0 px-0 sm:px-4 py-0 sm:py-6">
        {/* ── Sidebar overlay (mobile) ── */}
        {sidebarOpen && (
          <div
            className="fixed inset-0 z-40 bg-black/60 sm:hidden"
            onClick={() => setSidebarOpen(false)}
          />
        )}

        {/* ── Sidebar ── */}
        <aside
          className={`fixed sm:static inset-y-0 left-0 z-50 sm:z-auto w-72 shrink-0 overflow-y-auto border-r border-zinc-800 bg-[#09090f] sm:bg-transparent px-5 py-6 transition-transform duration-300 ${
            sidebarOpen ? "translate-x-0" : "-translate-x-full sm:translate-x-0"
          }`}
        >
          <div className="space-y-8">
            <FilterPanel filters={filters} onChange={handleFilterChange} />
            <div className="border-t border-zinc-800 pt-6">
              <WeightSliders weights={weights} onChange={setWeights} />
            </div>
          </div>
        </aside>

        {/* ── Main ── */}
        <main className="flex-1 min-w-0 px-4 sm:px-0 sm:pl-6 py-4 sm:py-0">
          {/* Toolbar */}
          <div className="mb-5 flex items-center justify-between gap-3">
            <div>
              <h1 className="text-lg font-bold text-white">
                {filters.query ? (
                  <>Results for &ldquo;<span className="text-amber-400">{filters.query}</span>&rdquo;</>
                ) : (
                  "All listings"
                )}
              </h1>
              {!loading && (
                <p className="text-sm text-zinc-500 mt-0.5">
                  {listings.length} {listings.length === 1 ? "listing" : "listings"} &middot; ranked by DreamCar score
                </p>
              )}
            </div>
            <button
              onClick={() => setSidebarOpen(true)}
              className="sm:hidden flex items-center gap-2 rounded-lg border border-zinc-700 bg-zinc-900 px-3 py-2 text-sm text-zinc-300"
            >
              <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2}
                  d="M3 4h18M3 12h12M3 20h6" />
              </svg>
              Filters
              {activeFilterCount > 0 && (
                <span className="rounded-full bg-amber-500 w-4 h-4 text-[10px] font-bold text-zinc-900 flex items-center justify-center">
                  {activeFilterCount}
                </span>
              )}
            </button>
          </div>

          {loading ? (
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 xl:grid-cols-3">
              {Array.from({ length: 6 }).map((_, i) => (
                <div
                  key={i}
                  className="rounded-xl border border-zinc-800 bg-[#111118] h-72 animate-pulse"
                />
              ))}
            </div>
          ) : listings.length === 0 && dataSource === "disabled" ? (
            <div className="flex flex-col items-center justify-center py-24 text-center">
              <div className="mb-4 rounded-full border border-amber-500/30 bg-amber-500/10 p-5">
                <svg className="h-8 w-8 text-amber-400" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={1.5}
                    d="M9.75 9.75l4.5 4.5m0-4.5l-4.5 4.5M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
              </div>
              <h2 className="text-lg font-semibold text-white mb-2">Live scraping is not enabled</h2>
              <p className="text-sm text-zinc-400 max-w-sm">
                Set the <code className="rounded bg-zinc-800 px-1.5 py-0.5 font-mono text-amber-400">SCRAPING_ENABLED=true</code> environment variable to fetch real listings from AutoTrader SA and Cars.co.za.
              </p>
            </div>
          ) : (
            <ListingGrid listings={listings} />
          )}
        </main>
      </div>
    </div>
  );
}

export default function SearchPage() {
  return (
    <Suspense>
      <SearchResults />
    </Suspense>
  );
}
