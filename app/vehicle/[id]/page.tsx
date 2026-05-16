import { notFound } from "next/navigation";
import Link from "next/link";
import { LISTINGS } from "@/lib/mock-data";
import { rankListings } from "@/lib/ranking";
import { DEFAULT_WEIGHTS, SOURCE_LABELS, SOURCE_BAR_COLORS } from "@/lib/types";
import type { Listing } from "@/lib/types";
import ScoreBadge from "@/components/ScoreBadge";
import SourceBadge from "@/components/SourceBadge";
import ListingCard from "@/components/ListingCard";
import Navbar from "@/components/Navbar";

interface PageProps {
  params: Promise<{ id: string }>;
}

function ScoreBar({ label, value, weight, contribution }: {
  label: string; value: number; weight: number; contribution: number;
}) {
  return (
    <div className="space-y-1.5">
      <div className="flex justify-between text-xs">
        <span className="text-zinc-400">{label}</span>
        <span className="text-zinc-200">
          {Math.round(value * 100)}/100 × {weight}% ={" "}
          <span className="font-semibold text-amber-400">{contribution.toFixed(1)} pts</span>
        </span>
      </div>
      <div className="h-1.5 rounded-full bg-zinc-800 overflow-hidden">
        <div
          className="h-full rounded-full bg-gradient-to-r from-amber-600 to-amber-400 transition-all"
          style={{ width: `${Math.round(value * 100)}%` }}
        />
      </div>
    </div>
  );
}

function SpecRow({ label, value }: { label: string; value: string }) {
  return (
    <div className="flex items-center justify-between border-b border-zinc-800 py-3 last:border-0">
      <span className="text-sm text-zinc-500">{label}</span>
      <span className="text-sm font-medium text-zinc-200">{value}</span>
    </div>
  );
}

export default async function VehicleDetailPage({ params }: PageProps) {
  const { id } = await params;
  const raw = LISTINGS.find((l) => l.id === id);
  if (!raw) notFound();

  const [ranked] = rankListings([raw, ...LISTINGS], DEFAULT_WEIGHTS);
  const listing: Listing = ranked.id === id ? ranked : { ...raw, score: 50 };

  const similar = rankListings(
    LISTINGS.filter((l) => l.make === listing.make && l.id !== id),
    DEFAULT_WEIGHTS
  ).slice(0, 3);

  const weights = DEFAULT_WEIGHTS;
  const totalWeight =
    weights.price + weights.mileage + weights.year + weights.condition + weights.serviceHistory;
  const bd = listing.scoreBreakdown;

  const barClass = SOURCE_BAR_COLORS[listing.source];

  return (
    <div className="min-h-screen flex flex-col">
      <Navbar />

      <div className="mx-auto w-full max-w-5xl px-4 py-6 space-y-6">
        <Link
          href="/search"
          className="inline-flex items-center gap-1.5 text-sm text-zinc-500 hover:text-amber-400 transition-colors"
        >
          <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 19l-7-7 7-7" />
          </svg>
          Back to results
        </Link>

        <div className="rounded-xl overflow-hidden border border-zinc-800">
          <div className={`h-1.5 w-full bg-gradient-to-r ${barClass}`} />
          <div className="bg-[#111118] px-6 py-5">
            <div className="flex flex-wrap items-start justify-between gap-4">
              <div className="space-y-1.5">
                <div className="flex flex-wrap items-center gap-2">
                  <SourceBadge source={listing.source} />
                  <span className="rounded bg-zinc-800 px-2 py-0.5 text-xs text-zinc-400 capitalize">
                    {listing.condition}
                  </span>
                  {listing.serviceHistory && (
                    <span className="rounded bg-amber-500/10 border border-amber-500/20 px-2 py-0.5 text-xs text-amber-400">
                      Full service history
                    </span>
                  )}
                </div>
                <h1 className="text-2xl font-black text-white">
                  {listing.make} {listing.model}
                </h1>
                <p className="text-zinc-400">{listing.variant}</p>
              </div>
              <div className="flex items-center gap-4">
                {listing.score !== undefined && (
                  <div className="text-center">
                    <ScoreBadge score={listing.score} size="lg" />
                    <p className="text-[10px] text-zinc-500 mt-1">DreamScore</p>
                  </div>
                )}
                <div className="text-right">
                  <p className="text-3xl font-black text-white">
                    R{listing.price.toLocaleString("en-ZA")}
                  </p>
                  <p className="text-sm text-zinc-500 mt-0.5">
                    {listing.mileage.toLocaleString("en-ZA")} km
                  </p>
                </div>
              </div>
            </div>
          </div>
        </div>

        <div className="grid gap-6 lg:grid-cols-2">
          <div className="rounded-xl border border-zinc-800 bg-[#111118] p-5 space-y-1">
            <h2 className="text-xs font-semibold uppercase tracking-widest text-zinc-500 mb-4">Specifications</h2>
            <SpecRow label="Year" value={String(listing.year)} />
            <SpecRow label="Make" value={listing.make} />
            <SpecRow label="Model" value={listing.model} />
            <SpecRow label="Variant" value={listing.variant} />
            <SpecRow label="Mileage" value={`${listing.mileage.toLocaleString("en-ZA")} km`} />
            <SpecRow label="Transmission" value={listing.transmission.charAt(0).toUpperCase() + listing.transmission.slice(1)} />
            <SpecRow label="Fuel type" value={listing.fuel.charAt(0).toUpperCase() + listing.fuel.slice(1)} />
            <SpecRow label="Colour" value={listing.color} />
            <SpecRow label="Condition" value={listing.condition.charAt(0).toUpperCase() + listing.condition.slice(1)} />
            <SpecRow label="Service history" value={listing.serviceHistory ? "Full history" : "Not available"} />
            <SpecRow label="Province" value={listing.province} />
            <SpecRow label="City" value={listing.city} />
            <SpecRow label="Listed" value={new Date(listing.listedDate).toLocaleDateString("en-ZA", { year: "numeric", month: "long", day: "numeric" })} />
            <SpecRow label="Source" value={SOURCE_LABELS[listing.source]} />
          </div>

          <div className="space-y-4">
            {bd && listing.score !== undefined && (
              <div className="rounded-xl border border-zinc-800 bg-[#111118] p-5 space-y-4">
                <div className="flex items-center justify-between">
                  <h2 className="text-xs font-semibold uppercase tracking-widest text-zinc-500">DreamScore Breakdown</h2>
                  <div className="flex items-center gap-2">
                    <span className="text-2xl font-black text-amber-400">{listing.score}</span>
                    <span className="text-zinc-600">/100</span>
                  </div>
                </div>
                <p className="text-xs text-zinc-600">
                  Score is calculated by normalizing each factor across all search results and applying your chosen weights.
                </p>
                <div className="space-y-4">
                  <ScoreBar label="Price" value={bd.price} weight={Math.round((weights.price / totalWeight) * 100)} contribution={bd.price * weights.price / totalWeight * 100} />
                  <ScoreBar label="Mileage" value={bd.mileage} weight={Math.round((weights.mileage / totalWeight) * 100)} contribution={bd.mileage * weights.mileage / totalWeight * 100} />
                  <ScoreBar label="Year" value={bd.year} weight={Math.round((weights.year / totalWeight) * 100)} contribution={bd.year * weights.year / totalWeight * 100} />
                  <ScoreBar label="Condition" value={bd.condition} weight={Math.round((weights.condition / totalWeight) * 100)} contribution={bd.condition * weights.condition / totalWeight * 100} />
                  <ScoreBar label="Service History" value={bd.serviceHistory} weight={Math.round((weights.serviceHistory / totalWeight) * 100)} contribution={bd.serviceHistory * weights.serviceHistory / totalWeight * 100} />
                </div>
              </div>
            )}

            <div className="rounded-xl border border-zinc-800 bg-[#111118] p-5 space-y-3">
              <h2 className="text-xs font-semibold uppercase tracking-widest text-zinc-500">Description</h2>
              <p className="text-sm text-zinc-300 leading-relaxed">{listing.description}</p>
            </div>

            <a
              href={`https://${listing.source === "autotrader" ? "autotrader.co.za" : listing.source === "carscoza" ? "cars.co.za" : listing.source === "olx" ? "olx.co.za" : listing.source === "gumtree" ? "gumtree.co.za" : listing.source === "facebook" ? "facebook.com/marketplace" : "iol.co.za/motoring"}/listing/${listing.id}`}
              target="_blank"
              rel="noopener noreferrer"
              className="flex items-center justify-center gap-2 rounded-xl bg-gradient-to-r from-amber-500 to-orange-500 py-4 font-bold text-zinc-900 hover:from-amber-400 hover:to-orange-400 transition-all shadow-[0_0_20px_rgba(245,158,11,0.2)]"
            >
              View Original Listing on {SOURCE_LABELS[listing.source]}
              <svg className="h-4 w-4" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M10 6H6a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2v-4M14 4h6m0 0v6m0-6L10 14" />
              </svg>
            </a>
          </div>
        </div>

        {similar.length > 0 && (
          <div className="space-y-4">
            <h2 className="text-lg font-bold text-white">More {listing.make} listings</h2>
            <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
              {similar.map((l) => (
                <ListingCard key={l.id} listing={l} />
              ))}
            </div>
          </div>
        )}
      </div>

      <footer className="border-t border-zinc-800 px-6 py-4 text-center text-xs text-zinc-600 mt-auto">
        DreamCar ZA — South African Automotive Aggregator
      </footer>
    </div>
  );
}
