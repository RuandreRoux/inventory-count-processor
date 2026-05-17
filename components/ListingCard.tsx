import type { Listing } from "@/lib/types";
import { SOURCE_BAR_COLORS } from "@/lib/types";
import ScoreBadge from "./ScoreBadge";
import SourceBadge from "./SourceBadge";

const CAR_COLORS: Record<string, string> = {
  White: "#e4e4e7", Silver: "#a1a1aa", Grey: "#71717a", Black: "#18181b",
  Red: "#dc2626", Blue: "#2563eb", "Tornado Red": "#dc2626", "Glacier White": "#f4f4f5",
  "Alpine White": "#fafafa", "Polar White": "#f4f4f5", "Arctic White": "#e4e4e7",
  "Soul Red Crystal": "#b91c1c", "Intense Blue": "#1d4ed8", "Obsidian Black": "#09090b",
  Graphite: "#52525b", "Phantom Black": "#09090b", "Deep Black": "#09090b",
  "Meteor Grey": "#6b7280", "Black Sapphire": "#1e3a5f", "Ravenna Blue": "#1e40af",
  "Aurora Black Pearl": "#18181b", "Mythos Black": "#09090b", "Florett Silver": "#9ca3af",
  "Gondwana Stone": "#a8896a", "Fuji White": "#f1f5f9", "Mineral Grey": "#6b7280",
  "Storm White": "#f1f5f9", "Pearl White": "#fafafa", "Crystal White": "#f4f4f5",
  "Attitude Black": "#18181b", "Sparkling Silver": "#9ca3af", "Shimmering Silver": "#a1a1aa",
  "Polymetal Grey": "#64748b", "Sapphire Blue": "#1d4ed8", "Phoenix Orange": "#ea580c",
  Orange: "#ea580c",
};

function CarPlaceholder({ color, make, model }: { color: string; make: string; model: string }) {
  const bg = CAR_COLORS[color] ?? "#52525b";
  return (
    <div className="relative flex h-44 w-full items-center justify-center overflow-hidden bg-zinc-900">
      <div
        className="absolute inset-0 opacity-10"
        style={{ background: `radial-gradient(ellipse at 50% 60%, ${bg}88 0%, transparent 70%)` }}
      />
      <svg viewBox="0 0 220 100" className="w-48 drop-shadow-2xl" aria-label={`${make} ${model}`}>
        <ellipse cx="110" cy="82" rx="90" ry="8" fill="#000" opacity="0.4" />
        <rect x="20" y="48" width="180" height="32" rx="8" fill={bg} />
        <path d="M45 48 Q60 20 95 18 L140 18 Q165 20 178 48 Z" fill={bg} />
        <path
          d="M50 48 Q63 24 96 22 L139 22 Q162 24 173 48 Z"
          fill="none" stroke="rgba(255,255,255,0.12)" strokeWidth="1"
        />
        <rect x="52" y="24" width="48" height="22" rx="4" fill="rgba(147,210,255,0.18)" stroke="rgba(255,255,255,0.1)" strokeWidth="0.5" />
        <rect x="105" y="24" width="48" height="22" rx="4" fill="rgba(147,210,255,0.18)" stroke="rgba(255,255,255,0.1)" strokeWidth="0.5" />
        <circle cx="55" cy="80" r="14" fill="#18181b" />
        <circle cx="55" cy="80" r="10" fill="#27272a" />
        <circle cx="55" cy="80" r="5" fill="#3f3f46" />
        <circle cx="165" cy="80" r="14" fill="#18181b" />
        <circle cx="165" cy="80" r="10" fill="#27272a" />
        <circle cx="165" cy="80" r="5" fill="#3f3f46" />
        <rect x="22" y="56" width="18" height="10" rx="4" fill="rgba(253,224,71,0.85)" />
        <rect x="180" y="56" width="18" height="10" rx="4" fill="rgba(239,68,68,0.7)" />
      </svg>
      <div className="absolute bottom-0 left-0 right-0 h-px bg-gradient-to-r from-transparent via-white/10 to-transparent" />
    </div>
  );
}

function CarThumbnail({ listing }: { listing: Listing }) {
  if (listing.imageUrl) {
    return (
      <div className="relative h-44 w-full overflow-hidden bg-zinc-900">
        {/* eslint-disable-next-line @next/next/no-img-element */}
        <img
          src={listing.imageUrl}
          alt={`${listing.make} ${listing.model}`}
          className="h-full w-full object-cover"
          onError={(e) => { (e.currentTarget as HTMLImageElement).style.display = 'none'; }}
        />
        <div className="absolute bottom-0 left-0 right-0 h-px bg-gradient-to-r from-transparent via-white/10 to-transparent" />
      </div>
    );
  }
  return <CarPlaceholder color={listing.color} make={listing.make} model={listing.model} />;
}

export default function ListingCard({ listing }: { listing: Listing }) {
  const barClass = SOURCE_BAR_COLORS[listing.source];

  // Link directly to the external listing when a URL is available;
  // fall back to the internal detail page for mock/offline listings.
  const href = listing.url || `/vehicle/${listing.id}`;
  const isExternal = !!listing.url;

  return (
    <a
      href={href}
      target={isExternal ? "_blank" : undefined}
      rel={isExternal ? "noopener noreferrer" : undefined}
      className="group block"
    >
      <div className="relative rounded-xl border border-zinc-800 bg-[#111118] overflow-hidden transition-all duration-200 hover:border-amber-500/40 hover:shadow-[0_0_20px_rgba(245,158,11,0.08)]">
        <div className={`h-0.5 w-full bg-gradient-to-r ${barClass}`} />
        <div className="relative">
          <CarThumbnail listing={listing} />
          {listing.score !== undefined && (
            <div className="absolute right-3 top-3">
              <ScoreBadge score={listing.score} size="md" />
            </div>
          )}
        </div>

        <div className="p-4 space-y-3">
          <div>
            <h3 className="font-bold text-white leading-tight group-hover:text-amber-400 transition-colors">
              {listing.make} {listing.model}
            </h3>
            <p className="text-xs text-zinc-400 mt-0.5 truncate">{listing.variant}</p>
          </div>

          <div className="flex flex-wrap gap-1.5">
            <span className="rounded bg-zinc-800 px-2 py-0.5 text-[10px] text-zinc-400">
              {listing.year}
            </span>
            <span className="rounded bg-zinc-800 px-2 py-0.5 text-[10px] text-zinc-400 capitalize">
              {listing.transmission}
            </span>
            <span className="rounded bg-zinc-800 px-2 py-0.5 text-[10px] text-zinc-400 capitalize">
              {listing.fuel}
            </span>
            {listing.serviceHistory && (
              <span className="rounded bg-amber-500/10 px-2 py-0.5 text-[10px] text-amber-400 border border-amber-500/20">
                Full history
              </span>
            )}
          </div>

          <div className="flex items-end justify-between">
            <div>
              <p className="text-xl font-black text-white">
                R{listing.price.toLocaleString("en-ZA")}
              </p>
              <p className="text-xs text-zinc-500 mt-0.5">
                {listing.mileage.toLocaleString("en-ZA")} km · {listing.city}
              </p>
            </div>
            <SourceBadge source={listing.source} />
          </div>
        </div>
      </div>
    </a>
  );
}
