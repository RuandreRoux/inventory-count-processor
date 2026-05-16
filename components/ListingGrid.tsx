import type { Listing } from "@/lib/types";
import ListingCard from "./ListingCard";

export default function ListingGrid({ listings }: { listings: Listing[] }) {
  if (listings.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center py-24 text-center">
        <div className="text-5xl mb-4">&#x1F50D;</div>
        <h3 className="text-lg font-semibold text-zinc-300">No listings found</h3>
        <p className="text-sm text-zinc-500 mt-2 max-w-xs">
          Try adjusting your search term or filters to see more results.
        </p>
      </div>
    );
  }

  return (
    <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 xl:grid-cols-3">
      {listings.map((listing) => (
        <ListingCard key={listing.id} listing={listing} />
      ))}
    </div>
  );
}
