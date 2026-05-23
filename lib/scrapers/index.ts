import type { Listing, SearchFilters, RankingWeights } from '@/lib/types';
import { rankListings } from '@/lib/ranking';
import { cacheGet, cacheSet, cacheBuildKey } from './cache';

export const FIRECRAWL_ENABLED = Boolean(process.env.FIRECRAWL_API_KEY);
/** @deprecated No longer used — kept for backwards compat with the API route header */
export const SCRAPING_ENABLED = FIRECRAWL_ENABLED;

const inFlight = new Map<string, Promise<Listing[]>>();

interface FetchOptions {
  filters: SearchFilters;
  weights: RankingWeights;
}

export async function fetchListings({ filters, weights }: FetchOptions): Promise<Listing[]> {
  const cacheKey = cacheBuildKey({ q: filters.query.trim().toLowerCase() });

  const cached = cacheGet(cacheKey);
  if (cached) {
    return rankListings(applyPostFilters(cached, filters), weights);
  }

  let promise = inFlight.get(cacheKey);
  if (!promise) {
    promise = scrapeAll(filters).finally(() => inFlight.delete(cacheKey));
    inFlight.set(cacheKey, promise);
  }

  const raw = await promise;
  cacheSet(cacheKey, raw);
  return rankListings(applyPostFilters(raw, filters), weights);
}

async function scrapeAll(filters: SearchFilters): Promise<Listing[]> {
  if (!FIRECRAWL_ENABLED) return [];

  const { scrapeCarsCozaFirecrawl } = await import('./carscoza-firecrawl');
  const listings = await scrapeCarsCozaFirecrawl(filters.query).catch(err => {
    console.error('[DreamCar] Firecrawl scrape failed:', err);
    return [] as Listing[];
  });

  console.log('[DreamCar] Total listings scraped:', listings.length);
  return listings;
}

function applyPostFilters(listings: Listing[], f: SearchFilters): Listing[] {
  let r = listings;
  if (f.query.trim()) {
    const q = f.query.trim().toLowerCase();
    r = r.filter((l) =>
      l.make.toLowerCase().includes(q) ||
      l.model.toLowerCase().includes(q) ||
      `${l.make} ${l.model}`.toLowerCase().includes(q) ||
      `${l.make} ${l.model} ${l.variant}`.toLowerCase().includes(q)
    );
  }
  if (f.maxPrice !== undefined) r = r.filter((l) => l.price <= f.maxPrice!);
  if (f.maxMileage !== undefined) r = r.filter((l) => l.mileage <= f.maxMileage!);
  if (f.minYear !== undefined) r = r.filter((l) => l.year >= f.minYear!);
  if (f.condition) r = r.filter((l) => l.condition === f.condition);
  if (f.serviceHistoryOnly) r = r.filter((l) => l.serviceHistory);
  if (f.transmission) r = r.filter((l) => l.transmission === f.transmission);
  if (f.fuel) r = r.filter((l) => l.fuel === f.fuel);
  if (f.province) r = r.filter((l) => l.province === f.province);
  return r;
}
