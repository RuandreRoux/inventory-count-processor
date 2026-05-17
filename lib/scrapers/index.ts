/**
 * Scraper orchestrator.
 *
 * Requires SCRAPING_ENABLED=true + a Playwright-capable runtime.
 */

import type { Listing, SearchFilters, RankingWeights } from '@/lib/types';
import { rankListings } from '@/lib/ranking';
import { cacheGet, cacheSet, cacheBuildKey } from './cache';

export const SCRAPING_ENABLED = process.env.SCRAPING_ENABLED === 'true';

// Per-query in-flight promise — prevents duplicate concurrent scrapes for the same query
const inFlight = new Map<string, Promise<Listing[]>>();

interface FetchOptions {
  filters: SearchFilters;
  weights: RankingWeights;
}

export async function fetchListings({ filters, weights }: FetchOptions): Promise<Listing[]> {
  // Cache key is query-only: scrape a broad result set once, then apply filters in memory.
  // This means changing price/mileage/year filters hits the cache instead of re-scraping.
  const cacheKey = cacheBuildKey({ q: filters.query.trim().toLowerCase() });

  const cached = cacheGet(cacheKey);
  if (cached) {
    return rankListings(applyPostFilters(cached, filters), weights);
  }

  // If a scrape is already running for this query, wait for it instead of launching a second browser
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
  if (!SCRAPING_ENABLED) return [];

  const { launchBrowser } = await import('./browser');
  const { scrapeAutoTrader } = await import('./autotrader');
  const { scrapeCarsCoza } = await import('./carscoza');
  const { scrapeChangeCars } = await import('./changecars');

  const browser = await launchBrowser();
  try {
    // Pass query only — numeric filters applied post-scrape from cache
    const [at, cz, cc] = await Promise.allSettled([
      scrapeAutoTrader(browser, filters.query, {}),
      scrapeCarsCoza(browser, filters.query, {}),
      scrapeChangeCars(browser, filters.query, {}),
    ]);

    if (at.status === 'rejected') console.error('[DreamCar] AutoTrader failed:', at.reason);
    if (cz.status === 'rejected') console.error('[DreamCar] Cars.co.za failed:', cz.reason);
    if (cc.status === 'rejected') console.error('[DreamCar] ChangeCars failed:', cc.reason);

    const all = [
      ...(at.status === 'fulfilled' ? at.value : []),
      ...(cz.status === 'fulfilled' ? cz.value : []),
      ...(cc.status === 'fulfilled' ? cc.value : []),
    ];
    console.log('[DreamCar] Total listings scraped:', all.length);
    return all;
  } finally {
    await browser.close();
  }
}

function applyPostFilters(listings: Listing[], f: SearchFilters): Listing[] {
  let r = listings;
  // Text query filter — keep only listings matching the search query
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
