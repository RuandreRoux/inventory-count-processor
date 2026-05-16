/**
 * Scraper orchestrator.
 *
 * Requires SCRAPING_ENABLED=true + a Playwright-capable runtime.
 *
 * Local dev:
 *   npx playwright install chromium
 *   SCRAPING_ENABLED=true npm run dev
 *
 * Vercel / serverless:
 *   Set SCRAPING_ENABLED=true in project environment variables.
 *   The @sparticuz/chromium-min binary is fetched from GitHub Releases at runtime.
 */

import type { Listing, SearchFilters, RankingWeights } from '@/lib/types';
import { rankListings } from '@/lib/ranking';
import { cacheGet, cacheSet, cacheBuildKey } from './cache';

export const SCRAPING_ENABLED = process.env.SCRAPING_ENABLED === 'true';

interface FetchOptions {
  filters: SearchFilters;
  weights: RankingWeights;
}

export async function fetchListings({ filters, weights }: FetchOptions): Promise<Listing[]> {
  const cacheKey = cacheBuildKey({
    q: filters.query,
    maxPrice: filters.maxPrice,
    maxMileage: filters.maxMileage,
    minYear: filters.minYear,
    condition: filters.condition,
    serviceHistoryOnly: filters.serviceHistoryOnly,
    transmission: filters.transmission,
    fuel: filters.fuel,
    province: filters.province,
  });

  const cached = cacheGet(cacheKey);
  if (cached) {
    return rankListings(applyPostFilters(cached, filters), weights);
  }

  const raw = await scrapeAll(filters);
  cacheSet(cacheKey, raw);
  return rankListings(applyPostFilters(raw, filters), weights);
}

async function scrapeAll(filters: SearchFilters): Promise<Listing[]> {
  if (!SCRAPING_ENABLED) {
    return [];
  }

  const { launchBrowser } = await import('./browser');
  const { scrapeAutoTrader } = await import('./autotrader');
  const { scrapeCarsCoza } = await import('./carscoza');

  const browser = await launchBrowser();
  try {
    const sf = {
      maxPrice: filters.maxPrice,
      maxMileage: filters.maxMileage,
      minYear: filters.minYear,
    };

    const [at, cz] = await Promise.allSettled([
      scrapeAutoTrader(browser, filters.query, sf),
      scrapeCarsCoza(browser, filters.query, sf),
    ]);

    if (at.status === 'rejected') console.error('[DreamCar] AutoTrader scrape failed:', at.reason);
    if (cz.status === 'rejected') console.error('[DreamCar] Cars.co.za scrape failed:', cz.reason);

    return [
      ...(at.status === 'fulfilled' ? at.value : []),
      ...(cz.status === 'fulfilled' ? cz.value : []),
    ];
  } finally {
    await browser.close();
  }
}

function applyPostFilters(listings: Listing[], f: SearchFilters): Listing[] {
  let r = listings;
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
