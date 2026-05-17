/**
 * Cars.co.za scraper
 *
 * Listings are loaded via an internal JSON API after page render.
 * We intercept those responses rather than scraping the DOM.
 */

import type { Browser, Response } from 'playwright-core';
import type { Listing } from '@/lib/types';
import { normalizeRaw } from './normalize';

const SEARCH_URL = 'https://www.cars.co.za/usedcars/';

function normalizeMake(query: string): { make: string; model: string } {
  const MAKES: Record<string, string> = {
    toyota: 'Toyota', volkswagen: 'Volkswagen', vw: 'Volkswagen', ford: 'Ford',
    bmw: 'BMW', 'mercedes-benz': 'Mercedes-Benz', mercedes: 'Mercedes-Benz',
    hyundai: 'Hyundai', kia: 'Kia', mazda: 'Mazda', isuzu: 'Isuzu',
    nissan: 'Nissan', honda: 'Honda', suzuki: 'Suzuki', renault: 'Renault',
    audi: 'Audi', haval: 'Haval', 'land rover': 'Land Rover',
  };
  const q = query.toLowerCase();
  const entry = Object.entries(MAKES).find(([k]) => q.includes(k));
  if (!entry) return { make: query, model: '' };
  const make = entry[1];
  const model = q.replace(entry[0], '').trim().replace(/\b\w/g, c => c.toUpperCase());
  return { make, model };
}

// Try to extract an array of vehicle listings from any JSON structure
function extractVehiclesFromJson(json: unknown): Array<Record<string, unknown>> {
  if (!json || typeof json !== 'object') return [];
  const obj = json as Record<string, unknown>;

  // Common API response shapes
  const candidates = [
    obj['listings'], obj['results'], obj['data'], obj['vehicles'],
    obj['adverts'], obj['items'], obj['cars'], obj['hits'],
    (obj['data'] as Record<string, unknown>)?.['listings'],
    (obj['data'] as Record<string, unknown>)?.['results'],
  ];
  for (const c of candidates) {
    if (Array.isArray(c) && c.length > 0 && typeof c[0] === 'object') {
      return c as Array<Record<string, unknown>>;
    }
  }
  if (Array.isArray(json) && json.length > 0 && typeof json[0] === 'object') {
    return json as Array<Record<string, unknown>>;
  }
  return [];
}

function vehicleToListing(v: Record<string, unknown>): Listing | null {
  // Extract common fields from various API response shapes
  const make = (v['make'] ?? v['Make'] ?? v['manufacturer'] ?? '') as string;
  const model = (v['model'] ?? v['Model'] ?? '') as string;
  const variant = (v['variant'] ?? v['Variant'] ?? v['trim'] ?? '') as string;
  const year = Number(v['year'] ?? v['Year'] ?? v['modelYear'] ?? 0);
  const price = Number(v['price'] ?? v['Price'] ?? v['sellingPrice'] ?? v['asking_price'] ?? 0);
  const mileage = Number(v['mileage'] ?? v['Mileage'] ?? v['km'] ?? v['odometer'] ?? 0);
  const url = (v['url'] ?? v['link'] ?? v['listingUrl'] ?? v['permalink'] ?? '') as string;

  const title = `${year} ${make} ${model} ${variant}`.trim();

  if (!make || !price || !year) return null;

  return normalizeRaw(
    {
      title,
      priceText: `R ${price}`,
      mileageText: `${mileage} km`,
      locationText: (v['province'] ?? v['region'] ?? v['location'] ?? '') as string,
      url: url.startsWith('http') ? url : url ? `https://www.cars.co.za${url}` : '',
      serviceHistory: Boolean(v['serviceHistory'] ?? v['fsh'] ?? v['fullServiceHistory'] ?? false),
    },
    'carscoza',
  );
}

export async function scrapeCarsCoza(
  browser: Browser,
  query: string,
  filters: { maxPrice?: number; maxMileage?: number; minYear?: number },
): Promise<Listing[]> {
  const page = await browser.newPage();
  const capturedJsonUrls: string[] = [];
  const capturedVehicles: Array<Record<string, unknown>> = [];

  // Intercept JSON API responses
  const onResponse = async (response: Response) => {
    try {
      const ct = response.headers()['content-type'] ?? '';
      if (!ct.includes('json')) return;
      const url = response.url();
      if (!url.includes('cars.co.za')) return;

      const json = await response.json().catch(() => null);
      if (!json) return;

      const vehicles = extractVehiclesFromJson(json);
      if (vehicles.length > 0) {
        capturedJsonUrls.push(url);
        capturedVehicles.push(...vehicles);
        console.log('[CarsCoza] Captured', vehicles.length, 'vehicles from:', url.slice(0, 120));
      }
    } catch {
      // ignore
    }
  };

  page.on('response', onResponse);

  try {
    await page.setViewportSize({ width: 1280, height: 900 });
    await page.setExtraHTTPHeaders({
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
      'Accept-Language': 'en-ZA,en-GB;q=0.9,en;q=0.8',
      'Accept-Encoding': 'gzip, deflate, br',
      'Upgrade-Insecure-Requests': '1',
    });

    const { make, model } = normalizeMake(query);
    const params = new URLSearchParams();
    params.set('cat_make', make);
    if (model) params.set('cat_model', model);
    if (filters.maxPrice) params.set('price_to', String(filters.maxPrice));
    if (filters.maxMileage) params.set('mileage_to', String(filters.maxMileage));
    if (filters.minYear) params.set('year_from', String(filters.minYear));

    const url = `${SEARCH_URL}?${params.toString()}`;
    console.log('[CarsCoza] Navigating to:', url);

    const response = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    const status = response?.status() ?? 0;
    console.log('[CarsCoza] HTTP status:', status);

    if (status === 503 || status === 403) {
      console.log('[CarsCoza] Blocked (status', status, ') — skipping');
      return [];
    }

    // Wait for the AJAX listing requests to fire and complete
    await page.waitForLoadState('networkidle').catch(() => null);
    await page.waitForTimeout(4000);

    // Scroll to trigger lazy loading
    await page.evaluate(() => window.scrollBy(0, 800));
    await page.waitForTimeout(2000);

    console.log('[CarsCoza] JSON API URLs captured:', capturedJsonUrls.length);
    console.log('[CarsCoza] Total vehicles captured:', capturedVehicles.length);

    // Log all unique href pattern samples if no API data found
    if (capturedVehicles.length === 0) {
      const hrefSamples = await page.evaluate(() => {
        const links = Array.from(document.querySelectorAll('a[href]'));
        const hrefs = links.map(a => (a as HTMLAnchorElement).href).filter(h => h.includes('cars.co.za'));
        const unique = [...new Set(hrefs)].slice(0, 20);
        return unique;
      });
      console.log('[CarsCoza] Sample hrefs on page:', JSON.stringify(hrefSamples));
    }

    // Build Listing objects from captured API data
    const seen = new Set<string>();
    const listings: Listing[] = [];
    for (const v of capturedVehicles.slice(0, 40)) {
      const listing = vehicleToListing(v);
      if (listing && !seen.has(listing.id)) {
        seen.add(listing.id);
        listings.push(listing);
      }
    }

    console.log('[CarsCoza] Listings extracted:', listings.length);
    return listings;
  } finally {
    page.off('response', onResponse);
    await page.close();
  }
}
