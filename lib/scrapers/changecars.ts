/**
 * ChangeCars SA scraper (changecars.co.za)
 *
 * Intercepts internal JSON API calls for listings, falls back to DOM link extraction.
 */

import type { Browser, Response } from 'playwright-core';
import type { Listing } from '@/lib/types';
import { BROWSER_HEADERS } from './browser';
import { normalizeRaw } from './normalize';

const BASE_URL = 'https://www.changecars.co.za';

const MAKES: Record<string, string> = {
  toyota: 'toyota', volkswagen: 'volkswagen', vw: 'volkswagen', ford: 'ford',
  bmw: 'bmw', 'mercedes-benz': 'mercedes-benz', mercedes: 'mercedes-benz',
  hyundai: 'hyundai', kia: 'kia', mazda: 'mazda', isuzu: 'isuzu',
  nissan: 'nissan', honda: 'honda', suzuki: 'suzuki', renault: 'renault',
  audi: 'audi', haval: 'haval', 'land rover': 'land-rover',
};

function buildSearchUrl(query: string, filters: { maxPrice?: number; maxMileage?: number; minYear?: number }): string {
  const q = query.toLowerCase();
  const makeEntry = Object.entries(MAKES).find(([k]) => q.includes(k));
  const make = makeEntry?.[1] ?? '';
  const model = make ? q.replace(makeEntry![0], '').trim().replace(/\s+/g, '-') : '';

  const params = new URLSearchParams();
  if (make) params.set('make', make);
  if (model) params.set('model', model);
  if (filters.maxPrice) params.set('price_max', String(filters.maxPrice));
  if (filters.maxMileage) params.set('mileage_max', String(filters.maxMileage));
  if (filters.minYear) params.set('year_min', String(filters.minYear));

  const path = make && model ? `/used/${make}/${model}/` : make ? `/used/${make}/` : '/used-cars/';
  return `${BASE_URL}${path}?${params.toString()}`;
}

function extractVehicles(json: unknown): Array<Record<string, unknown>> {
  if (!json || typeof json !== 'object') return [];
  const obj = json as Record<string, unknown>;
  const candidates = [
    obj['listings'], obj['results'], obj['data'], obj['vehicles'],
    obj['adverts'], obj['items'], obj['cars'], obj['hits'],
    (obj['data'] as Record<string, unknown>)?.['listings'],
    (obj['data'] as Record<string, unknown>)?.['results'],
  ];
  for (const c of candidates) {
    if (Array.isArray(c) && c.length > 0 && typeof c[0] === 'object') return c as Array<Record<string, unknown>>;
  }
  if (Array.isArray(json) && json.length > 0 && typeof json[0] === 'object') return json as Array<Record<string, unknown>>;
  return [];
}

function vehicleToListing(v: Record<string, unknown>): Listing | null {
  const make = (v['make'] ?? v['Make'] ?? v['manufacturer'] ?? '') as string;
  const model = (v['model'] ?? v['Model'] ?? '') as string;
  const variant = (v['variant'] ?? v['Variant'] ?? v['trim'] ?? v['derivative'] ?? '') as string;
  const year = Number(v['year'] ?? v['Year'] ?? v['modelYear'] ?? 0);
  const price = Number(v['price'] ?? v['Price'] ?? v['sellingPrice'] ?? v['asking_price'] ?? 0);
  const mileage = Number(v['mileage'] ?? v['Mileage'] ?? v['km'] ?? v['odometer'] ?? 0);
  const url = (v['url'] ?? v['link'] ?? v['listingUrl'] ?? v['permalink'] ?? '') as string;

  if (!make || !price || !year) return null;

  const title = `${year} ${make} ${model} ${variant}`.trim();
  return normalizeRaw(
    {
      title,
      priceText: `R ${price}`,
      mileageText: `${mileage} km`,
      locationText: (v['province'] ?? v['region'] ?? v['location'] ?? v['city'] ?? '') as string,
      url: url.startsWith('http') ? url : url ? `${BASE_URL}${url}` : '',
      serviceHistory: Boolean(v['serviceHistory'] ?? v['fsh'] ?? v['fullServiceHistory'] ?? false),
    },
    'changecars',
  );
}

export async function scrapeChangeCars(
  browser: Browser,
  query: string,
  filters: { maxPrice?: number; maxMileage?: number; minYear?: number },
): Promise<Listing[]> {
  const page = await browser.newPage();
  const capturedVehicles: Array<Record<string, unknown>> = [];
  const capturedUrls: string[] = [];

  const onResponse = async (response: Response) => {
    try {
      const ct = response.headers()['content-type'] ?? '';
      if (!ct.includes('json')) return;
      if (!response.url().includes('changecars.co.za')) return;
      const json = await response.json().catch(() => null);
      if (!json) return;
      const vehicles = extractVehicles(json);
      if (vehicles.length > 0) {
        capturedUrls.push(response.url());
        capturedVehicles.push(...vehicles);
        console.log('[ChangeCars] Captured', vehicles.length, 'vehicles from:', response.url().slice(0, 120));
      }
    } catch { /* ignore */ }
  };

  page.on('response', onResponse);

  try {
    await page.setViewportSize({ width: 1280, height: 900 });
    await page.setExtraHTTPHeaders(BROWSER_HEADERS);

    const url = buildSearchUrl(query, filters);
    console.log('[ChangeCars] Navigating to:', url);

    const response = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    const status = response?.status() ?? 0;
    console.log('[ChangeCars] HTTP status:', status);

    if (status === 503 || status === 403) {
      console.log('[ChangeCars] Blocked (status', status, ') — skipping');
      return [];
    }

    await page.waitForLoadState('networkidle').catch(() => null);
    await page.waitForTimeout(4000);
    await page.evaluate(() => window.scrollBy(0, 800));
    await page.waitForTimeout(2000);

    const pageTitle = await page.title();
    console.log('[ChangeCars] Page title:', pageTitle);
    console.log('[ChangeCars] API vehicles captured:', capturedVehicles.length);

    if (capturedVehicles.length === 0) {
      // Log debug info to help identify link patterns and page structure
      const debugInfo = await page.evaluate(() => {
        const links = Array.from(document.querySelectorAll('a[href]'))
          .map(a => (a as HTMLAnchorElement).href)
          .filter(h => h.includes('changecars.co.za') && h.length > 35);
        const unique = [...new Set(links)].slice(0, 20);
        const bodyPreview = document.body?.innerText?.slice(0, 400) ?? '';
        const bodyLen = document.body?.innerHTML?.length ?? 0;
        return { bodyLen, bodyPreview, sampleLinks: unique };
      });
      console.log('[ChangeCars] bodyLen:', debugInfo.bodyLen);
      console.log('[ChangeCars] Body preview:', debugInfo.bodyPreview);
      console.log('[ChangeCars] Sample links:', JSON.stringify(debugInfo.sampleLinks));

      // DOM fallback — try to find listing links
      const raw = await page.evaluate(() => {
        // Try common listing link patterns
        const patterns = ['a[href*="/listing/"]', 'a[href*="/advert/"]', 'a[href*="/vehicle/"]', 'a[href*="/cars/"]', 'a[href*="/used/"]'];
        const adLinks = patterns.flatMap(p => Array.from(document.querySelectorAll(p)))
          .filter(a => {
            const href = (a as HTMLAnchorElement).href ?? '';
            return href.split('/').length >= 5;
          });

        type Extract = { title: string; priceText: string; mileageText: string; locationText: string; url: string; hasServiceHistory: boolean };
        const seen = new Set<Element>();
        const results: Extract[] = [];

        for (const link of adLinks.slice(0, 30)) {
          let container: Element | null = link.parentElement;
          for (let i = 0; i < 8; i++) {
            if (!container) break;
            const text = container.textContent ?? '';
            if (text.includes('R ') && text.length > 50 && text.length < 3000) break;
            container = container.parentElement;
          }
          if (!container || seen.has(container)) continue;
          seen.add(container);

          const fullText = container.textContent ?? '';
          const href = (link as HTMLAnchorElement).href ?? link.getAttribute('href') ?? '';
          const absUrl = href.startsWith('http') ? href : `https://www.changecars.co.za${href}`;
          const heading = container.querySelector('h1, h2, h3, h4');
          const title = heading?.textContent?.trim() ?? link.textContent?.trim() ?? '';
          const priceMatch = fullText.match(/R\s?[\d\s,]+/);
          const priceText = priceMatch ? priceMatch[0].trim() : '';
          const kmMatch = fullText.match(/[\d\s,]+\s*km/i);
          const mileageText = kmMatch ? kmMatch[0].trim() : '';
          const yearMatch = fullText.match(/\b(19|20)\d{2}\b/);
          const provinces = ['gauteng','western cape','kwazulu-natal','eastern cape','limpopo','mpumalanga','north west','free state','northern cape'];
          const matchedProvince = provinces.find(p => fullText.toLowerCase().includes(p)) ?? '';
          const hasServiceHistory = /\b(full service|fsh|service history)\b/i.test(fullText);

          if (title && priceText) {
            results.push({
              title: yearMatch ? `${yearMatch[0]} ${title}` : title,
              priceText, mileageText, locationText: matchedProvince, url: absUrl, hasServiceHistory,
            });
          }
        }
        return results;
      });

      console.log('[ChangeCars] DOM fallback extracts:', raw.length);
      const listings: Listing[] = [];
      for (const r of raw) {
        const n = normalizeRaw(
          { title: r.title, priceText: r.priceText, mileageText: r.mileageText, locationText: r.locationText, url: r.url, serviceHistory: r.hasServiceHistory },
          'changecars',
        );
        if (n) listings.push(n);
      }
      console.log('[ChangeCars] Listings extracted:', listings.length);
      return listings;
    }

    // Build listings from API data
    const seen = new Set<string>();
    const listings: Listing[] = [];
    for (const v of capturedVehicles.slice(0, 40)) {
      const listing = vehicleToListing(v);
      if (listing && !seen.has(listing.id)) {
        seen.add(listing.id);
        listings.push(listing);
      }
    }
    console.log('[ChangeCars] Listings extracted:', listings.length);
    return listings;
  } finally {
    page.off('response', onResponse);
    await page.close();
  }
}
