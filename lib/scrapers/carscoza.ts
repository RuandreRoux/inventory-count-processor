/**
 * Cars.co.za scraper
 *
 * Listing URLs follow the pattern:
 *   /for-sale/used/{year}-{make}-{model}-{variant}-{city}/{id}/
 * The slug contains year/make/model/variant/location — parsed directly from href.
 * Price and mileage are extracted from the card container text.
 */

import type { Browser, Response } from 'playwright-core';
import type { Listing } from '@/lib/types';
import { BROWSER_HEADERS } from './browser';
import { normalizeRaw, parsePrice, parseMileage } from './normalize';

const SEARCH_URL = 'https://www.cars.co.za/usedcars/';

const MAKES: Record<string, string> = {
  toyota: 'Toyota', volkswagen: 'Volkswagen', vw: 'Volkswagen', ford: 'Ford',
  bmw: 'BMW', 'mercedes-benz': 'Mercedes-Benz', mercedes: 'Mercedes-Benz',
  hyundai: 'Hyundai', kia: 'Kia', mazda: 'Mazda', isuzu: 'Isuzu',
  nissan: 'Nissan', honda: 'Honda', suzuki: 'Suzuki', renault: 'Renault',
  audi: 'Audi', haval: 'Haval', 'land rover': 'Land Rover',
};

function normalizeMake(query: string): { make: string; model: string } {
  const q = query.toLowerCase();
  const entry = Object.entries(MAKES).find(([k]) => q.includes(k));
  if (!entry) return { make: query, model: '' };
  const model = q.replace(entry[0], '').trim().replace(/\b\w/g, c => c.toUpperCase());
  return { make: entry[1], model };
}

/** Parse the Cars.co.za URL slug into structured fields.
 *  e.g. "2021-Toyota-Hilux-2.8-GD-6-Raised-Body-Legend-4x4-Auto-Double-Cab-Gauteng-Pretoria"
 */
function parseSlug(slug: string): { year: number; make: string; model: string; variant: string; city: string; province: string } {
  const SA_PROVINCES: Record<string, string> = {
    gauteng: 'Gauteng',
    'western-cape': 'Western Cape',
    'kwazulu-natal': 'KwaZulu-Natal',
    'eastern-cape': 'Eastern Cape',
    limpopo: 'Limpopo',
    mpumalanga: 'Mpumalanga',
    'north-west': 'North West',
    'north-west-province': 'North West',
    'free-state': 'Free State',
    'northern-cape': 'Northern Cape',
  };

  const parts = slug.split('-');
  const yearIdx = parts.findIndex(p => /^(19|20)\d{2}$/.test(p));
  const year = yearIdx >= 0 ? parseInt(parts[yearIdx], 10) : 0;

  const afterYear = parts.slice(yearIdx + 1);

  // Detect province by matching from the end backwards
  let provinceKey = '';
  let provinceEndIdx = afterYear.length;
  for (let len = 3; len >= 1; len--) {
    for (let i = afterYear.length - len; i >= 0; i--) {
      const candidate = afterYear.slice(i, i + len).join('-').toLowerCase();
      if (SA_PROVINCES[candidate]) {
        provinceKey = candidate;
        provinceEndIdx = i;
        break;
      }
    }
    if (provinceKey) break;
  }

  const province = SA_PROVINCES[provinceKey] ?? '';
  const city = afterYear.slice(provinceEndIdx + (provinceKey.split('-').length)).join(' ');

  // First word = make, second = model, rest = variant (up to province)
  const carParts = afterYear.slice(0, provinceEndIdx);
  const make = carParts[0] ?? '';
  const model = carParts[1] ?? '';
  const variant = carParts.slice(2).join(' ');

  return { year, make, model, variant, city, province };
}

function pickStr(v: Record<string, unknown>, ...keys: string[]): string {
  for (const k of keys) {
    const val = v[k];
    if (typeof val === 'string' && val.trim()) return val.trim();
  }
  return '';
}

function pickNum(v: Record<string, unknown>, ...keys: string[]): number {
  for (const k of keys) {
    const val = Number(v[k]);
    if (val > 0) return val;
  }
  return 0;
}

function extractFromVehicleArray(raw: Array<Record<string, unknown>>, source: string): Listing[] {
  const listings: Listing[] = [];
  const seenIds = new Set<string>();
  for (const v of raw.slice(0, 60)) {
    const make = pickStr(v, 'make', 'Make', 'manufacturer', 'makeDescription', 'makeName', 'brand');
    const model = pickStr(v, 'model', 'Model', 'modelDescription', 'modelName');
    const variant = pickStr(v, 'variant', 'Variant', 'derivative', 'trim', 'variantDescription', 'description', 'title');
    const year = pickNum(v, 'year', 'Year', 'modelYear', 'vehicleYear');
    const price = pickNum(v, 'price', 'Price', 'sellingPrice', 'askingPrice', 'listPrice', 'vehiclePrice', 'retail');
    const mileage = pickNum(v, 'mileage', 'Mileage', 'km', 'odometer', 'kilometres', 'kilometers', 'kms');
    const rawUrl = pickStr(v, 'url', 'link', 'listingUrl', 'permalink', 'adUrl', 'href', 'detailUrl');
    const imageUrl = pickStr(v, 'image', 'imageUrl', 'thumbnail', 'photo', 'primaryImage', 'mainImage', 'heroImage', 'imgUrl', 'picture');
    const locationText = pickStr(v, 'province', 'region', 'city', 'location', 'area', 'suburb');
    if (!make || !price || !year) {
      console.log(`[CarsCoza] Skip (${source}): make=${make} price=${price} year=${year} keys=${Object.keys(v).slice(0, 10).join(',')}`);
      continue;
    }
    const title = `${year} ${make} ${model} ${variant}`.trim();
    const listing = normalizeRaw(
      {
        title,
        priceText: `R ${price}`,
        mileageText: `${mileage} km`,
        locationText,
        url: rawUrl.startsWith('http') ? rawUrl : rawUrl ? `https://www.cars.co.za${rawUrl}` : '',
        serviceHistory: Boolean(v['serviceHistory'] ?? v['fsh'] ?? v['fullServiceHistory'] ?? false),
      },
      'carscoza',
    );
    if (listing && !seenIds.has(listing.id)) {
      listing.make = make; listing.model = model; listing.variant = variant; listing.year = year;
      if (imageUrl) listing.imageUrl = imageUrl.startsWith('http') ? imageUrl : `https://www.cars.co.za${imageUrl}`;
      seenIds.add(listing.id);
      listings.push(listing);
    }
  }
  return listings;
}

export async function scrapeCarsCoza(
  browser: Browser,
  query: string,
  filters: { maxPrice?: number; maxMileage?: number; minYear?: number },
): Promise<Listing[]> {
  const page = await browser.newPage();

  // Also try intercepting JSON API in case it fires
  const capturedVehicles: Array<Record<string, unknown>> = [];
  const onResponse = async (response: Response) => {
    try {
      const ct = response.headers()['content-type'] ?? '';
      if (!ct.includes('json') || !response.url().includes('cars.co.za')) return;
      const json = await response.json().catch(() => null);
      if (!json) return;
      const candidates = [
        (json as Record<string,unknown>)['listings'],
        (json as Record<string,unknown>)['results'],
        (json as Record<string,unknown>)['data'],
        Array.isArray(json) ? json : null,
      ];
      for (const c of candidates) {
        if (Array.isArray(c) && c.length > 0 && typeof c[0] === 'object') {
          capturedVehicles.push(...(c as Array<Record<string,unknown>>));
          console.log('[CarsCoza] JSON API captured', c.length, 'items');
          break;
        }
      }
    } catch { /* ignore */ }
  };
  page.on('response', onResponse);

  try {
    await page.setViewportSize({ width: 1280, height: 900 });
    await page.setExtraHTTPHeaders(BROWSER_HEADERS);

    const { make, model } = normalizeMake(query);
    const params = new URLSearchParams({ cat_make: make });
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
      console.log('[CarsCoza] Blocked — skipping');
      return [];
    }

    await page.waitForLoadState('networkidle').catch(() => null);
    await page.waitForTimeout(3000);
    await page.evaluate(() => window.scrollBy(0, 800));
    await page.waitForTimeout(2000);

    console.log('[CarsCoza] Page title:', await page.title());

    console.log('[CarsCoza] Captured XHR vehicles:', capturedVehicles.length);

    // Strategy 0: use captured XHR/fetch JSON API responses
    // (Cars.co.za may fetch search results client-side after hydration)
    if (capturedVehicles.length > 0) {
      const listings = extractFromVehicleArray(capturedVehicles, 'xhrapi');
      if (listings.length > 0) {
        console.log('[CarsCoza] Listings from XHR API:', listings.length);
        return listings;
      }
    }

    // Strategy 1: Extract from window.__NEXT_DATA__ (Next.js SSR payload)
    // Use numeric heuristics: price 50k–10M ZAR, year 2000–2030
    const nextDataResult = await page.evaluate(() => {
      try {
        const el = document.getElementById('__NEXT_DATA__');
        if (!el) return { found: false, topKeys: [] as string[], ppKeys: [] as string[], listings: null as unknown[] | null };
        const json = JSON.parse(el.textContent ?? '{}') as Record<string, unknown>;
        const topKeys = Object.keys(json);
        const ppKeys = json['props'] && typeof json['props'] === 'object'
          ? Object.keys((json['props'] as Record<string, unknown>)['pageProps'] as Record<string, unknown> ?? {})
          : [];

        function isVehicleArray(arr: unknown[]): boolean {
          if (arr.length < 2) return false;
          const obj = arr[0] as Record<string, unknown>;
          const vals = Object.values(obj);
          const hasPrice = vals.some(v => typeof v === 'number' && v > 50000 && v < 10000000);
          const hasYear = vals.some(v =>
            (typeof v === 'number' && v >= 2000 && v <= 2030) ||
            (typeof v === 'string' && /^20\d{2}$/.test(v as string))
          );
          return hasPrice && hasYear;
        }

        function findListings(obj: unknown, depth = 0): unknown[] | null {
          if (depth > 14 || !obj || typeof obj !== 'object') return null;
          if (Array.isArray(obj)) {
            return isVehicleArray(obj) ? obj : null;
          }
          const rec = obj as Record<string, unknown>;
          // Check high-priority keys first
          for (const k of ['vehicles', 'listings', 'results', 'items', 'cars', 'ads', 'adverts', 'data', 'records', 'stock']) {
            if (rec[k]) {
              const found = findListings(rec[k], depth + 1);
              if (found) return found;
            }
          }
          for (const v of Object.values(rec)) {
            const found = findListings(v, depth + 1);
            if (found) return found;
          }
          return null;
        }

        const listings = findListings(json);
        return { found: true, topKeys, ppKeys, listings, size: listings?.length ?? 0 };
      } catch (e) {
        return { found: false, topKeys: [] as string[], ppKeys: [] as string[], listings: null as unknown[] | null };
      }
    });

    console.log('[CarsCoza] __NEXT_DATA__ found:', nextDataResult.found, '| topKeys:', nextDataResult.topKeys.join(','), '| ppKeys:', nextDataResult.ppKeys.join(','), '| array size:', nextDataResult.size);

    if (nextDataResult.listings && nextDataResult.listings.length > 0) {
      const listings = extractFromVehicleArray(nextDataResult.listings as Array<Record<string, unknown>>, 'nextdata');
      if (listings.length > 0) {
        console.log('[CarsCoza] Listings from __NEXT_DATA__:', listings.length);
        return listings;
      }
    }

    // Strategy 2: parse listing links — URL slug contains year/make/model/location
    const rawLinks = await page.evaluate(() => {
      const links = Array.from(document.querySelectorAll('a[href*="/for-sale/used/"]')) as HTMLAnchorElement[];
      const seen = new Set<string>();
      return links
        .filter(a => { if (seen.has(a.href)) return false; seen.add(a.href); return true; })
        .slice(0, 40)
        .map(a => {
          let container: Element | null = a.parentElement;
          for (let i = 0; i < 8; i++) {
            if (!container) break;
            const text = container.textContent ?? '';
            if (text.includes('R ') && text.length > 30 && text.length < 3000) break;
            container = container.parentElement;
          }
          const leaves = container
            ? Array.from(container.querySelectorAll('*')).filter(
                (el): el is HTMLElement => el.children.length === 0 && !!(el as HTMLElement).innerText?.trim()
              )
            : [];
          const priceEl = leaves.find(el => /R\s?\d/.test(el.innerText));
          const kmEl = leaves.find(el => /\b\d[\d ,]*\s*km\b/i.test(el.innerText));
          // Pick a car photo: prefer CDN/imgix/non-logo images, skip small icons
          const imgs = Array.from(container?.querySelectorAll('img') ?? []) as HTMLImageElement[];
          const photoImg = imgs.find(img => {
            const src = img.src ?? '';
            return src && !src.includes('logo') && !src.includes('icon') && !src.includes('sprite') &&
              (img.naturalWidth === 0 || img.naturalWidth > 100); // loaded or large
          });
          const cardText = container?.textContent ?? '';
          return {
            href: a.href,
            priceText: priceEl?.innerText?.trim() ?? cardText.match(/R\s?[\d,]+/)?.[0] ?? '',
            mileageText: kmEl?.innerText?.trim() ?? cardText.match(/\b(\d{1,3}(?:[, ]\d{3})?)\s*km\b/i)?.[0] ?? '',
            hasServiceHistory: /\b(full service|fsh|service history)\b/i.test(cardText),
            imageUrl: photoImg?.src ?? photoImg?.dataset?.src ?? '',
          };
        });
    });

    console.log('[CarsCoza] Listing links found:', rawLinks.length);

    const listings: Listing[] = [];
    const seenIds = new Set<string>();

    for (const r of rawLinks) {
      const slugMatch = r.href.match(/\/for-sale\/used\/([^/]+)\/(\d+)/);
      if (!slugMatch) continue;
      const parsed = parseSlug(slugMatch[1]);
      if (!parsed.year || !parsed.make) continue;
      const title = `${parsed.year} ${parsed.make} ${parsed.model} ${parsed.variant}`.trim();
      const listing = normalizeRaw(
        { title, priceText: r.priceText, mileageText: r.mileageText, locationText: [parsed.city, parsed.province].filter(Boolean).join(', '), url: r.href, serviceHistory: r.hasServiceHistory },
        'carscoza',
      );
      if (listing && !seenIds.has(listing.id)) {
        listing.make = parsed.make; listing.model = parsed.model; listing.variant = parsed.variant; listing.year = parsed.year;
        if (parsed.city) listing.city = parsed.city;
        if (parsed.province) listing.province = parsed.province;
        if (r.imageUrl) listing.imageUrl = r.imageUrl;
        seenIds.add(listing.id);
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

export { parsePrice, parseMileage };
