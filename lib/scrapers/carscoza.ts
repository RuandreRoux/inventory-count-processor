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

    // Primary: parse listing links — URL slug has all the car info we need
    const rawLinks = await page.evaluate(() => {
      const links = Array.from(document.querySelectorAll('a[href*="/for-sale/used/"]')) as HTMLAnchorElement[];
      // De-dupe by href
      const seen = new Set<string>();
      return links
        .filter(a => {
          if (seen.has(a.href)) return false;
          seen.add(a.href);
          return true;
        })
        .slice(0, 40)
        .map(a => {
          // Walk up to find the listing card container
          let container: Element | null = a.parentElement;
          for (let i = 0; i < 8; i++) {
            if (!container) break;
            const text = container.textContent ?? '';
            if (text.includes('R ') && text.length > 30 && text.length < 3000) break;
            container = container.parentElement;
          }

          // Use leaf elements for precise extraction — avoids cross-line regex issues
          const leaves = container
            ? Array.from(container.querySelectorAll('*')).filter(
                (el): el is HTMLElement => el.children.length === 0 && !!(el as HTMLElement).innerText?.trim()
              )
            : [];

          const priceEl = leaves.find(el => /R\s?\d/.test(el.innerText));
          const kmEl = leaves.find(el => /\b\d[\d ,]*\s*km\b/i.test(el.innerText));
          const imgEl = container?.querySelector('img[src*="cars.co.za"], img[src*="imgix"], img[data-src], img[src]') as HTMLImageElement | null;

          const cardText = container?.textContent ?? '';
          const hasServiceHistory = /\b(full service|fsh|service history)\b/i.test(cardText);

          // Fallback regex restricted to standard number formats (no cross-line matching)
          const priceText = priceEl?.innerText?.trim() ?? cardText.match(/R\s?[\d,]+/)?.[0] ?? '';
          const mileageText = kmEl?.innerText?.trim() ??
            cardText.match(/\b(\d{1,3}(?:[, ]\d{3})?)\s*km\b/i)?.[0] ?? '';

          return {
            href: a.href,
            priceText,
            mileageText,
            hasServiceHistory,
            imageUrl: imgEl?.src ?? imgEl?.dataset?.src ?? '',
          };
        });
    });

    console.log('[CarsCoza] Listing links found:', rawLinks.length);

    const listings: Listing[] = [];
    const seenIds = new Set<string>();

    for (const r of rawLinks) {
      // Extract slug from URL: /for-sale/used/{slug}/{id}/
      const slugMatch = r.href.match(/\/for-sale\/used\/([^/]+)\/(\d+)/);
      if (!slugMatch) continue;

      const slug = slugMatch[1];
      const parsed = parseSlug(slug);
      if (!parsed.year || !parsed.make) continue;

      const title = `${parsed.year} ${parsed.make} ${parsed.model} ${parsed.variant}`.trim();
      const location = [parsed.city, parsed.province].filter(Boolean).join(', ');

      const listing = normalizeRaw(
        {
          title,
          priceText: r.priceText,
          mileageText: r.mileageText,
          locationText: location,
          url: r.href,
          serviceHistory: r.hasServiceHistory,
        },
        'carscoza',
      );

      if (listing && !seenIds.has(listing.id)) {
        // Override parsed fields since slug is more reliable than normalizeRaw's title parsing
        listing.make = parsed.make;
        listing.model = parsed.model;
        listing.variant = parsed.variant;
        listing.year = parsed.year;
        if (parsed.city) listing.city = parsed.city;
        if (parsed.province) listing.province = parsed.province;
        if (r.imageUrl) listing.imageUrl = r.imageUrl;

        seenIds.add(listing.id);
        listings.push(listing);
      }
    }

    // Supplement with any JSON API data captured
    if (capturedVehicles.length > 0 && listings.length < 5) {
      console.log('[CarsCoza] Supplementing with', capturedVehicles.length, 'API vehicles');
      for (const v of capturedVehicles.slice(0, 20)) {
        const make = (v['make'] ?? v['Make'] ?? '') as string;
        const model = (v['model'] ?? v['Model'] ?? '') as string;
        const year = Number(v['year'] ?? v['Year'] ?? 0);
        const price = Number(v['price'] ?? v['Price'] ?? 0);
        const mileage = Number(v['mileage'] ?? v['Mileage'] ?? v['km'] ?? 0);
        const url = (v['url'] ?? v['link'] ?? '') as string;
        if (!make || !price) continue;
        const title = `${year} ${make} ${model}`.trim();
        const n = normalizeRaw(
          { title, priceText: `R ${price}`, mileageText: `${mileage} km`, locationText: '', url: url.startsWith('http') ? url : url ? `https://www.cars.co.za${url}` : '', serviceHistory: false },
          'carscoza',
        );
        if (n && !seenIds.has(n.id)) { seenIds.add(n.id); listings.push(n); }
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
