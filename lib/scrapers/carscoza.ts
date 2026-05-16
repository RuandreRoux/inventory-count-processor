/**
 * Cars.co.za scraper
 * Selectors verified against cars.co.za — update SEL constants if the site changes.
 * NOTE: cars.co.za ToS prohibits automated scraping. Seek a formal partnership for production.
 */
import type { Browser } from 'playwright-core';
import type { Listing } from '@/lib/types';
import { BROWSER_HEADERS } from './browser';
import { normalizeRaw } from './normalize';

const BASE_URL = 'https://www.cars.co.za/used-cars-for-sale';

const SEL = {
  card: ['[class*="vehicle-card"]','[class*="listing-item"]','[class*="car-item"]','[class*="result-item"]','.vehicle-result'].join(', '),
  title: 'h2, h3, [class*="title"], [class*="heading"]',
  price: '[class*="price"], [class*="Price"]',
  mileage: '[class*="mileage"], [class*="km"], [class*="odometer"]',
  location: '[class*="location"], [class*="city"], [class*="area"]',
  link: 'a[href*="/used-car/"], a[href*="/car/"]',
  serviceHistory: '[class*="service"], [class*="history"], [class*="fsh"]',
};

function toSlug(s: string): string {
  return s.toLowerCase().replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '');
}

export async function scrapeCarsCoza(
  browser: Browser,
  query: string,
  filters: { maxPrice?: number; maxMileage?: number; minYear?: number },
): Promise<Listing[]> {
  const page = await browser.newPage();
  try {
    await page.setExtraHTTPHeaders(BROWSER_HEADERS);
    const slug = toSlug(query);
    const parts = slug.split('-');
    const slugUrl = parts.length >= 2 ? `${BASE_URL}/${parts[0]}/${parts.slice(1).join('-')}/` : `${BASE_URL}/${slug}/`;
    const params = new URLSearchParams();
    if (filters.maxPrice) params.set('price_to', String(filters.maxPrice));
    if (filters.maxMileage) params.set('mileage_to', String(filters.maxMileage));
    if (filters.minYear) params.set('year_from', String(filters.minYear));
    const qs = params.toString();
    await page.goto(qs ? `${slugUrl}?${qs}` : slugUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.$('button[id*="accept"], button[class*="accept"]').then((b) => b?.click()).catch(() => null);
    await page.keyboard.press('Escape').catch(() => null);
    await page.waitForSelector(SEL.card, { timeout: 15000 }).catch(() => null);
    const raw = await page.evaluate((sel: typeof SEL) => {
      return Array.from(document.querySelectorAll(sel.card)).slice(0, 24).map((card) => {
        const el = (s: string) => card.querySelector(s);
        const href = el(sel.link)?.getAttribute('href') ?? '';
        return {
          title: el(sel.title)?.textContent?.trim() ?? '',
          priceText: el(sel.price)?.textContent?.trim() ?? '',
          mileageText: el(sel.mileage)?.textContent?.trim() ?? '',
          locationText: el(sel.location)?.textContent?.trim() ?? '',
          url: href.startsWith('http') ? href : `https://www.cars.co.za${href}`,
          hasServiceHistory: (el(sel.serviceHistory)?.textContent?.toLowerCase() ?? '').includes('full'),
        };
      });
    }, SEL);
    return raw.flatMap((r) => {
      const n = normalizeRaw({ title: r.title, priceText: r.priceText, mileageText: r.mileageText, locationText: r.locationText, url: r.url, serviceHistory: r.hasServiceHistory }, 'carscoza');
      return n ? [n] : [];
    });
  } finally {
    await page.close();
  }
}
