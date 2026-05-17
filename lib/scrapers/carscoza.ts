/**
 * Cars.co.za scraper
 */

import type { Browser } from 'playwright-core';
import type { Listing } from '@/lib/types';
import { BROWSER_HEADERS } from './browser';
import { normalizeRaw } from './normalize';

const BASE_URL = 'https://www.cars.co.za/used-cars-for-sale';

const SEL = {
  card: [
    '[class*="vehicle-card"]',
    '[class*="listing-item"]',
    '[class*="car-item"]',
    '[class*="result-item"]',
    'article[class*="car"]',
    '.vehicle-result',
    '[class*="VehicleCard"]',
    '[class*="SearchResult"]',
    'li[class*="vehicle"]',
  ].join(', '),

  title: 'h2, h3, [class*="title"], [class*="heading"], [class*="Title"]',
  price: '[class*="price"], [class*="Price"]',
  mileage: '[class*="mileage"], [class*="km"], [class*="odometer"], [class*="Mileage"]',
  location: '[class*="location"], [class*="city"], [class*="area"], [class*="Location"]',
  link: 'a[href*="/used-car/"], a[href*="/car/"], a[href]',
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
    const urlWithSlug = parts.length >= 2
      ? `${BASE_URL}/${parts[0]}/${parts.slice(1).join('-')}/`
      : `${BASE_URL}/${slug}/`;

    const params = new URLSearchParams();
    if (filters.maxPrice) params.set('price_to', String(filters.maxPrice));
    if (filters.maxMileage) params.set('mileage_to', String(filters.maxMileage));
    if (filters.minYear) params.set('year_from', String(filters.minYear));

    const qs = params.toString();
    const url = qs ? `${urlWithSlug}?${qs}` : urlWithSlug;

    console.log('[CarsCoza] Navigating to:', url);
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });

    const pageTitle = await page.title();
    console.log('[CarsCoza] Page title:', pageTitle);

    await page.$('button[id*="accept"], button[class*="accept"]').then((b) => b?.click()).catch(() => null);
    await page.keyboard.press('Escape').catch(() => null);
    await page.waitForSelector(SEL.card, { timeout: 15000 }).catch(() => null);

    const debug = await page.evaluate((sel) => {
      const cards = Array.from(document.querySelectorAll(sel.card));
      const bodyClasses = document.body.className;
      const firstClasses = document.querySelector('[class]')?.className?.slice(0, 200) ?? '';
      return { cardCount: cards.length, bodyClasses: bodyClasses.slice(0, 200), firstClasses };
    }, SEL);
    console.log('[CarsCoza] Cards found:', debug.cardCount);
    console.log('[CarsCoza] Body classes:', debug.bodyClasses);
    console.log('[CarsCoza] First element classes:', debug.firstClasses);

    const raw = await page.evaluate((sel: typeof SEL) => {
      const cards = Array.from(document.querySelectorAll(sel.card));
      return cards.slice(0, 24).map((card) => {
        const el = (s: string) => card.querySelector(s);
        const title = el(sel.title)?.textContent?.trim() ?? '';
        const priceText = el(sel.price)?.textContent?.trim() ?? '';
        const mileageText = el(sel.mileage)?.textContent?.trim() ?? '';
        const locationText = el(sel.location)?.textContent?.trim() ?? '';
        const href = el(sel.link)?.getAttribute('href') ?? '';
        const listingUrl = href.startsWith('http') ? href : `https://www.cars.co.za${href}`;
        const historyText = el(sel.serviceHistory)?.textContent?.toLowerCase() ?? '';
        const hasServiceHistory = historyText.includes('full') || historyText.includes('fsh');
        return { title, priceText, mileageText, locationText, url: listingUrl, hasServiceHistory };
      });
    }, SEL);

    const listings: Listing[] = [];
    for (const r of raw) {
      const normalized = normalizeRaw(
        { title: r.title, priceText: r.priceText, mileageText: r.mileageText, locationText: r.locationText, url: r.url, serviceHistory: r.hasServiceHistory },
        'carscoza',
      );
      if (normalized) listings.push(normalized);
    }
    console.log('[CarsCoza] Listings extracted:', listings.length);
    return listings;
  } finally {
    await page.close();
  }
}
