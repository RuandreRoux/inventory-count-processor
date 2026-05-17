/**
 * AutoTrader SA scraper
 */

import type { Page, Browser } from 'playwright-core';
import type { Listing } from '@/lib/types';
import { BROWSER_HEADERS } from './browser';
import { normalizeRaw, parsePrice, parseMileage, parseYear } from './normalize';

const BASE_URL = 'https://www.autotrader.co.za/cars-for-sale';

const SEL = {
  card: [
    '[class*="listing-tile"]',
    '[class*="ListingTile"]',
    '[data-testid="listing"]',
    '.b-listing',
    'article[class*="listing"]',
    '[class*="result"]',
    '[class*="vehicle"]',
    'li[class*="tile"]',
  ].join(', '),

  title: 'h2, h3, [class*="title"], [data-testid="title"], [class*="heading"]',
  price: '[class*="price"], [class*="Price"], [data-testid="price"]',
  mileage: '[class*="odometer"], [class*="mileage"], [class*="km"]',
  location: '[class*="location"], [class*="city"], [class*="province"]',
  link: 'a[href*="/ad/"], a[href*="/used-car/"], a[href]',
  serviceHistory: '[class*="service"], [class*="history"]',
};

interface PageExtract {
  title: string;
  priceText: string;
  mileageText: string;
  locationText: string;
  url: string;
  hasServiceHistory: boolean;
}

export async function scrapeAutoTrader(
  browser: Browser,
  query: string,
  filters: { maxPrice?: number; maxMileage?: number; minYear?: number },
): Promise<Listing[]> {
  const page = await browser.newPage();
  try {
    await page.setExtraHTTPHeaders(BROWSER_HEADERS);

    const params = new URLSearchParams();
    const q = query.toLowerCase();
    const MAKES: Record<string, string> = {
      toyota: 'Toyota', volkswagen: 'Volkswagen', vw: 'Volkswagen', ford: 'Ford',
      bmw: 'BMW', 'mercedes-benz': 'Mercedes-Benz', mercedes: 'Mercedes-Benz',
      hyundai: 'Hyundai', kia: 'Kia', mazda: 'Mazda', isuzu: 'Isuzu',
      nissan: 'Nissan', honda: 'Honda', suzuki: 'Suzuki', renault: 'Renault',
      audi: 'Audi', haval: 'Haval', 'land rover': 'Land Rover',
    };
    const matchedMake = Object.entries(MAKES).find(([k]) => q.includes(k))?.[1];
    if (matchedMake) {
      params.set('search[car_make]', matchedMake);
      const modelPart = q.replace(matchedMake.toLowerCase(), '').trim();
      if (modelPart) params.set('search[car_model]', modelPart);
    } else {
      params.set('search[model]', query);
    }
    if (filters.maxPrice) params.set('search[price_to]', String(filters.maxPrice));
    if (filters.maxMileage) params.set('search[mileage_to]', String(filters.maxMileage));
    if (filters.minYear) params.set('search[year_from]', String(filters.minYear));

    const url = `${BASE_URL}?${params.toString()}`;
    console.log('[AutoTrader] Navigating to:', url);
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });

    const pageTitle = await page.title();
    console.log('[AutoTrader] Page title:', pageTitle);

    await dismissOverlays(page);
    await page.waitForSelector(SEL.card, { timeout: 15000 }).catch(() => null);

    const debug = await page.evaluate((sel) => {
      const cards = Array.from(document.querySelectorAll(sel.card));
      const bodyClasses = document.body.className;
      const firstClasses = document.querySelector('[class]')?.className?.slice(0, 200) ?? '';
      return { cardCount: cards.length, bodyClasses: bodyClasses.slice(0, 200), firstClasses };
    }, SEL);
    console.log('[AutoTrader] Cards found:', debug.cardCount);
    console.log('[AutoTrader] Body classes:', debug.bodyClasses);
    console.log('[AutoTrader] First element classes:', debug.firstClasses);

    const raw = await page.evaluate((sel: typeof SEL) => {
      const cards = Array.from(document.querySelectorAll(sel.card));
      return cards.slice(0, 24).map((card): PageExtract => {
        const el = (s: string) => card.querySelector(s);
        const title = el(sel.title)?.textContent?.trim() ?? '';
        const priceText = el(sel.price)?.textContent?.trim() ?? '';
        const mileageText = el(sel.mileage)?.textContent?.trim() ?? '';
        const locationText = el(sel.location)?.textContent?.trim() ?? '';
        const href = el(sel.link)?.getAttribute('href') ?? '';
        const url = href.startsWith('http') ? href : `https://www.autotrader.co.za${href}`;
        const historyEl = el(sel.serviceHistory)?.textContent?.toLowerCase() ?? '';
        const hasServiceHistory = historyEl.includes('full') || historyEl.includes('service');
        return { title, priceText, mileageText, locationText, url, hasServiceHistory };
      });
    }, SEL);

    const listings: Listing[] = [];
    for (const r of raw) {
      const normalized = normalizeRaw(
        { title: r.title, priceText: r.priceText, mileageText: r.mileageText, locationText: r.locationText, url: r.url, serviceHistory: r.hasServiceHistory },
        'autotrader',
      );
      if (normalized) listings.push(normalized);
    }
    console.log('[AutoTrader] Listings extracted:', listings.length);
    return listings;
  } finally {
    await page.close();
  }
}

async function dismissOverlays(page: Page): Promise<void> {
  const cookieBtn = await page.$('button[id*="accept"], button[class*="accept"], [aria-label*="accept"]');
  if (cookieBtn) await cookieBtn.click().catch(() => null);
  await page.keyboard.press('Escape').catch(() => null);
}

export { parsePrice, parseMileage, parseYear };
