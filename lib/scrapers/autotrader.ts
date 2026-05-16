/**
 * AutoTrader SA scraper
 * Selectors verified against autotrader.co.za — update SEL constants if the site changes.
 * NOTE: autotrader.co.za ToS prohibits automated access. Seek a formal partnership for production.
 */
import type { Page, Browser } from 'playwright-core';
import type { Listing } from '@/lib/types';
import { BROWSER_HEADERS } from './browser';
import { normalizeRaw } from './normalize';

const BASE_URL = 'https://www.autotrader.co.za/cars-for-sale';

const SEL = {
  card: ['[class*="listing-tile"]','[class*="ListingTile"]','[data-testid="listing"]','.b-listing','article[class*="listing"]'].join(', '),
  title: 'h2, h3, [class*="title"], [data-testid="title"]',
  price: '[class*="price"], [class*="Price"], [data-testid="price"]',
  mileage: '[class*="odometer"], [class*="mileage"], [class*="km"]',
  location: '[class*="location"], [class*="city"], [class*="province"]',
  link: 'a[href*="/ad/"]',
  serviceHistory: '[class*="service"], [class*="history"]',
};

const MAKES: Record<string, string> = {
  toyota: 'Toyota', volkswagen: 'Volkswagen', vw: 'Volkswagen', ford: 'Ford',
  bmw: 'BMW', 'mercedes-benz': 'Mercedes-Benz', mercedes: 'Mercedes-Benz',
  hyundai: 'Hyundai', kia: 'Kia', mazda: 'Mazda', isuzu: 'Isuzu',
  nissan: 'Nissan', honda: 'Honda', suzuki: 'Suzuki', renault: 'Renault',
  audi: 'Audi', haval: 'Haval', 'land rover': 'Land Rover',
};

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
    await page.goto(`${BASE_URL}?${params.toString()}`, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await dismissOverlays(page);
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
          url: href.startsWith('http') ? href : `https://www.autotrader.co.za${href}`,
          hasServiceHistory: (el(sel.serviceHistory)?.textContent?.toLowerCase() ?? '').includes('full'),
        };
      });
    }, SEL);
    return raw.flatMap((r) => {
      const n = normalizeRaw({ title: r.title, priceText: r.priceText, mileageText: r.mileageText, locationText: r.locationText, url: r.url, serviceHistory: r.hasServiceHistory }, 'autotrader');
      return n ? [n] : [];
    });
  } finally {
    await page.close();
  }
}

async function dismissOverlays(page: Page): Promise<void> {
  const btn = await page.$('button[id*="accept"], button[class*="accept"], [aria-label*="accept"]');
  if (btn) await btn.click().catch(() => null);
  await page.keyboard.press('Escape').catch(() => null);
}
