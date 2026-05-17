/**
 * AutoTrader SA scraper.
 *
 * autotrader.co.za blocks data-center IPs with HTTP 503.
 * Set SCRAPERAPI_KEY env var to route through residential proxies and bypass the block.
 * Without the key, this scraper returns [] on a 503.
 */

import type { Browser, BrowserContext } from 'playwright-core';
import type { Listing } from '@/lib/types';
import { BROWSER_HEADERS, newProxiedContext } from './browser';
import { normalizeRaw, parsePrice, parseMileage, parseYear } from './normalize';

const BASE_URL = 'https://www.autotrader.co.za/cars-for-sale';

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
  // Use a proxied context if ScraperAPI key is available, else use plain browser
  let context: BrowserContext | null = await newProxiedContext(browser);
  const page = context ? await context.newPage() : await browser.newPage();

  try {
    await page.setViewportSize({ width: 1280, height: 900 });
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

    const url = `${BASE_URL}?${params.toString()}`;
    console.log('[AutoTrader] Navigating to:', url);

    const response = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    const status = response?.status() ?? 0;
    console.log('[AutoTrader] HTTP status:', status);

    if (status === 503 || status === 403) {
      console.log('[AutoTrader] Blocked — set SCRAPERAPI_KEY env var to bypass');
      return [];
    }

    await page.waitForLoadState('networkidle').catch(() => null);
    await page.waitForTimeout(2000);

    const pageTitle = await page.title();
    console.log('[AutoTrader] Page title:', pageTitle);

    const raw = await page.evaluate(() => {
      const adLinks = Array.from(document.querySelectorAll('a[href*="/ad/"], a[href*="/used/"]'));
      const seen = new Set<Element>();
      type Extract = { title: string; priceText: string; mileageText: string; locationText: string; url: string; hasServiceHistory: boolean };
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
        const href = link.getAttribute('href') ?? '';
        const absUrl = href.startsWith('http') ? href : `https://www.autotrader.co.za${href}`;
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

    const debug = await page.evaluate(() => ({
      adLinks: document.querySelectorAll('a[href*="/ad/"]').length,
      h2: document.querySelectorAll('h2').length,
      bodyLen: document.body?.innerHTML?.length ?? 0,
    }));
    console.log('[AutoTrader] Debug:', JSON.stringify(debug));
    console.log('[AutoTrader] Raw extracts:', raw.length);

    const listings: Listing[] = [];
    for (const r of raw) {
      const n = normalizeRaw(
        { title: r.title, priceText: r.priceText, mileageText: r.mileageText, locationText: r.locationText, url: r.url, serviceHistory: r.hasServiceHistory },
        'autotrader',
      );
      if (n) listings.push(n);
    }
    console.log('[AutoTrader] Listings extracted:', listings.length);
    return listings;
  } finally {
    if (context) {
      await context.close();
    } else {
      await page.close();
    }
  }
}

export { parsePrice, parseMileage, parseYear };
