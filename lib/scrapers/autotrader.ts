/**
 * AutoTrader SA scraper
 */

import type { Browser } from 'playwright-core';
import type { Listing } from '@/lib/types';
import { normalizeRaw, parsePrice, parseMileage, parseYear } from './normalize';

const BASE_URL = 'https://www.autotrader.co.za/cars-for-sale';

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
    // Full realistic browser context to avoid bot detection
    await page.setViewportSize({ width: 1280, height: 900 });
    await page.setExtraHTTPHeaders({
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
      'Accept-Language': 'en-ZA,en-GB;q=0.9,en;q=0.8',
      'Accept-Encoding': 'gzip, deflate, br',
      'Upgrade-Insecure-Requests': '1',
      'Sec-Fetch-Dest': 'document',
      'Sec-Fetch-Mode': 'navigate',
      'Sec-Fetch-Site': 'none',
      'Sec-Fetch-User': '?1',
      'Cache-Control': 'max-age=0',
    });

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

    const response = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    const status = response?.status() ?? 0;
    console.log('[AutoTrader] HTTP status:', status);

    // Wait for JavaScript to render content
    await page.waitForLoadState('networkidle').catch(() => null);
    await page.waitForTimeout(3000);

    const pageTitle = await page.title();
    console.log('[AutoTrader] Page title:', pageTitle);
    const bodyText = await page.evaluate(() => document.body?.innerText?.slice(0, 300) ?? '');
    console.log('[AutoTrader] Body text preview:', bodyText);

    // Try JSON-LD structured data first
    const jsonLdListings = await extractFromJsonLd(page, 'autotrader');
    if (jsonLdListings.length > 0) {
      console.log('[AutoTrader] JSON-LD listings:', jsonLdListings.length);
      return jsonLdListings;
    }

    // DOM-based extraction — find listing links and walk up to their container
    const raw = await page.evaluate((): PageExtract[] => {
      const adLinks = Array.from(document.querySelectorAll('a[href*="/ad/"], a[href*="/used/"]'));
      console.log('AutoTrader ad links:', adLinks.length);

      const seen = new Set<Element>();
      const results: PageExtract[] = [];

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
        const provinces = ['gauteng', 'western cape', 'kwazulu-natal', 'eastern cape', 'limpopo', 'mpumalanga', 'north west', 'free state', 'northern cape'];
        const matchedProvince = provinces.find(p => (fullText.toLowerCase()).includes(p)) ?? '';
        const hasServiceHistory = /\b(full service|fsh|service history)\b/i.test(fullText);

        if (title && priceText) {
          results.push({
            title: yearMatch ? `${yearMatch[0]} ${title}` : title,
            priceText,
            mileageText,
            locationText: matchedProvince,
            url: absUrl,
            hasServiceHistory,
          });
        }
      }
      return results;
    });

    const debug = await page.evaluate(() => ({
      adLinkCount: document.querySelectorAll('a[href*="/ad/"]').length,
      usedLinkCount: document.querySelectorAll('a[href*="/used/"]').length,
      h2Count: document.querySelectorAll('h2').length,
      bodyLen: document.body?.innerHTML?.length ?? 0,
    }));
    console.log('[AutoTrader] Debug:', JSON.stringify(debug));
    console.log('[AutoTrader] Raw extracts:', raw.length);

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

async function extractFromJsonLd(page: import('playwright-core').Page, source: 'autotrader' | 'carscoza'): Promise<Listing[]> {
  try {
    const jsonLdData = await page.evaluate(() => {
      return Array.from(document.querySelectorAll('script[type="application/ld+json"]'))
        .map(s => { try { return JSON.parse(s.textContent ?? '{}'); } catch { return null; } })
        .filter(Boolean);
    });

    const listings: Listing[] = [];
    for (const data of jsonLdData) {
      const items: unknown[] = Array.isArray(data['@graph']) ? data['@graph'] :
                    Array.isArray(data) ? data :
                    data['@type'] ? [data] : [];
      for (const raw of items) {
        const item = raw as Record<string, unknown>;
        if (!['Car', 'Vehicle', 'Product'].includes(item['@type'] as string)) continue;
        const name = item.name as string ?? '';
        const offerObj = item.offers as Record<string, unknown> | undefined;
        const price = (offerObj?.price ?? item.price ?? 0) as number;
        const mileageObj = item.mileageFromOdometer as Record<string, unknown> | undefined;
        const mileage = (mileageObj?.value ?? 0) as number;
        const url = (item.url as string) ?? '';
        if (!name || !price) continue;
        const normalized = normalizeRaw(
          { title: name, priceText: `R ${price}`, mileageText: `${mileage} km`, locationText: '', url, serviceHistory: false },
          source,
        );
        if (normalized) listings.push(normalized);
      }
    }
    return listings;
  } catch {
    return [];
  }
}

export { parsePrice, parseMileage, parseYear };
