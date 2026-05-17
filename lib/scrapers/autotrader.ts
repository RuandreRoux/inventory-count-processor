/**
 * AutoTrader SA scraper
 */

import type { Page, Browser } from 'playwright-core';
import type { Listing } from '@/lib/types';
import { BROWSER_HEADERS } from './browser';
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
    await page.waitForTimeout(2000);

    const pageTitle = await page.title();
    console.log('[AutoTrader] Page title:', pageTitle);

    await dismissOverlays(page);

    // Try JSON-LD structured data first
    const jsonLdListings = await extractFromJsonLd(page, 'autotrader');
    if (jsonLdListings.length > 0) {
      console.log('[AutoTrader] JSON-LD listings found:', jsonLdListings.length);
      return jsonLdListings;
    }

    // Fallback: find listing cards by URL pattern — links to /ad/ pages
    const raw = await page.evaluate((): PageExtract[] => {
      // Strategy 1: find anchor tags linking to /ad/ pages
      const adLinks = Array.from(document.querySelectorAll('a[href*="/ad/"]'));
      console.log('AutoTrader ad links found:', adLinks.length);

      // Group by closest repeated container
      const seen = new Set<Element>();
      const results: PageExtract[] = [];

      for (const link of adLinks.slice(0, 30)) {
        // Walk up to find a container that looks like a card (has price + title info)
        let container: Element | null = link.parentElement;
        for (let i = 0; i < 6; i++) {
          if (!container) break;
          const text = container.textContent ?? '';
          // Container with "R " price text is likely a listing card
          if (text.includes('R ') && text.length > 30 && text.length < 2000) {
            break;
          }
          container = container.parentElement;
        }
        if (!container || seen.has(container)) continue;
        seen.add(container);

        const fullText = container.textContent ?? '';
        const href = link.getAttribute('href') ?? '';
        const absUrl = href.startsWith('http') ? href : `https://www.autotrader.co.za${href}`;

        // Extract title from h2/h3 or link text
        const heading = container.querySelector('h1, h2, h3, h4');
        const title = heading?.textContent?.trim() ?? link.textContent?.trim() ?? '';

        // Extract price — find text matching R + digits
        const priceMatch = fullText.match(/R\s?[\d\s,]+/);
        const priceText = priceMatch ? priceMatch[0].trim() : '';

        // Extract mileage — find text matching digits + km
        const kmMatch = fullText.match(/[\d\s,]+\s*km/i);
        const mileageText = kmMatch ? kmMatch[0].trim() : '';

        // Extract year — 4-digit year
        const yearMatch = fullText.match(/\b(19|20)\d{2}\b/);
        const yearText = yearMatch ? yearMatch[0] : '';

        // Location — look for province names
        const provinces = ['gauteng', 'western cape', 'kwazulu-natal', 'eastern cape', 'limpopo', 'mpumalanga', 'north west', 'free state', 'northern cape'];
        const lowerText = fullText.toLowerCase();
        const matchedProvince = provinces.find(p => lowerText.includes(p)) ?? '';

        // Service history
        const hasServiceHistory = /\b(full service|fsh|service history)\b/i.test(fullText);

        if (title && priceText) {
          results.push({
            title: `${yearText} ${title}`.trim(),
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

    console.log('[AutoTrader] Raw extracts:', raw.length);

    // Log debug info about the page structure
    const debugInfo = await page.evaluate(() => {
      const bodyClasses = document.body.className?.slice(0, 200) ?? '';
      const adLinkCount = document.querySelectorAll('a[href*="/ad/"]').length;
      const h2Count = document.querySelectorAll('h2').length;
      const h3Count = document.querySelectorAll('h3').length;
      const priceEls = document.querySelectorAll('[class*="price"], [class*="Price"]').length;
      return { bodyClasses, adLinkCount, h2Count, h3Count, priceEls };
    });
    console.log('[AutoTrader] Debug:', JSON.stringify(debugInfo));

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

async function extractFromJsonLd(page: Page, source: 'autotrader' | 'carscoza'): Promise<Listing[]> {
  try {
    const jsonLdData = await page.evaluate(() => {
      const scripts = Array.from(document.querySelectorAll('script[type="application/ld+json"]'));
      return scripts.map(s => {
        try { return JSON.parse(s.textContent ?? '{}'); } catch { return null; }
      }).filter(Boolean);
    });

    const listings: Listing[] = [];
    for (const data of jsonLdData) {
      const items = Array.isArray(data['@graph']) ? data['@graph'] :
                    Array.isArray(data) ? data :
                    data['@type'] ? [data] : [];

      for (const item of items) {
        if (!item || !['Car', 'Vehicle', 'Product', 'Offer'].includes(item['@type'])) continue;
        const name = item.name ?? '';
        const price = item.offers?.price ?? item.price ?? 0;
        const mileage = item.mileageFromOdometer?.value ?? item.vehicleSpecialUsage ?? 0;
        const url = item.url ?? '';
        if (!name || !price) continue;

        const normalized = normalizeRaw(
          {
            title: name,
            priceText: `R ${price}`,
            mileageText: `${mileage} km`,
            locationText: item.areaServed ?? item.address?.addressRegion ?? '',
            url,
            serviceHistory: false,
          },
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

async function dismissOverlays(page: Page): Promise<void> {
  const cookieBtn = await page.$('button[id*="accept"], button[class*="accept"], [aria-label*="accept"]');
  if (cookieBtn) await cookieBtn.click().catch(() => null);
  await page.keyboard.press('Escape').catch(() => null);
}

export { parsePrice, parseMileage, parseYear };
