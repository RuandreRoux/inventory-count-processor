/**
 * Cars.co.za scraper
 */

import type { Page, Browser } from 'playwright-core';
import type { Listing } from '@/lib/types';
import { BROWSER_HEADERS } from './browser';
import { normalizeRaw } from './normalize';

const BASE_URL = 'https://www.cars.co.za/used-cars-for-sale';

interface PageExtract {
  title: string;
  priceText: string;
  mileageText: string;
  locationText: string;
  url: string;
  hasServiceHistory: boolean;
}

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
    await page.waitForTimeout(2000);

    const pageTitle = await page.title();
    console.log('[CarsCoza] Page title:', pageTitle);

    await page.$('button[id*="accept"], button[class*="accept"]').then((b) => b?.click()).catch(() => null);
    await page.keyboard.press('Escape').catch(() => null);

    // Try JSON-LD first
    const jsonLdListings = await extractFromJsonLd(page, 'carscoza');
    if (jsonLdListings.length > 0) {
      console.log('[CarsCoza] JSON-LD listings found:', jsonLdListings.length);
      return jsonLdListings;
    }

    // Fallback: DOM-based extraction using URL patterns
    const raw = await page.evaluate((): PageExtract[] => {
      // Cars.co.za links typically contain /used-car/ or /car/
      const adLinks = Array.from(document.querySelectorAll('a[href*="/used-car/"], a[href*="/car/"]'))
        .filter(a => {
          const href = (a as HTMLAnchorElement).href ?? '';
          // Filter out navigation/make/model links — listing URLs typically have at least 3 path segments
          return href.split('/').length >= 6;
        });

      console.log('CarsCoza ad links found:', adLinks.length);

      const seen = new Set<Element>();
      const results: PageExtract[] = [];

      for (const link of adLinks.slice(0, 30)) {
        // Walk up to find listing container
        let container: Element | null = link.parentElement;
        for (let i = 0; i < 6; i++) {
          if (!container) break;
          const text = container.textContent ?? '';
          if (text.includes('R ') && text.length > 30 && text.length < 2000) break;
          container = container.parentElement;
        }
        if (!container || seen.has(container)) continue;
        seen.add(container);

        const fullText = container.textContent ?? '';
        const href = (link as HTMLAnchorElement).href ?? link.getAttribute('href') ?? '';
        const absUrl = href.startsWith('http') ? href : `https://www.cars.co.za${href}`;

        const heading = container.querySelector('h1, h2, h3, h4');
        const title = heading?.textContent?.trim() ?? link.textContent?.trim() ?? '';

        const priceMatch = fullText.match(/R\s?[\d\s,]+/);
        const priceText = priceMatch ? priceMatch[0].trim() : '';

        const kmMatch = fullText.match(/[\d\s,]+\s*km/i);
        const mileageText = kmMatch ? kmMatch[0].trim() : '';

        const yearMatch = fullText.match(/\b(19|20)\d{2}\b/);
        const yearText = yearMatch ? yearMatch[0] : '';

        const provinces = ['gauteng', 'western cape', 'kwazulu-natal', 'eastern cape', 'limpopo', 'mpumalanga', 'north west', 'free state', 'northern cape'];
        const lowerText = fullText.toLowerCase();
        const matchedProvince = provinces.find(p => lowerText.includes(p)) ?? '';

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

    console.log('[CarsCoza] Raw extracts:', raw.length);

    const debugInfo = await page.evaluate(() => {
      const bodyClasses = document.body.className?.slice(0, 200) ?? '';
      const carLinks = document.querySelectorAll('a[href*="/used-car/"], a[href*="/car/"]').length;
      const h2Count = document.querySelectorAll('h2').length;
      const h3Count = document.querySelectorAll('h3').length;
      return { bodyClasses, carLinks, h2Count, h3Count };
    });
    console.log('[CarsCoza] Debug:', JSON.stringify(debugInfo));

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
        const mileage = item.mileageFromOdometer?.value ?? 0;
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
