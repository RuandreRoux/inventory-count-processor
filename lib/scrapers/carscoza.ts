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
    const raw = v[k];
    if (typeof raw === 'number' && raw > 0) return raw;
    // Handle mileage/price stored as a formatted string e.g. "279 504" or "279504"
    if (typeof raw === 'string' && raw.length > 0) {
      const val = parseInt(raw.replace(/[^0-9]/g, ''), 10);
      if (val > 0) return val;
    }
  }
  return 0;
}

function extractFromVehicleArray(raw: Array<Record<string, unknown>>, source: string): Listing[] {
  const listings: Listing[] = [];
  const seenIds = new Set<string>();
  for (const rawV of raw.slice(0, 500)) {
    // Flatten JSON:API attributes into top-level (Cars.co.za uses {id, type, attributes:{...}})
    const attrs = rawV['attributes'] as Record<string, unknown> | undefined;
    const v: Record<string, unknown> = attrs && typeof attrs === 'object' ? { ...attrs, ...rawV } : rawV;

    // Log first item's full key structure to identify field names in Render logs
    if (rawV === raw[0] && source !== 'link') {
      const allKeys = Object.entries(v).map(([k, val]) => {
        if (val === null || val === undefined) return `${k}:null`;
        if (Array.isArray(val)) return `${k}:[${(val as unknown[]).length}]`;
        if (typeof val === 'object') return `${k}:{${Object.keys(val as object).slice(0, 4).join(',')}}`;
        const s = String(val);
        return `${k}=${s.length > 40 ? s.substring(0, 40) + '…' : s}`;
      });
      console.log(`[CarsCoza] First item (${source}) keys (${allKeys.length}): ${allKeys.join(' | ')}`);
      const numFields = Object.entries(v).filter(([,val]) => typeof val === 'number').map(([k,val]) => `${k}=${val}`);
      const strFields = Object.entries(v).filter(([,val]) => typeof val === 'string' && (val as string).length < 80).map(([k,val]) => `${k}="${val}"`);
      console.log(`[CarsCoza] First item (${source}) num fields: ${numFields.join(', ')}`);
      console.log(`[CarsCoza] First item (${source}) str fields: ${strFields.join(', ')}`);
    }

    const make = pickStr(v, 'make', 'Make', 'manufacturer', 'makeDescription', 'makeName', 'brand', 'make_description');
    const model = pickStr(v, 'model', 'Model', 'modelDescription', 'modelName', 'model_description');
    const variant = pickStr(v, 'variant', 'Variant', 'derivative', 'trim', 'variantDescription', 'description', 'variant_description', 'title');
    const year = pickNum(v, 'year', 'Year', 'modelYear', 'vehicleYear', 'model_year', 'vehicle_year');
    const price = pickNum(v, 'price', 'Price', 'sellingPrice', 'askingPrice', 'listPrice', 'vehiclePrice', 'retail', 'selling_price', 'asking_price');
    const mileage = pickNum(v,
      'mileage', 'Mileage', 'km', 'odometer', 'kilometres', 'Kilometres',
      'kilometers', 'kms', 'Kms', 'vehicleKilometres', 'vehicle_kilometres',
      'totalKm', 'total_km', 'kmReading', 'km_reading', 'odometerReading',
      'odometer_reading', 'kmDriven', 'km_driven', 'usedKm', 'used_km',
      'vehicle_mileage', 'distance', 'milage', 'mileage_value',
      'odometerKm', 'odometer_km', 'listedKm', 'listed_km',
    );

    // Only use links.self if it's a real Cars.co.za listing page URL, not an internal API URL
    const linksObj = (rawV['links'] ?? (v['links'])) as Record<string, unknown> | undefined;
    const linksUrl = linksObj && typeof linksObj === 'object' ? (typeof linksObj['self'] === 'string' ? linksObj['self'] as string : '') : '';
    const isPublicListingUrl = (url: string) => url.includes('cars.co.za/for-sale/') || url.startsWith('/for-sale/used/');
    const rawUrl = pickStr(v, 'url', 'link', 'listingUrl', 'permalink', 'adUrl', 'href', 'detailUrl', 'slug', 'detail_url', 'listing_url', 'ad_url');

    // JSON:API images may be nested: images[0].url or photos[0].url
    let imageUrl = pickStr(v,
      'image', 'imageUrl', 'thumbnail', 'photo', 'primaryImage', 'mainImage', 'heroImage',
      'imgUrl', 'picture', 'featured_image_url', 'primary_image_url', 'image_url', 'photo_url',
      'img', 'thumbnailUrl', 'leadImage', 'vehicleImage', 'mainPhoto', 'leadPhoto',
      'primaryPhoto', 'featuredImage', 'carImage', 'listingImage', 'coverImage',
    );
    if (!imageUrl) {
      const imgArr = (
        v['images'] ?? v['photos'] ?? v['media'] ?? v['gallery'] ??
        v['vehicleImages'] ?? v['carsImages'] ?? v['listingImages'] ?? v['photoList']
      ) as Array<Record<string, unknown>> | undefined;
      if (Array.isArray(imgArr) && imgArr.length > 0) {
        imageUrl = pickStr(imgArr[0], 'url', 'src', 'href', 'uri', 'path', 'original', 'large', 'medium', 'small', 'thumbnail', 'imageUrl', 'fullUrl');
      }
    }

    const cityText = pickStr(v, 'city', 'town', 'suburb', 'area', 'dealer_city', 'dealerCity', 'location_city', 'locationCity');
    const provinceText = pickStr(v, 'province', 'region', 'dealer_province', 'dealerProvince', 'location_province', 'locationProvince');
    const locationText = [cityText, provinceText].filter(Boolean).join(', ')
      || pickStr(v, 'location', 'address');

    // Build URL: prefer a real Cars.co.za listing URL > fallback construction from id
    const listingId = String(rawV['id'] ?? v['id'] ?? '');
    let finalUrl = (linksUrl && isPublicListingUrl(linksUrl)) ? linksUrl
                 : (rawUrl && isPublicListingUrl(rawUrl)) ? rawUrl
                 : '';
    if (!finalUrl && listingId) {
      // Construct a valid Cars.co.za URL — the id alone is sufficient to find the listing
      const slug = `${year}-${make}-${model}`
        .toLowerCase()
        .replace(/\s+/g, '-')
        .replace(/[^a-z0-9-]/g, '');
      finalUrl = `/for-sale/used/${slug}/${listingId}/`;
    }

    // Construct image URL from CDN pattern if no image was found in the API response
    // Pattern: https://img-ik.cars.co.za/ik-seo/carsimages/{id}/{year}-{Make}-{Model}-{Variant}.jpg
    if (!imageUrl && listingId && year && make && model) {
      const imgSlug = `${year}-${make}-${model}${variant ? '-' + variant : ''}`
        .replace(/\./g, '')
        .replace(/\s+/g, '-')
        .replace(/[^a-zA-Z0-9-]/g, '');
      imageUrl = `https://img-ik.cars.co.za/ik-seo/carsimages/${listingId}/${imgSlug}.jpg?tr=f-auto,h-267,w-400,q-80`;
    }

    if (!make || !price || !year) {
      if (source !== 'link') console.log(`[CarsCoza] Skip (${source}): make=${make} price=${price} year=${year} keys=${Object.keys(v).slice(0, 12).join(',')}`);
      continue;
    }
    const title = `${year} ${make} ${model} ${variant}`.trim();
    const listing = normalizeRaw(
      {
        title,
        priceText: `R ${price}`,
        mileageText: `${mileage} km`,
        locationText,
        url: finalUrl.startsWith('http') ? finalUrl : finalUrl ? `https://www.cars.co.za${finalUrl}` : '',
        serviceHistory: Boolean(v['serviceHistory'] ?? v['fsh'] ?? v['fullServiceHistory'] ?? v['full_service_history'] ?? false),
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

function isVehicleArray(arr: unknown[]): boolean {
  if (arr.length < 2) return false;
  const obj = arr[0] as Record<string, unknown>;

  function hasVehicleFields(o: Record<string, unknown>): boolean {
    const vals = Object.values(o);
    const hasPrice = vals.some(v => typeof v === 'number' && v > 50000 && v < 10000000);
    const hasYear = vals.some(v =>
      (typeof v === 'number' && v >= 2000 && v <= 2030) ||
      (typeof v === 'string' && /^20\d{2}$/.test(v as string))
    );
    return hasPrice && hasYear;
  }

  // Direct fields
  if (hasVehicleFields(obj)) return true;
  // JSON:API: {id, type, attributes: {price, year, ...}}
  const attrs = obj['attributes'];
  if (attrs && typeof attrs === 'object' && !Array.isArray(attrs)) {
    if (hasVehicleFields(attrs as Record<string, unknown>)) return true;
  }
  return false;
}

function findVehicleArray(obj: unknown, depth = 0): Array<Record<string, unknown>> | null {
  if (depth > 14 || !obj || typeof obj !== 'object') return null;
  if (Array.isArray(obj)) {
    return isVehicleArray(obj) ? (obj as Array<Record<string, unknown>>) : null;
  }
  const rec = obj as Record<string, unknown>;
  for (const k of ['vehicles', 'listings', 'results', 'items', 'cars', 'ads', 'adverts', 'data', 'records', 'stock', 'hits', 'content']) {
    if (rec[k]) {
      const found = findVehicleArray(rec[k], depth + 1);
      if (found) return found;
    }
  }
  for (const v of Object.values(rec)) {
    const found = findVehicleArray(v, depth + 1);
    if (found) return found;
  }
  return null;
}

export async function scrapeCarsCoza(
  browser: Browser,
  query: string,
  filters: { maxPrice?: number; maxMileage?: number; minYear?: number },
): Promise<Listing[]> {
  const page = await browser.newPage();

  // Capture ALL JSON responses from cars.co.za with URL logging
  const capturedJsons: Array<{ url: string; json: unknown }> = [];
  const onResponse = async (response: Response) => {
    try {
      const ct = response.headers()['content-type'] ?? '';
      if (!ct.includes('json') || !response.url().includes('cars.co.za')) return;
      const json = await response.json().catch(() => null);
      if (!json) return;
      console.log('[CarsCoza] XHR JSON from:', response.url().replace(/^https?:\/\/[^/]+/, '').substring(0, 80));
      capturedJsons.push({ url: response.url(), json });
    } catch { /* ignore */ }
  };
  page.on('response', onResponse);

  // Capture auth headers from Cars.co.za's own /fw/public/v3/vehicle requests.
  // The page fires facets calls on load; those carry whatever auth the API needs.
  let capturedAuthHeaders: Record<string, string> = {};
  page.on('request', req => {
    if (req.url().includes('/fw/public/v3/vehicle') && Object.keys(capturedAuthHeaders).length === 0) {
      const h = req.headers();
      // Keep only custom / auth headers — skip browser-managed forbidden headers
      capturedAuthHeaders = Object.fromEntries(
        Object.entries(h).filter(([k]) =>
          k.startsWith('x-') || k === 'authorization' || k === 'accept'
        )
      );
      console.log('[CarsCoza] Captured auth header keys:', Object.keys(capturedAuthHeaders).join(',') || 'none');
    }
  });

  try {
    await page.setViewportSize({ width: 1280, height: 900 });
    await page.setExtraHTTPHeaders(BROWSER_HEADERS);

    const { make, model } = normalizeMake(query);

    const mmv = model ? `${make}[${model}]` : make;
    const searchUrl = `${SEARCH_URL}?make_model_variant=${encodeURIComponent(mmv)}&sort=sort_rank&price_type=listing_price&P=1`;
    console.log('[CarsCoza] Navigating to:', searchUrl);

    const response = await page.goto(searchUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
    const status = response?.status() ?? 0;
    console.log('[CarsCoza] HTTP status:', status);

    if (status === 503 || status === 403) {
      console.log('[CarsCoza] Blocked — skipping');
      return [];
    }

    // Wait for Cars.co.za JS to fire its facets API calls (gives us the auth headers)
    await page.waitForTimeout(3000);
    console.log('[CarsCoza] Auth headers captured:', Object.keys(capturedAuthHeaders).join(',') || 'none');

    // Strategy A: use the captured auth headers to call the /fw/public/v3/vehicle API directly.
    // This is the only reliable multi-page approach — SSR only renders page 1.
    const apiItems = await page.evaluate(async ({ mmv, extraHeaders }: { mmv: string; extraHeaders: Record<string, string> }) => {
      const PAGE_SIZE = 20;
      const MAX_ITEMS = 300;
      const base = `/fw/public/v3/vehicle?make_model_variant=${encodeURIComponent(mmv)}&sort=sort_rank&price_type=listing_price&page[limit]=${PAGE_SIZE}`;
      const headers: Record<string, string> = { 'Accept': 'application/vnd.api+json, application/json', ...extraHeaders };

      const fetchPage = (offset: number) =>
        fetch(`${base}&page[offset]=${offset}`, { credentials: 'include', headers })
          .then(r => { console.log('[fw API] offset=' + offset + ' status=' + r.status); return r.ok ? r.json() as Promise<Record<string, unknown>> : null; })
          .catch(e => { console.log('[fw API] fetch error:', String(e)); return null; });

      try {
        const first = await fetchPage(0);
        if (!first) { console.log('[fw API] first page null'); return null; }
        const firstData = (first['data'] as unknown[]) ?? [];
        if (!firstData.length) { console.log('[fw API] first page empty data'); return null; }

        const allItems: unknown[] = [...firstData];
        let offset = PAGE_SIZE;
        let emptyBatches = 0;

        while (allItems.length < MAX_ITEMS && emptyBatches < 2) {
          const batchOffsets: number[] = [];
          for (let j = 0; j < 10; j++) {
            if (allItems.length + (batchOffsets.length + 1) * PAGE_SIZE > MAX_ITEMS) break;
            batchOffsets.push(offset + j * PAGE_SIZE);
          }
          if (!batchOffsets.length) break;

          const results = await Promise.all(batchOffsets.map(async o => {
            const p = await fetchPage(o);
            return p ? ((p['data'] as unknown[]) ?? []) : [];
          }));

          let gotAny = false;
          for (const chunk of results) { if (chunk.length) { allItems.push(...chunk); gotAny = true; } }
          if (!gotAny) emptyBatches++; else emptyBatches = 0;
          offset += batchOffsets.length * PAGE_SIZE;
        }

        console.log('[fw API] fetched total:', allItems.length);
        return allItems;
      } catch (e) { console.log('[fw API] error:', String(e)); return null; }
    }, { mmv, extraHeaders: capturedAuthHeaders });

    if (apiItems && Array.isArray(apiItems) && apiItems.length > 0) {
      const listings = extractFromVehicleArray(apiItems as Array<Record<string, unknown>>, 'api');
      if (listings.length > 0) {
        console.log('[CarsCoza] Listings from API:', listings.length);
        return listings;
      }
    }

    // Strategy B: SSR only gives page 1 (20 items) — use as fallback when API is blocked
    console.log('[CarsCoza] API returned nothing, falling back to SSR page 1');
    const ssrText = await page.evaluate(() => {
      const el = document.getElementById('__NEXT_DATA__');
      return el?.textContent ?? null;
    }).catch(() => null);

    if (ssrText) {
      try {
        const arr = findVehicleArray(JSON.parse(ssrText));
        if (arr?.length) {
          const listings = extractFromVehicleArray(arr, 'ssr-p1');
          if (listings.length > 0) {
            console.log('[CarsCoza] Listings from SSR p1:', listings.length);
            return listings;
          }
        }
      } catch { /* fall through */ }
    }

    // Last resort: parse listing links from the final page in the browser tab
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
