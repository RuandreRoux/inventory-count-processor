import type { Listing } from '@/lib/types';
import { buildId, guessTransmission, guessFuel } from './normalize';

const FIRECRAWL_SCRAPE_URL = 'https://api.firecrawl.dev/v1/scrape';

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

interface ExtractedListing {
  title?: string;
  make?: string;
  model?: string;
  variant?: string;
  year?: number;
  price?: number;
  mileage?: number;
  city?: string;
  province?: string;
  url?: string;
  imageUrl?: string;
  condition?: string;
  transmission?: string;
  serviceHistory?: boolean;
}

type FirecrawlResponse = {
  success?: boolean;
  data?: { json?: { listings?: ExtractedListing[] } };
};

const EXTRACTION_SCHEMA = {
  type: 'object',
  properties: {
    listings: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          title:          { type: 'string' },
          make:           { type: 'string' },
          model:          { type: 'string' },
          variant:        { type: 'string' },
          year:           { type: 'number' },
          price:          { type: 'number' },
          mileage:        { type: 'number' },
          city:           { type: 'string' },
          province:       { type: 'string' },
          url:            { type: 'string' },
          imageUrl:       { type: 'string' },
          condition:      { type: 'string' },
          transmission:   { type: 'string' },
          serviceHistory: { type: 'boolean' },
        },
      },
    },
  },
  required: ['listings'],
};

async function fetchPage(url: string, apiKey: string): Promise<ExtractedListing[]> {
  let res: Response;
  try {
    res = await fetch(FIRECRAWL_SCRAPE_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
      body: JSON.stringify({
        url,
        formats: ['json'],
        jsonOptions: {
          prompt:
            'Extract all car listings visible on this page. For each listing include: ' +
            'title, make, model, variant, year (number), price (ZAR number), mileage (km number), ' +
            'city, province, url (full https://www.cars.co.za listing URL), imageUrl, ' +
            'condition, transmission, serviceHistory (boolean).',
          schema: EXTRACTION_SCHEMA,
        },
        waitFor: 3000,
        location: { country: 'ZA' },
      }),
    });
  } catch (e) {
    console.error('[CarsCozaFirecrawl] Network error:', e);
    return [];
  }

  if (!res.ok) {
    const text = await res.text().catch(() => '');
    console.error(`[CarsCozaFirecrawl] HTTP ${res.status} for ${url}: ${text.slice(0, 200)}`);
    return [];
  }

  const body = (await res.json()) as FirecrawlResponse;
  const listings = body?.data?.json?.listings;
  if (!Array.isArray(listings)) {
    console.log(`[CarsCozaFirecrawl] No listings array in response for page ${url}`);
    return [];
  }
  return listings;
}

function mapToListing(item: ExtractedListing): Listing | null {
  const make = item.make?.trim();
  const year = item.year;
  const price = item.price;
  if (!make || !year || !price) return null;

  const variantText = `${item.variant ?? ''} ${item.title ?? ''}`;
  const transmission =
    item.transmission?.toLowerCase().includes('auto') ? 'automatic' as const
    : item.transmission?.toLowerCase().includes('man') ? 'manual' as const
    : guessTransmission(variantText);
  const fuel = guessFuel(variantText);

  const url = item.url ?? '';
  const id = buildId('carscoza', url || `${year}-${make}-${item.model}-${item.variant}-${price}`);

  return {
    id,
    source: 'carscoza',
    make,
    model: item.model?.trim() ?? '',
    variant: item.variant?.trim() ?? '',
    year,
    price,
    mileage: item.mileage ?? 0,
    condition: 'good',
    serviceHistory: item.serviceHistory ?? false,
    transmission,
    fuel,
    color: '',
    province: item.province ?? '',
    city: item.city ?? '',
    listedDate: new Date().toISOString().slice(0, 10),
    description: item.title ?? `${year} ${make} ${item.model ?? ''}`.trim(),
    url: url || undefined,
    imageUrl: item.imageUrl || undefined,
  };
}

export async function scrapeCarsCozaFirecrawl(query: string, maxPages = 8): Promise<Listing[]> {
  const apiKey = process.env.FIRECRAWL_API_KEY;
  if (!apiKey) {
    console.log('[CarsCozaFirecrawl] FIRECRAWL_API_KEY not set, skipping');
    return [];
  }

  const { make, model } = normalizeMake(query);
  const mmv = model ? `${make}[${model}]` : make;
  // Cars.co.za requires literal brackets, not percent-encoded
  const mmvEncoded = encodeURIComponent(mmv).replace(/%5B/gi, '[').replace(/%5D/gi, ']');
  const buildUrl = (page: number) =>
    `https://www.cars.co.za/usedcars/?make_model_variant=${mmvEncoded}&sort=sort_rank&P=${page}`;

  const allListings: Listing[] = [];
  const seenIds = new Set<string>();

  const addItems = (items: ExtractedListing[]) => {
    for (const item of items) {
      const listing = mapToListing(item);
      if (listing && !seenIds.has(listing.id)) {
        seenIds.add(listing.id);
        allListings.push(listing);
      }
    }
  };

  // Fetch page 1 first — if it's empty the query returned no results
  const firstPage = await fetchPage(buildUrl(1), apiKey);
  console.log(`[CarsCozaFirecrawl] Page 1: ${firstPage.length} items`);
  if (firstPage.length === 0) return [];
  addItems(firstPage);

  // Fetch remaining pages in batches of 3 in parallel
  const BATCH = 3;
  for (let start = 2; start <= maxPages; start += BATCH) {
    const pageNums = Array.from(
      { length: Math.min(BATCH, maxPages - start + 1) },
      (_, i) => start + i,
    );
    const results = await Promise.all(pageNums.map(p => fetchPage(buildUrl(p), apiKey)));

    let anyResults = false;
    for (const items of results) {
      if (items.length > 0) { anyResults = true; addItems(items); }
    }
    console.log(`[CarsCozaFirecrawl] Pages ${pageNums[0]}-${pageNums[pageNums.length - 1]}: ${results.map(r => r.length).join(',')} items`);
    if (!anyResults) break;
  }

  console.log(`[CarsCozaFirecrawl] Total: ${allListings.length} listings for "${query}"`);
  return allListings;
}
