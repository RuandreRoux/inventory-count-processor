import type { Listing } from '@/lib/types';
import { buildId, guessTransmission, guessFuel } from './normalize';

const FIRECRAWL_BASE = 'https://api.firecrawl.dev/v1';

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

const EXTRACTION_PROMPT =
  'Extract ALL car listings on this page — there should be around 20. For each listing include: ' +
  'title, make, model, variant, year (number), price (ZAR number), mileage (km number), ' +
  'city, province, url (full https://www.cars.co.za listing URL), imageUrl (full image URL), ' +
  'condition, transmission, serviceHistory (boolean).';

// Scrapes a single URL via Firecrawl extract, returns empty on timeout or error.
async function scrapeUrl(url: string, apiKey: string, timeoutMs = 25_000): Promise<ExtractedListing[]> {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const res = await fetch(`${FIRECRAWL_BASE}/scrape`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
      signal: controller.signal,
      body: JSON.stringify({
        url,
        formats: ['extract'],
        extract: { prompt: EXTRACTION_PROMPT, schema: EXTRACTION_SCHEMA },
        waitFor: 2000,
        location: { country: 'ZA' },
      }),
    });
    if (!res.ok) {
      const text = await res.text().catch(() => '');
      console.error(`[CarsCozaFirecrawl] HTTP ${res.status} for ${url}: ${text.slice(0, 200)}`);
      return [];
    }
    const data = (await res.json()) as { data?: { extract?: { listings?: ExtractedListing[] } } };
    return data.data?.extract?.listings ?? [];
  } catch (e: unknown) {
    if (e instanceof Error && e.name === 'AbortError') {
      console.error(`[CarsCozaFirecrawl] Timeout for ${url}`);
    } else {
      console.error(`[CarsCozaFirecrawl] Error for ${url}:`, e);
    }
    return [];
  } finally {
    clearTimeout(timer);
  }
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

  // Fallback: construct CDN image URL from listing ID when Firecrawl doesn't return one
  let imageUrl = item.imageUrl || '';
  if (!imageUrl) {
    const listingId = url.match(/\/(\d{6,})\//)?.[1];
    if (listingId) {
      const slug = `${year}-${make}-${item.model ?? ''}${item.variant ? '-' + item.variant : ''}`
        .replace(/\./g, '').replace(/\s+/g, '-').replace(/[^a-zA-Z0-9-]/g, '');
      imageUrl = `https://img-ik.cars.co.za/ik-seo/carsimages/${listingId}/${slug}.jpg?tr=f-auto,h-267,w-400,q-80`;
    }
  }

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
    imageUrl: imageUrl || undefined,
  };
}

export async function scrapeCarsCozaFirecrawl(query: string, pagesPerSort = 3): Promise<Listing[]> {
  const apiKey = process.env.FIRECRAWL_API_KEY;
  if (!apiKey) {
    console.log('[CarsCozaFirecrawl] FIRECRAWL_API_KEY not set, skipping');
    return [];
  }

  const { make, model } = normalizeMake(query);
  const mmv = model ? `${make}[${model}]` : make;
  const mmvEncoded = encodeURIComponent(mmv).replace(/%5B/gi, '[').replace(/%5D/gi, ']');

  // Two sort orders × pagesPerSort = 6 URLs by default, fired in parallel
  const urls = ['sort_rank', 'price_asc'].flatMap(sort =>
    Array.from({ length: pagesPerSort }, (_, i) =>
      `https://www.cars.co.za/usedcars/?make_model_variant=${mmvEncoded}&sort=${sort}&P=${i + 1}`,
    ),
  );

  const results = await Promise.allSettled(urls.map(url => scrapeUrl(url, apiKey)));
  const extracted = results.flatMap(r => r.status === 'fulfilled' ? r.value : []);

  const allListings: Listing[] = [];
  const seenIds = new Set<string>();
  for (const item of extracted) {
    const listing = mapToListing(item);
    if (listing && !seenIds.has(listing.id)) {
      seenIds.add(listing.id);
      allListings.push(listing);
    }
  }

  console.log(`[CarsCozaFirecrawl] ${allListings.length} unique listings for "${query}"`);
  return allListings;
}
