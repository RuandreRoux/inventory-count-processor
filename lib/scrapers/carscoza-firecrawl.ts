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

type BatchStatusResponse = {
  status?: 'processing' | 'completed' | 'failed' | 'cancelled';
  completed?: number;
  total?: number;
  data?: Array<{ json?: { listings?: ExtractedListing[] }; metadata?: { sourceURL?: string } }>;
};

// Sends all URLs to Firecrawl's batch endpoint, polls until complete.
// Firecrawl manages concurrency on their end — avoids the timeout/rate-limit issues
// that occur when firing many individual scrape requests simultaneously.
async function batchScrape(urls: string[], apiKey: string): Promise<ExtractedListing[]> {
  // Start the batch job
  let startRes: Response;
  try {
    startRes = await fetch(`${FIRECRAWL_BASE}/batch/scrape`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
      body: JSON.stringify({
        urls,
        formats: ['json'],
        jsonOptions: { prompt: EXTRACTION_PROMPT, schema: EXTRACTION_SCHEMA },
        waitFor: 2000,
        location: { country: 'ZA' },
      }),
    });
  } catch (e) {
    console.error('[CarsCozaFirecrawl] Batch start network error:', e);
    return [];
  }

  if (!startRes.ok) {
    const text = await startRes.text().catch(() => '');
    console.error(`[CarsCozaFirecrawl] Batch start HTTP ${startRes.status}: ${text.slice(0, 200)}`);
    return [];
  }

  const { id } = (await startRes.json()) as { id?: string };
  if (!id) {
    console.error('[CarsCozaFirecrawl] No batch ID in response');
    return [];
  }
  console.log(`[CarsCozaFirecrawl] Batch ${id} started — ${urls.length} URLs`);

  // Poll until completed or deadline reached (API route maxDuration is 60s)
  const deadline = Date.now() + 50_000;
  while (Date.now() < deadline) {
    await new Promise(r => setTimeout(r, 3000));

    const pollRes = await fetch(`${FIRECRAWL_BASE}/batch/scrape/${id}`, {
      headers: { Authorization: `Bearer ${apiKey}` },
    }).catch(() => null);

    if (!pollRes?.ok) continue;

    const status = (await pollRes.json()) as BatchStatusResponse;
    console.log(`[CarsCozaFirecrawl] Batch ${id}: ${status.status} (${status.completed ?? 0}/${status.total ?? urls.length})`);

    if (status.status === 'completed') {
      return (status.data ?? []).flatMap(page => page.json?.listings ?? []);
    }
    if (status.status === 'failed' || status.status === 'cancelled') {
      console.error(`[CarsCozaFirecrawl] Batch ${id} ${status.status}`);
      return [];
    }
  }

  console.error(`[CarsCozaFirecrawl] Batch ${id} did not complete within deadline`);
  return [];
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

  // Two sort orders × pagesPerSort = 6 URLs by default, sent as one batch job
  const urls = ['sort_rank', 'price_asc'].flatMap(sort =>
    Array.from({ length: pagesPerSort }, (_, i) =>
      `https://www.cars.co.za/usedcars/?make_model_variant=${mmvEncoded}&sort=${sort}&P=${i + 1}`,
    ),
  );

  const extracted = await batchScrape(urls, apiKey);

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
