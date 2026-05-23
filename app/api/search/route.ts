import { NextRequest, NextResponse } from 'next/server';
import { fetchListings, SCRAPING_ENABLED, FIRECRAWL_ENABLED } from '@/lib/scrapers/index';
import type { SearchFilters, RankingWeights, Condition, Transmission, Fuel } from '@/lib/types';
import { DEFAULT_WEIGHTS } from '@/lib/types';

export const runtime = 'nodejs';
export const maxDuration = 60;

function parseWeights(raw: string | null): RankingWeights {
  if (!raw) return DEFAULT_WEIGHTS;
  const parts = raw.split(',').map(Number);
  if (parts.length !== 5 || parts.some(isNaN)) return DEFAULT_WEIGHTS;
  return { price: parts[0], mileage: parts[1], year: parts[2], condition: parts[3], serviceHistory: parts[4] };
}

export async function GET(req: NextRequest) {
  console.log('[DreamCar] SCRAPING_ENABLED env:', process.env.SCRAPING_ENABLED);
  console.log('[DreamCar] SCRAPING_ENABLED flag:', SCRAPING_ENABLED);

  const sp = req.nextUrl.searchParams;

  const filters: SearchFilters = {
    query: sp.get('q') ?? '',
    maxPrice: sp.get('maxPrice') ? Number(sp.get('maxPrice')) : undefined,
    maxMileage: sp.get('maxMileage') ? Number(sp.get('maxMileage')) : undefined,
    minYear: sp.get('minYear') ? Number(sp.get('minYear')) : undefined,
    condition: (sp.get('condition') as Condition) || undefined,
    serviceHistoryOnly: sp.get('serviceHistoryOnly') === 'true' || undefined,
    transmission: (sp.get('transmission') as Transmission) || undefined,
    fuel: (sp.get('fuel') as Fuel) || undefined,
    province: sp.get('province') || undefined,
  };

  const weights = parseWeights(sp.get('weights'));

  const ranked = await fetchListings({ filters, weights });

  return NextResponse.json(ranked, {
    headers: {
      'X-Data-Source': (SCRAPING_ENABLED || FIRECRAWL_ENABLED) ? 'live' : 'disabled',
      'X-Scraping-Env': process.env.SCRAPING_ENABLED ?? 'undefined',
      'Cache-Control': 'no-store',
    },
  });
}
