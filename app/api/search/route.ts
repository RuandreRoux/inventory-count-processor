import { NextRequest, NextResponse } from 'next/server';
import { searchDb } from '@/lib/db-search';
import { SUPABASE_ENABLED } from '@/lib/supabase';
import type { SearchFilters, RankingWeights, Condition, Transmission, Fuel } from '@/lib/types';
import { DEFAULT_WEIGHTS } from '@/lib/types';

export const runtime = 'nodejs';
export const maxDuration = 60;

function parseWeights(raw: string | null): RankingWeights {
  if (!raw) return DEFAULT_WEIGHTS;
  const parts = raw.split(',').map(Number);
  if (parts.length !== 6 || parts.some(isNaN)) return DEFAULT_WEIGHTS;
  return { price: parts[0], mileage: parts[1], year: parts[2], warranty: parts[3], condition: parts[4], serviceHistory: parts[5] };
}

export async function GET(req: NextRequest) {
  const sp = req.nextUrl.searchParams;

  const filters: SearchFilters = {
    query:            sp.get('q') ?? '',
    maxPrice:         sp.get('maxPrice') ? Number(sp.get('maxPrice')) : undefined,
    maxMileage:       sp.get('maxMileage') ? Number(sp.get('maxMileage')) : undefined,
    minYear:          sp.get('minYear') ? Number(sp.get('minYear')) : undefined,
    condition:        (sp.get('condition') as Condition) || undefined,
    serviceHistoryOnly: sp.get('serviceHistoryOnly') === 'true' || undefined,
    transmission:     (sp.get('transmission') as Transmission) || undefined,
    fuel:             (sp.get('fuel') as Fuel) || undefined,
    province:         sp.get('province') || undefined,
  };

  const weights = parseWeights(sp.get('weights'));

  const ranked = await searchDb(filters, weights);

  return NextResponse.json(ranked, {
    headers: {
      'X-Data-Source': SUPABASE_ENABLED ? 'database' : 'disabled',
      'Cache-Control':  'no-store',
    },
  });
}
