import type { Listing, SearchFilters, RankingWeights, Condition, Transmission, Fuel, Source } from './types';
import { rankListings } from './ranking';
import { getSupabase } from './supabase';

// Maps a snake_case DB row back to the camelCase Listing type.
function rowToListing(row: Record<string, unknown>): Listing {
  return {
    id:             String(row.id),
    source:         String(row.source) as Source,
    make:           String(row.make),
    model:          String(row.model),
    variant:        String(row.variant ?? ''),
    year:           Number(row.year),
    price:          Number(row.price),
    mileage:        Number(row.mileage),
    condition:      String(row.condition ?? 'good') as Condition,
    serviceHistory: Boolean(row.service_history),
    transmission:   String(row.transmission ?? 'manual') as Transmission,
    fuel:           String(row.fuel ?? 'petrol') as Fuel,
    color:          String(row.color ?? ''),
    province:       String(row.province ?? ''),
    city:           String(row.city ?? ''),
    listedDate:     String(row.listed_date ?? ''),
    description:    String(row.description ?? ''),
    url:            row.url ? String(row.url) : undefined,
    imageUrl:       row.image_url ? String(row.image_url) : undefined,
  };
}

export async function searchDb(filters: SearchFilters, weights: RankingWeights): Promise<Listing[]> {
  const sb = getSupabase();

  let q = sb.from('listings').select('*').eq('is_active', true);

  // Text search: each token must match at least one of make/model/variant
  const queryText = filters.query.trim();
  if (queryText) {
    const tokens = queryText.toLowerCase().split(/\s+/).filter(Boolean);
    for (const token of tokens) {
      q = q.or(`make.ilike.%${token}%,model.ilike.%${token}%,variant.ilike.%${token}%`);
    }
  }

  if (filters.maxPrice !== undefined)  q = q.lte('price',    filters.maxPrice);
  if (filters.maxMileage !== undefined) q = q.lte('mileage',  filters.maxMileage);
  if (filters.minYear !== undefined)    q = q.gte('year',     filters.minYear);
  if (filters.condition)                q = q.eq('condition', filters.condition);
  if (filters.serviceHistoryOnly)       q = q.eq('service_history', true);
  if (filters.transmission)             q = q.eq('transmission', filters.transmission);
  if (filters.fuel)                     q = q.eq('fuel',      filters.fuel);
  if (filters.province)                 q = q.ilike('province', `%${filters.province}%`);

  const { data, error } = await q.order('price', { ascending: true }).limit(300);

  if (error) {
    console.error('[db-search] Query failed:', error.message);
    return [];
  }

  if (!data || data.length === 0) return [];

  const listings = (data as Record<string, unknown>[]).map(rowToListing);
  return rankListings(listings, weights);
}
