import type { Listing } from '@/lib/types';
import { getSupabase } from '@/lib/supabase';

export interface ScrapeStats {
  found: number;
  upserted: number;
  deactivated: number;
}

function listingToRow(l: Listing) {
  return {
    id:              l.id,
    source:          l.source,
    make:            l.make,
    model:           l.model,
    variant:         l.variant,
    year:            l.year,
    price:           l.price,
    mileage:         l.mileage,
    condition:       l.condition,
    service_history: l.serviceHistory,
    transmission:    l.transmission,
    fuel:            l.fuel,
    color:           l.color,
    province:        l.province,
    city:            l.city,
    listed_date:     l.listedDate,
    description:     l.description,
    url:             l.url ?? null,
    image_url:       l.imageUrl ?? null,
    last_seen_at:    new Date().toISOString(),
    is_active:       true,
  };
}

export async function scrapeToDb(make: string, model: string): Promise<ScrapeStats> {
  const { scrapeCarsCozaFirecrawl } = await import('./carscoza-firecrawl');
  const query = model ? `${make} ${model}` : make;
  const listings = await scrapeCarsCozaFirecrawl(query);

  const rows = listings.map(listingToRow);

  // Batch upsert in chunks of 100
  for (let i = 0; i < rows.length; i += 100) {
    const batch = rows.slice(i, i + 100);
    const { error } = await getSupabase()
      .from('listings')
      .upsert(batch, { onConflict: 'id' });
    if (error) throw new Error(`Supabase upsert failed: ${error.message}`);
  }

  // Soft delete: mark inactive if not seen in the last 48 hours
  const cutoff = new Date(Date.now() - 48 * 60 * 60 * 1000).toISOString();
  const { data: deactivated, error: deactivateError } = await getSupabase()
    .from('listings')
    .update({ is_active: false, updated_at: new Date().toISOString() })
    .ilike('make', make)
    .ilike('model', model)
    .eq('is_active', true)
    .lt('last_seen_at', cutoff)
    .select('id');

  if (deactivateError) {
    console.error('[carscoza-db] Soft delete failed:', deactivateError.message);
  }

  return {
    found:       listings.length,
    upserted:    rows.length,
    deactivated: deactivated?.length ?? 0,
  };
}
