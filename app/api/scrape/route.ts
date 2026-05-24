import { NextRequest, NextResponse } from 'next/server';
import { getSupabase, SUPABASE_ENABLED } from '@/lib/supabase';
import { scrapeToDb } from '@/lib/scrapers/carscoza-db';

export const runtime = 'nodejs';
export const maxDuration = 60;

export async function POST(req: NextRequest) {
  const authHeader = req.headers.get('authorization') ?? '';
  const secret = process.env.SCRAPE_SECRET;
  if (!secret || authHeader !== `Bearer ${secret}`) {
    return NextResponse.json({ error: 'Unauthorized' }, { status: 401 });
  }

  if (!SUPABASE_ENABLED) {
    return NextResponse.json({ error: 'Supabase not configured' }, { status: 503 });
  }

  let make: string, model: string;
  try {
    const body = await req.json();
    make = String(body.make ?? '').trim();
    model = String(body.model ?? '').trim();
    if (!make) throw new Error('make is required');
  } catch (e) {
    return NextResponse.json({ error: (e as Error).message }, { status: 400 });
  }

  const supabase = getSupabase();

  const { data: job, error: jobError } = await supabase
    .from('scrape_jobs')
    .insert({ make, model, status: 'running' })
    .select('id')
    .single();

  if (jobError) {
    return NextResponse.json({ error: `Failed to create job: ${jobError.message}` }, { status: 500 });
  }

  const jobId: string = job.id;

  try {
    const stats = await scrapeToDb(make, model);

    await supabase
      .from('scrape_jobs')
      .update({
        status:               'completed',
        completed_at:         new Date().toISOString(),
        listings_found:       stats.found,
        listings_upserted:    stats.upserted,
        listings_deactivated: stats.deactivated,
      })
      .eq('id', jobId);

    return NextResponse.json({ ok: true, jobId, stats });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    console.error('[/api/scrape] Error:', message);

    await supabase
      .from('scrape_jobs')
      .update({ status: 'failed', completed_at: new Date().toISOString(), error: message })
      .eq('id', jobId);

    return NextResponse.json({ error: message }, { status: 500 });
  }
}
