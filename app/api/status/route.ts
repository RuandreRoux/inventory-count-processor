import { NextResponse } from 'next/server';
import { FIRECRAWL_ENABLED } from '@/lib/scrapers/index';

export const runtime = 'nodejs';

export async function GET() {
  return NextResponse.json({
    firecrawl_enabled: FIRECRAWL_ENABLED,
    firecrawl_key_set: Boolean(process.env.FIRECRAWL_API_KEY),
  });
}
