import { NextResponse } from 'next/server';

export const runtime = 'nodejs';

export async function GET() {
  return NextResponse.json({
    SCRAPING_ENABLED_env: process.env.SCRAPING_ENABLED ?? 'undefined',
    SCRAPING_ENABLED_parsed: process.env.SCRAPING_ENABLED === 'true',
  });
}
