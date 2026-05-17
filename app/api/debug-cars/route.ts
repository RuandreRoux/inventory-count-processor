import { NextRequest, NextResponse } from 'next/server';

export const runtime = 'nodejs';
export const maxDuration = 60;

export async function GET(req: NextRequest) {
  const q = req.nextUrl.searchParams.get('q') ?? 'Toyota Fortuner';

  const { launchBrowser } = await import('@/lib/scrapers/browser');
  const { BROWSER_HEADERS } = await import('@/lib/scrapers/browser');

  const browser = await launchBrowser();
  try {
    const page = await browser.newPage();
    await page.setExtraHTTPHeaders(BROWSER_HEADERS);

    // Determine make/model the same way the scraper does
    const MAKES: Record<string, string> = {
      toyota: 'Toyota', volkswagen: 'Volkswagen', vw: 'Volkswagen', ford: 'Ford',
      bmw: 'BMW', hyundai: 'Hyundai', kia: 'Kia', mazda: 'Mazda', isuzu: 'Isuzu',
      nissan: 'Nissan', honda: 'Honda', suzuki: 'Suzuki', renault: 'Renault',
      audi: 'Audi', haval: 'Haval',
    };
    const qLower = q.toLowerCase();
    const entry = Object.entries(MAKES).find(([k]) => qLower.includes(k));
    const make = entry ? entry[1] : q;
    const model = entry ? qLower.replace(entry[0], '').trim().replace(/\b\w/g, (c: string) => c.toUpperCase()) : '';
    const mmv = model ? `${make}[${model}]` : make;

    const searchUrl = `https://www.cars.co.za/usedcars/?make_model_variant=${encodeURIComponent(mmv)}&sort=sort_rank&price_type=listing_price&P=1`;
    await page.goto(searchUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(2000);

    // Fetch page 0 and page 1 and return the raw responses
    const result = await page.evaluate(async (mmv: string) => {
      const PAGE_SIZE = 20;
      const opts = { credentials: 'include' as RequestCredentials, headers: { 'Accept': 'application/vnd.api+json, application/json' } };

      const tryFetch = async (url: string) => {
        try {
          const r = await fetch(url, opts);
          const status = r.status;
          const body = r.ok ? await r.json() : await r.text().catch(() => 'body read failed');
          return { status, body };
        } catch (e) {
          return { status: 0, error: String(e) };
        }
      };

      // Try both encoded and unencoded bracket forms, and page[number] style
      const urls = [
        // offset style — encoded brackets
        `/fw/public/v3/vehicle?make_model_variant=${encodeURIComponent(mmv)}&sort=sort_rank&price_type=listing_price&page%5Blimit%5D=${PAGE_SIZE}&page%5Boffset%5D=0`,
        // offset style — unencoded brackets
        `/fw/public/v3/vehicle?make_model_variant=${encodeURIComponent(mmv)}&sort=sort_rank&price_type=listing_price&page[limit]=${PAGE_SIZE}&page[offset]=0`,
        // page number style
        `/fw/public/v3/vehicle?make_model_variant=${encodeURIComponent(mmv)}&sort=sort_rank&price_type=listing_price&page%5Bsize%5D=${PAGE_SIZE}&page%5Bnumber%5D=1`,
        // P= style (like the HTML page)
        `/fw/public/v3/vehicle?make_model_variant=${encodeURIComponent(mmv)}&sort=sort_rank&price_type=listing_price&P=1`,
      ];

      const page0Results: Record<string, unknown> = {};
      for (const url of urls) {
        const res = await tryFetch(url);
        const key = url.includes('offset') ? (url.includes('%5B') ? 'offset_encoded' : 'offset_plain')
                  : url.includes('number') ? 'page_number'
                  : 'P_style';
        if (res.status === 200 && typeof res.body === 'object') {
          const body = res.body as Record<string, unknown>;
          page0Results[key] = {
            status: res.status,
            dataLength: Array.isArray(body['data']) ? (body['data'] as unknown[]).length : '?',
            meta: body['meta'],
            links: body['links'],
            topKeys: Object.keys(body),
          };
        } else {
          page0Results[key] = { status: res.status, error: typeof res.body === 'string' ? res.body.substring(0, 200) : 'non-200' };
        }
      }

      // For whichever offset style worked, also try offset=20 (page 2)
      const offsetUrl = `/fw/public/v3/vehicle?make_model_variant=${encodeURIComponent(mmv)}&sort=sort_rank&price_type=listing_price&page%5Blimit%5D=${PAGE_SIZE}&page%5Boffset%5D=20`;
      const page1Res = await tryFetch(offsetUrl);
      let page1Result: Record<string, unknown> = { status: page1Res.status };
      if (page1Res.status === 200 && typeof page1Res.body === 'object') {
        const body = page1Res.body as Record<string, unknown>;
        const data0 = (page0Results['offset_encoded'] as Record<string, unknown> | undefined);
        const firstIds = Array.isArray(body['data']) ? (body['data'] as Array<Record<string, unknown>>).slice(0, 3).map(i => i['id']) : [];
        page1Result = {
          status: 200,
          dataLength: Array.isArray(body['data']) ? (body['data'] as unknown[]).length : '?',
          firstItemIds: firstIds,
          sameAsPage0: JSON.stringify(firstIds) === JSON.stringify(data0),
          meta: body['meta'],
        };
      } else {
        page1Result = { status: page1Res.status, error: typeof page1Res.body === 'string' ? page1Res.body.substring(0, 200) : 'non-200' };
      }

      return { mmv, page0Results, page1Result };
    }, mmv);

    return NextResponse.json(result, { headers: { 'Cache-Control': 'no-store' } });
  } finally {
    await browser.close();
  }
}
