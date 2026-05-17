import { NextRequest, NextResponse } from 'next/server';
import { launchBrowser } from '@/lib/scrapers/browser';

export const runtime = 'nodejs';
export const maxDuration = 60;

export async function GET(req: NextRequest) {
  const sp = req.nextUrl.searchParams;
  const site = sp.get('site') ?? 'autotrader';
  const query = sp.get('q') ?? 'Toyota';

  if (process.env.SCRAPING_ENABLED !== 'true') {
    return NextResponse.json({ error: 'SCRAPING_ENABLED is not true' }, { status: 403 });
  }

  let browser;
  try {
    browser = await launchBrowser();
    const page = await browser.newPage();

    let url: string;
    if (site === 'carscoza') {
      const slug = query.toLowerCase().replace(/\s+/g, '-').replace(/[^a-z0-9-]/g, '');
      url = `https://www.cars.co.za/used-cars-for-sale/${slug}/`;
    } else {
      const params = new URLSearchParams({ 'search[car_make]': query });
      url = `https://www.autotrader.co.za/cars-for-sale?${params}`;
    }

    await page.setExtraHTTPHeaders({
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      'Accept-Language': 'en-ZA,en;q=0.9',
    });

    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(2000);

    const info = await page.evaluate(() => {
      const html = document.documentElement.outerHTML;
      // Extract all unique class names and data-* attributes from the page
      const allEls = Array.from(document.querySelectorAll('*'));
      const classNames = new Set<string>();
      const dataAttrs = new Set<string>();
      for (const el of allEls) {
        for (const cls of Array.from(el.classList)) {
          classNames.add(cls);
        }
        for (const attr of Array.from(el.attributes)) {
          if (attr.name.startsWith('data-')) dataAttrs.add(attr.name);
        }
      }
      // Find likely listing containers (elements repeated many times with same class)
      const classCounts: Record<string, number> = {};
      for (const el of allEls) {
        const cls = el.className && typeof el.className === 'string' ? el.className.trim() : '';
        if (cls) classCounts[cls] = (classCounts[cls] ?? 0) + 1;
      }
      const repeated = Object.entries(classCounts)
        .filter(([, count]) => count >= 3 && count <= 50)
        .sort(([, a], [, b]) => b - a)
        .slice(0, 30)
        .map(([cls, count]) => ({ cls: cls.slice(0, 100), count }));

      // Get JSON-LD scripts
      const jsonLd = Array.from(document.querySelectorAll('script[type="application/ld+json"]'))
        .map(s => s.textContent?.slice(0, 500) ?? '');

      // Get first 5000 chars of body HTML to see structure
      const bodyHtml = document.body.innerHTML.slice(0, 5000);

      return {
        title: document.title,
        url: location.href,
        repeated,
        dataAttrs: Array.from(dataAttrs).slice(0, 50),
        jsonLd,
        bodyHtml,
      };
    });

    return NextResponse.json(info, {
      headers: { 'Cache-Control': 'no-store' },
    });
  } catch (err) {
    return NextResponse.json({ error: String(err) }, { status: 500 });
  } finally {
    await browser?.close();
  }
}
