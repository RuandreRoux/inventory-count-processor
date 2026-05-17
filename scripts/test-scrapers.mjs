/**
 * Local scraper test — run with:
 *   node scripts/test-scrapers.mjs [query]
 *
 * Requires Playwright Chromium installed:
 *   npx playwright install chromium
 */

import { chromium } from 'playwright';

const query = process.argv[2] ?? 'Toyota Hilux';
const SEARCH_URL = 'https://www.cars.co.za/usedcars/';
const AT_URL = 'https://www.autotrader.co.za/cars-for-sale';

const HEADERS = {
  'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
  'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
  'Accept-Language': 'en-ZA,en-GB;q=0.9,en;q=0.8',
  'Accept-Encoding': 'gzip, deflate, br',
  'Upgrade-Insecure-Requests': '1',
};

function normalizeMake(q) {
  const MAKES = {
    toyota: 'Toyota', volkswagen: 'Volkswagen', vw: 'Volkswagen', ford: 'Ford',
    bmw: 'BMW', hyundai: 'Hyundai', kia: 'Kia', mazda: 'Mazda', isuzu: 'Isuzu',
    nissan: 'Nissan', honda: 'Honda', suzuki: 'Suzuki', renault: 'Renault',
    audi: 'Audi', haval: 'Haval',
  };
  const lower = q.toLowerCase();
  const entry = Object.entries(MAKES).find(([k]) => lower.includes(k));
  if (!entry) return { make: q, model: '' };
  const model = lower.replace(entry[0], '').trim().replace(/\b\w/g, c => c.toUpperCase());
  return { make: entry[1], model };
}

async function testAutoTrader(browser) {
  console.log('\n=== AutoTrader SA ===');
  const page = await browser.newPage();
  try {
    await page.setViewportSize({ width: 1280, height: 900 });
    await page.setExtraHTTPHeaders(HEADERS);

    const { make, model } = normalizeMake(query);
    const params = new URLSearchParams({ 'search[car_make]': make });
    if (model) params.set('search[car_model]', model);
    const url = `${AT_URL}?${params}`;
    console.log('URL:', url);

    const res = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    console.log('HTTP status:', res?.status());
    await page.waitForLoadState('networkidle').catch(() => {});
    await page.waitForTimeout(2000);

    console.log('Title:', await page.title());
    console.log('Body preview:', (await page.evaluate(() => document.body?.innerText?.slice(0, 200))));

    const adCount = await page.evaluate(() => document.querySelectorAll('a[href*="/ad/"]').length);
    console.log('Links matching /ad/:', adCount);

    const hrefSamples = await page.evaluate(() =>
      [...new Set(Array.from(document.querySelectorAll('a[href]')).map(a => a.href).filter(h => h.includes('autotrader')))].slice(0, 10)
    );
    console.log('Sample hrefs:', hrefSamples);
  } finally {
    await page.close();
  }
}

async function testCarsCoza(browser) {
  console.log('\n=== Cars.co.za ===');
  const page = await browser.newPage();
  const captured = [];

  page.on('response', async (response) => {
    try {
      const ct = response.headers()['content-type'] ?? '';
      if (!ct.includes('json')) return;
      if (!response.url().includes('cars.co.za')) return;
      const json = await response.json().catch(() => null);
      if (!json) return;

      // Try common shapes
      const candidates = [
        json?.listings, json?.results, json?.data, json?.vehicles,
        json?.adverts, json?.items, json?.cars, json?.hits,
        json?.data?.listings, json?.data?.results,
        Array.isArray(json) ? json : null,
      ];
      for (const c of candidates) {
        if (Array.isArray(c) && c.length > 0 && typeof c[0] === 'object') {
          captured.push(...c);
          console.log(`  [API] Captured ${c.length} items from: ${response.url().slice(0, 100)}`);
          break;
        }
      }
    } catch {}
  });

  try {
    await page.setViewportSize({ width: 1280, height: 900 });
    await page.setExtraHTTPHeaders(HEADERS);

    const { make, model } = normalizeMake(query);
    const params = new URLSearchParams({ cat_make: make });
    if (model) params.set('cat_model', model);
    const url = `${SEARCH_URL}?${params}`;
    console.log('URL:', url);

    const res = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
    console.log('HTTP status:', res?.status());
    await page.waitForLoadState('networkidle').catch(() => {});
    await page.waitForTimeout(4000);
    await page.evaluate(() => window.scrollBy(0, 800));
    await page.waitForTimeout(2000);

    console.log('Title:', await page.title());
    console.log('Total API vehicles captured:', captured.length);

    if (captured.length > 0) {
      console.log('\nFirst vehicle sample keys:', Object.keys(captured[0]));
      console.log('First vehicle:', JSON.stringify(captured[0], null, 2).slice(0, 600));
    } else {
      const hrefSamples = await page.evaluate(() =>
        [...new Set(Array.from(document.querySelectorAll('a[href]')).map(a => a.href).filter(h => h.includes('cars.co.za')))].slice(0, 15)
      );
      console.log('No API data captured. Sample hrefs:', hrefSamples);
    }
  } finally {
    page.off('response', () => {});
    await page.close();
  }
}

const browser = await chromium.launch({ headless: true });
try {
  await testAutoTrader(browser);
  await testCarsCoza(browser);
} finally {
  await browser.close();
}
