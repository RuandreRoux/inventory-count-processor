import { chromium } from 'playwright-core';
import type { Browser, BrowserContext } from 'playwright-core';

const CHROMIUM_REMOTE_URL =
  'https://github.com/Sparticuz/chromium/releases/download/v133.0.0/chromium-v133.0.0-pack.tar';

export async function launchBrowser(): Promise<Browser> {
  if (process.env.VERCEL === '1' || process.env.AWS_LAMBDA_FUNCTION_NAME) {
    const chromiumMin = await import('@sparticuz/chromium-min');
    const executablePath = await chromiumMin.default.executablePath(CHROMIUM_REMOTE_URL);
    return chromium.launch({
      args: chromiumMin.default.args,
      executablePath,
      headless: true,
    });
  }
  return chromium.launch({ headless: true });
}

/**
 * Creates a browser context routed through ScraperAPI's residential proxy network.
 * Requires SCRAPERAPI_KEY env var. Returns null if key is not set.
 * Use this for sites that block data-center IPs (e.g. AutoTrader SA).
 */
export async function newProxiedContext(browser: Browser): Promise<BrowserContext | null> {
  const key = process.env.SCRAPERAPI_KEY;
  if (!key) return null;
  console.log('[Browser] Using ScraperAPI residential proxy for this context');
  return browser.newContext({
    proxy: {
      server: 'http://proxy-server.scraperapi.com:8001',
      username: 'scraperapi',
      password: key,
    },
    ignoreHTTPSErrors: true,
  });
}

export const BROWSER_HEADERS = {
  'User-Agent':
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
  'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
  'Accept-Language': 'en-ZA,en-GB;q=0.9,en;q=0.8',
  'Accept-Encoding': 'gzip, deflate, br',
  'Upgrade-Insecure-Requests': '1',
};
