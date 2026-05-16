import { chromium } from 'playwright-core';
import type { Browser } from 'playwright-core';

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

export const BROWSER_HEADERS = {
  'User-Agent':
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
  'Accept-Language': 'en-ZA,en;q=0.9',
};
