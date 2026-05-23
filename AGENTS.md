<!-- BEGIN:nextjs-agent-rules -->
# This is NOT the Next.js you know

This version has breaking changes — APIs, conventions, and file structure may all differ from your training data. Read the relevant guide in `node_modules/next/dist/docs/` before writing any code. Heed deprecation notices.
<!-- END:nextjs-agent-rules -->

# DreamCar ZA — Agent Notes

See CLAUDE.md for the full architecture guide.

Key things to know before changing scraping code:
- Scraping uses **Firecrawl only** — no Playwright, no browser, no `@sparticuz/chromium`.
- The only scraper is `lib/scrapers/carscoza-firecrawl.ts`. Do not add Playwright dependencies.
- `FIRECRAWL_API_KEY` is the only env var required. `SCRAPING_ENABLED` does not exist.
- Pagination via `P=` URL param is unreliable without session state — the multi-sort strategy in `scrapeCarsCozaFirecrawl` is intentional.
- The in-memory cache has a 30-min TTL. Restart the server to clear it during development.
