# DreamCar ZA — Claude Code Guide

## What this is
A Next.js car-listing aggregator for South Africa. Users search by make/model, listings are scraped live from Cars.co.za via Firecrawl, ranked by a configurable scoring algorithm, and displayed in a filterable grid.

## Tech stack
- **Next.js 16** (App Router, React 19, Turbopack in dev)
- **Firecrawl** — cloud scraping API, no browser/Playwright required
- **Tailwind CSS 4**
- **TypeScript 5**
- No database — results cached in-memory per query for 30 minutes

## Architecture

```
User search
  → GET /api/search?q=Toyota+Fortuner&...
      → lib/scrapers/index.ts              (cache check + orchestration)
          → lib/scrapers/carscoza-firecrawl.ts  (Firecrawl → Cars.co.za)
              → 4 sort orders × 3 pages = 12 parallel Firecrawl requests
              → AI JSON extraction per page (~10-20 listings each)
      → lib/ranking.ts                     (weighted score 0-100)
  → JSON[]  →  app/search/page.tsx         (ListingGrid + filters + sliders)
```

## Key files

| File | Purpose |
|---|---|
| `lib/scrapers/carscoza-firecrawl.ts` | Firecrawl scraper — multi-sort parallel fetching |
| `lib/scrapers/index.ts` | Cache + orchestration, exports `FIRECRAWL_ENABLED` |
| `lib/scrapers/cache.ts` | In-memory cache, 30-min TTL keyed by query |
| `lib/scrapers/normalize.ts` | `buildId`, `guessTransmission`, `guessFuel`, `normalizeRaw` |
| `lib/ranking.ts` | Weighted min-max scoring, returns `score` 0-100 |
| `lib/types.ts` | `Listing`, `SearchFilters`, `RankingWeights` types |
| `lib/mock-data.ts` | Fallback listings shown when `FIRECRAWL_API_KEY` is not set |
| `app/api/search/route.ts` | Main API route — Node runtime, 60 s max duration |
| `app/api/status/route.ts` | Health check — returns `{ firecrawl_enabled, firecrawl_key_set }` |
| `app/search/page.tsx` | Search results UI with filters and ranking weight sliders |
| `components/ListingCard.tsx` | Listing card with score badge, source tag, and image |

## Environment variables

| Variable | Required | Description |
|---|---|---|
| `FIRECRAWL_API_KEY` | Yes | Firecrawl API key — enables live scraping |

No other env vars are needed. `SCRAPING_ENABLED` and `SCRAPERAPI_KEY` are gone.

## Local development

```bash
npm install
npm run dev          # http://localhost:3000
```

`.env` must contain `FIRECRAWL_API_KEY=fc-...` — the file is git-ignored so add it manually.

## Deployment (Vercel)

1. Import the GitHub repo in the Vercel dashboard
2. Add environment variable: `FIRECRAWL_API_KEY`
3. Deploy — Next.js is auto-detected, no extra build config needed

Requires **Vercel Pro** for the 60-second `maxDuration` on the API route. On Hobby, reduce `pagesPerSort` in `carscoza-firecrawl.ts` to `1` to stay within the 10-second limit.

## Scraping strategy

Cars.co.za is scraped via Firecrawl's AI JSON extraction. Because the site uses session state for pagination, incrementing `P=` across parallel requests is unreliable — pages 2+ return empty or duplicate results without a live session. Instead the scraper fires **4 sort orders × 3 pages = 12 parallel requests**, each returning a genuinely different slice of the inventory. After deduplication by listing ID this yields 40-80 unique listings per query in roughly the same time as a single request (~8-10 s).

Sort orders: `sort_rank`, `price_asc`, `price_desc`, `mileage`.

AutoTrader and ChangeCars are not scraped — Firecrawl blocks AutoTrader outright and ChangeCars returns AI-hallucinated results rather than real listings.

## Ranking weights (user-adjustable via sliders)

| Factor | Default |
|---|---|
| Price (lower = better) | 35% |
| Mileage (lower = better) | 25% |
| Year (newer = better) | 20% |
| Condition | 12% |
| Service history | 8% |
