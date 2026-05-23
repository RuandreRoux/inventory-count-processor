import type { Listing } from '@/lib/types';

interface CacheEntry {
  listings: Listing[];
  expiresAt: number;
}

const store = new Map<string, CacheEntry>();
const TTL_MS = 30 * 60 * 1000;

export function cacheGet(key: string): Listing[] | null {
  const entry = store.get(key);
  if (!entry) return null;
  if (Date.now() > entry.expiresAt) {
    store.delete(key);
    return null;
  }
  return entry.listings;
}

export function cacheSet(key: string, listings: Listing[]): void {
  store.set(key, { listings, expiresAt: Date.now() + TTL_MS });
}

export function cacheBuildKey(params: Record<string, string | number | boolean | undefined>): string {
  return Object.entries(params)
    .filter(([, v]) => v !== undefined)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([k, v]) => `${k}=${v}`)
    .join('&');
}
