"use server";

import { scrapeToDb, type ScrapeStats } from "@/lib/scrapers/carscoza-db";
import { getSupabase, SUPABASE_ENABLED } from "@/lib/supabase";

export async function triggerScrape(
  make: string,
  model: string
): Promise<{ ok: boolean; stats?: ScrapeStats; error?: string }> {
  if (!SUPABASE_ENABLED) return { ok: false, error: "Supabase not configured" };
  if (!make.trim()) return { ok: false, error: "Make is required" };
  try {
    const stats = await scrapeToDb(make.trim(), model.trim());
    return { ok: true, stats };
  } catch (e) {
    return { ok: false, error: (e as Error).message };
  }
}

export async function fetchRecentJobs() {
  if (!SUPABASE_ENABLED) return [];
  const { data } = await getSupabase()
    .from("scrape_jobs")
    .select("*")
    .order("started_at", { ascending: false })
    .limit(15);
  return (data ?? []) as {
    id: string;
    make: string;
    model: string;
    started_at: string;
    completed_at: string | null;
    listings_found: number | null;
    listings_upserted: number | null;
    listings_deactivated: number | null;
    status: string;
    error: string | null;
  }[];
}
