import { NextRequest, NextResponse } from "next/server";
import { LISTINGS } from "@/lib/mock-data";
import { rankListings } from "@/lib/ranking";
import type { SearchFilters, RankingWeights, Condition, Transmission, Fuel } from "@/lib/types";
import { DEFAULT_WEIGHTS } from "@/lib/types";

export const runtime = "nodejs";

function parseWeights(raw: string | null): RankingWeights {
  if (!raw) return DEFAULT_WEIGHTS;
  const parts = raw.split(",").map(Number);
  if (parts.length !== 5 || parts.some(isNaN)) return DEFAULT_WEIGHTS;
  return {
    price: parts[0],
    mileage: parts[1],
    year: parts[2],
    condition: parts[3],
    serviceHistory: parts[4],
  };
}

export async function GET(req: NextRequest) {
  const sp = req.nextUrl.searchParams;

  const filters: SearchFilters = {
    query: sp.get("q") ?? "",
    maxPrice: sp.get("maxPrice") ? Number(sp.get("maxPrice")) : undefined,
    maxMileage: sp.get("maxMileage") ? Number(sp.get("maxMileage")) : undefined,
    minYear: sp.get("minYear") ? Number(sp.get("minYear")) : undefined,
    condition: (sp.get("condition") as Condition) || undefined,
    serviceHistoryOnly: sp.get("serviceHistoryOnly") === "true" || undefined,
    transmission: (sp.get("transmission") as Transmission) || undefined,
    fuel: (sp.get("fuel") as Fuel) || undefined,
    province: sp.get("province") || undefined,
  };

  const weights = parseWeights(sp.get("weights"));

  let results = LISTINGS;

  if (filters.query.trim()) {
    const q = filters.query.toLowerCase();
    results = results.filter(
      (l) =>
        l.make.toLowerCase().includes(q) ||
        l.model.toLowerCase().includes(q) ||
        l.variant.toLowerCase().includes(q) ||
        `${l.make} ${l.model}`.toLowerCase().includes(q)
    );
  }

  if (filters.maxPrice !== undefined) {
    results = results.filter((l) => l.price <= filters.maxPrice!);
  }
  if (filters.maxMileage !== undefined) {
    results = results.filter((l) => l.mileage <= filters.maxMileage!);
  }
  if (filters.minYear !== undefined) {
    results = results.filter((l) => l.year >= filters.minYear!);
  }
  if (filters.condition) {
    results = results.filter((l) => l.condition === filters.condition);
  }
  if (filters.serviceHistoryOnly) {
    results = results.filter((l) => l.serviceHistory);
  }
  if (filters.transmission) {
    results = results.filter((l) => l.transmission === filters.transmission);
  }
  if (filters.fuel) {
    results = results.filter((l) => l.fuel === filters.fuel);
  }
  if (filters.province) {
    results = results.filter((l) => l.province === filters.province);
  }

  const ranked = rankListings(results, weights);

  return NextResponse.json(ranked);
}
