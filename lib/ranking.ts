import type { Listing, RankingWeights, ScoreBreakdown } from './types';

const CONDITION_SCORES: Record<string, number> = {
  excellent: 1.0,
  good: 0.7,
  fair: 0.4,
  poor: 0.1,
};

function normalize(value: number, min: number, max: number, invert = false): number {
  if (max === min) return 0.5;
  const n = (value - min) / (max - min);
  return invert ? 1 - n : n;
}

export function rankListings(listings: Listing[], weights: RankingWeights): Listing[] {
  if (listings.length === 0) return [];

  const prices = listings.map((l) => l.price);
  const mileages = listings.map((l) => l.mileage);
  const years = listings.map((l) => l.year);

  const minPrice = Math.min(...prices);
  const maxPrice = Math.max(...prices);
  const minMileage = Math.min(...mileages);
  const maxMileage = Math.max(...mileages);
  const minYear = Math.min(...years);
  const maxYear = Math.max(...years);

  const totalWeight =
    weights.price + weights.mileage + weights.year + weights.condition + weights.serviceHistory;

  const ranked = listings.map((listing) => {
    const breakdown: ScoreBreakdown = {
      price: normalize(listing.price, minPrice, maxPrice, true),
      mileage: normalize(listing.mileage, minMileage, maxMileage, true),
      year: normalize(listing.year, minYear, maxYear, false),
      condition: CONDITION_SCORES[listing.condition] ?? 0.5,
      serviceHistory: listing.serviceHistory ? 1.0 : 0.0,
    };

    const weightedSum =
      breakdown.price * weights.price +
      breakdown.mileage * weights.mileage +
      breakdown.year * weights.year +
      breakdown.condition * weights.condition +
      breakdown.serviceHistory * weights.serviceHistory;

    const score = totalWeight > 0 ? Math.round((weightedSum / totalWeight) * 100) : 0;

    return { ...listing, score, scoreBreakdown: breakdown };
  });

  return ranked.sort((a, b) => (b.score ?? 0) - (a.score ?? 0));
}
