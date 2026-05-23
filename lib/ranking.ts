import type { Listing, RankingWeights, ScoreBreakdown } from './types';

// Standard manufacturer warranty periods (years) for SA market.
// Source: AutoTrader SA, official brand websites, Cars.co.za (verified May 2026).
const WARRANTY_YEARS: Record<string, number> = {
  // 7 year
  gwm: 7,
  // 6 year
  nissan: 6,
  // 5 year
  hyundai: 5, kia: 5, mazda: 5, honda: 5, renault: 5, isuzu: 5,
  haval: 5, subaru: 5, opel: 5, mitsubishi: 5, chery: 5,
  // 4 year
  ford: 4,
  // 3 year
  toyota: 3, volkswagen: 3, vw: 3, 'land rover': 3, landrover: 3,
  jeep: 3, suzuki: 3, peugeot: 3, 'alfa romeo': 3, fiat: 3, jaguar: 3,
  // 2 year
  bmw: 2, 'mercedes-benz': 2, mercedes: 2, porsche: 2, mini: 2,
  // 1 year
  audi: 1,
  // others
  volvo: 5, mg: 5, byd: 5, omoda: 5, jaecoo: 5, jetour: 5, gac: 5,
  lexus: 3, seat: 3, dodge: 3, ram: 3, citroën: 3, citroen: 3,
  mahindra: 3, ssangyong: 3, skoda: 3,
};

const CURRENT_YEAR = new Date().getFullYear();

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

// Gradient score: 1.0 when brand new, 0.0 at expiry, 0.0 beyond.
function warrantyScore(make: string, year: number): number {
  const key = make.trim().toLowerCase();
  const years = WARRANTY_YEARS[key] ?? 3;
  const age = CURRENT_YEAR - year;
  return Math.max(0, (years - age) / years);
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
    weights.price + weights.mileage + weights.year + weights.warranty +
    weights.condition + weights.serviceHistory;

  const ranked = listings.map((listing) => {
    const breakdown: ScoreBreakdown = {
      price:          normalize(listing.price, minPrice, maxPrice, true),
      mileage:        normalize(listing.mileage, minMileage, maxMileage, true),
      year:           normalize(listing.year, minYear, maxYear, false),
      warranty:       warrantyScore(listing.make, listing.year),
      condition:      CONDITION_SCORES[listing.condition] ?? 0.5,
      serviceHistory: listing.serviceHistory ? 1.0 : 0.0,
    };

    const weightedSum =
      breakdown.price          * weights.price +
      breakdown.mileage        * weights.mileage +
      breakdown.year           * weights.year +
      breakdown.warranty       * weights.warranty +
      breakdown.condition      * weights.condition +
      breakdown.serviceHistory * weights.serviceHistory;

    const score = totalWeight > 0 ? Math.round((weightedSum / totalWeight) * 100) : 0;

    return { ...listing, score, scoreBreakdown: breakdown };
  });

  return ranked.sort((a, b) => (b.score ?? 0) - (a.score ?? 0));
}
