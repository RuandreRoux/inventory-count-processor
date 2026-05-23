import type { Listing, Source, Condition, Transmission, Fuel } from '@/lib/types';

export function parsePrice(raw: string): number {
  const digits = raw.replace(/[^0-9]/g, '');
  return digits ? parseInt(digits, 10) : 0;
}

export function parseMileage(raw: string): number {
  const digits = raw.replace(/[^0-9]/g, '');
  return digits ? parseInt(digits, 10) : 0;
}

export function parseYear(raw: string): number {
  const m = raw.match(/\b(19|20)\d{2}\b/);
  return m ? parseInt(m[0], 10) : new Date().getFullYear();
}

export function parseTitle(title: string): { year: number; make: string; model: string; variant: string } {
  const year = parseYear(title);
  const withoutYear = title.replace(/\b(19|20)\d{2}\b/, '').trim();
  const words = withoutYear.split(/\s+/);
  const make = words[0] ?? 'Unknown';
  const model = words[1] ?? '';
  const variant = words.slice(2).join(' ');
  return { year, make, model, variant };
}

export function guessTransmission(text: string): Transmission {
  const t = text.toLowerCase();
  if (t.includes('auto') || t.includes('cvt') || t.includes('dct') || t.includes('dsg') || t.includes('tiptronic')) return 'automatic';
  return 'manual';
}

export function guessFuel(text: string): Fuel {
  const t = text.toLowerCase();
  if (t.includes('electric') || t.includes(' ev ') || /\bev\b/.test(t)) return 'electric';
  if (t.includes('hybrid') || t.includes('phev')) return 'hybrid';
  // SA diesel markers: "diesel", TDI/TDCi/CDi/BlueTEC, GD-6, and the common "X.XD" engine suffix
  if (
    t.includes('diesel') || t.includes('tdi') || t.includes('tdci') ||
    t.includes('cdi') || t.includes('bluetec') ||
    t.includes('gd-6') || t.includes('gd 6') || t.includes('gd6') || // Toyota GD-6 diesel
    t.includes('crdi') || t.includes('dci') || t.includes('d4d') ||
    /\b\d+(\.\d+)?d\b/.test(t)  // e.g. 3.0d, 2.8d, 2.4d, 2.0d
  ) return 'diesel';
  return 'petrol';
}

export function buildId(source: Source, uniqueStr: string): string {
  let hash = 0;
  for (let i = 0; i < uniqueStr.length; i++) hash = (hash * 31 + uniqueStr.charCodeAt(i)) & 0xffffffff;
  return `${source}-${Math.abs(hash).toString(36)}`;
}

interface RawListing {
  title: string;
  priceText: string;
  mileageText: string;
  locationText: string;
  url: string;
  description?: string;
  condition?: string;
  serviceHistory?: boolean;
}

export function normalizeRaw(raw: RawListing, source: Source): Listing | null {
  const price = parsePrice(raw.priceText);
  const mileage = parseMileage(raw.mileageText);
  if (!price || !raw.title) return null;
  const { year, make, model, variant } = parseTitle(raw.title);
  const combined = `${raw.title} ${raw.description ?? ''} ${variant}`;
  const locationParts = raw.locationText.split(/[,|·\-]/).map((s) => s.trim());
  const city = locationParts[0] ?? '';
  const province = locationParts[1] ?? inferProvince(city);
  const conditionMap: Record<string, Condition> = { excellent: 'excellent', 'very good': 'good', good: 'good', fair: 'fair', poor: 'poor' };
  const condition: Condition = raw.condition ? (conditionMap[raw.condition.toLowerCase()] ?? 'good') : 'good';
  return {
    id: buildId(source, raw.url || raw.title),
    source, make, model, variant, year, price, mileage, condition,
    serviceHistory: raw.serviceHistory ?? false,
    transmission: guessTransmission(combined),
    fuel: guessFuel(combined),
    color: 'Unknown', province, city,
    url: raw.url || undefined,
    listedDate: new Date().toISOString().slice(0, 10),
    description: raw.description ?? raw.title,
  };
}

function inferProvince(city: string): string {
  const c = city.toLowerCase();
  if (['johannesburg','pretoria','sandton','midrand','centurion','soweto','randburg'].some((x) => c.includes(x))) return 'Gauteng';
  if (['cape town','stellenbosch','paarl','george','bellville','somerset west'].some((x) => c.includes(x))) return 'Western Cape';
  if (['durban','pietermaritzburg','umhlanga'].some((x) => c.includes(x))) return 'KwaZulu-Natal';
  if (['port elizabeth','east london','gqeberha'].some((x) => c.includes(x))) return 'Eastern Cape';
  if (['polokwane','tzaneen','mokopane'].some((x) => c.includes(x))) return 'Limpopo';
  if (['nelspruit','mbombela'].some((x) => c.includes(x))) return 'Mpumalanga';
  if (['rustenburg','klerksdorp'].some((x) => c.includes(x))) return 'North West';
  if (['bloemfontein'].some((x) => c.includes(x))) return 'Free State';
  return 'Gauteng';
}
