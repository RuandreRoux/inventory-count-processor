export type Source = 'autotrader' | 'carscoza' | 'changecars' | 'olx' | 'gumtree' | 'facebook' | 'iol';
export type Condition = 'excellent' | 'good' | 'fair' | 'poor';
export type Transmission = 'manual' | 'automatic';
export type Fuel = 'petrol' | 'diesel' | 'hybrid' | 'electric';

export interface Listing {
  id: string;
  source: Source;
  make: string;
  model: string;
  variant: string;
  year: number;
  price: number;
  mileage: number;
  condition: Condition;
  serviceHistory: boolean;
  transmission: Transmission;
  fuel: Fuel;
  color: string;
  province: string;
  city: string;
  listedDate: string;
  description: string;
  score?: number;
  scoreBreakdown?: ScoreBreakdown;
}

export interface ScoreBreakdown {
  price: number;
  mileage: number;
  year: number;
  condition: number;
  serviceHistory: number;
}

export interface SearchFilters {
  query: string;
  maxPrice?: number;
  maxMileage?: number;
  minYear?: number;
  condition?: Condition;
  serviceHistoryOnly?: boolean;
  transmission?: Transmission;
  fuel?: Fuel;
  province?: string;
  sources?: Source[];
}

export interface RankingWeights {
  price: number;
  mileage: number;
  year: number;
  condition: number;
  serviceHistory: number;
}

export const DEFAULT_WEIGHTS: RankingWeights = {
  price: 35,
  mileage: 25,
  year: 20,
  condition: 12,
  serviceHistory: 8,
};

export const SOURCE_LABELS: Record<Source, string> = {
  autotrader: 'AutoTrader SA',
  carscoza: 'Cars.co.za',
  changecars: 'ChangeCars',
  olx: 'OLX',
  gumtree: 'Gumtree',
  facebook: 'Facebook',
  iol: 'IOL Motoring',
};

export const SOURCE_COLORS: Record<Source, string> = {
  autotrader: 'bg-blue-500/20 text-blue-300 border-blue-500/30',
  carscoza: 'bg-red-500/20 text-red-300 border-red-500/30',
  changecars: 'bg-teal-500/20 text-teal-300 border-teal-500/30',
  olx: 'bg-purple-500/20 text-purple-300 border-purple-500/30',
  gumtree: 'bg-green-500/20 text-green-300 border-green-500/30',
  facebook: 'bg-indigo-500/20 text-indigo-300 border-indigo-500/30',
  iol: 'bg-orange-500/20 text-orange-300 border-orange-500/30',
};

export const SOURCE_BAR_COLORS: Record<Source, string> = {
  autotrader: 'from-blue-600 to-blue-500',
  carscoza: 'from-red-600 to-red-500',
  changecars: 'from-teal-600 to-teal-500',
  olx: 'from-purple-600 to-purple-500',
  gumtree: 'from-green-700 to-green-500',
  facebook: 'from-indigo-700 to-indigo-500',
  iol: 'from-orange-600 to-orange-500',
};
