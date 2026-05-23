"use client";

import type { RankingWeights } from "@/lib/types";

interface WeightSlidersProps {
  weights: RankingWeights;
  onChange: (w: RankingWeights) => void;
}

const FACTORS: { key: keyof RankingWeights; label: string; description: string }[] = [
  { key: "price", label: "Price", description: "Lower price scores higher" },
  { key: "mileage", label: "Mileage", description: "Fewer kilometres scores higher" },
  { key: "year", label: "Year", description: "Newer vehicles score higher" },
  { key: "condition", label: "Condition", description: "Excellent condition scores highest" },
  { key: "serviceHistory", label: "Service History", description: "Full history scores higher" },
];

export default function WeightSliders({ weights, onChange }: WeightSlidersProps) {
  const total = Object.values(weights).reduce((a, b) => a + b, 0);

  const handleChange = (key: keyof RankingWeights, newVal: number) => {
    onChange({ ...weights, [key]: newVal });
  };

  return (
    <div className="space-y-4">
      <div className="flex items-center justify-between">
        <h2 className="text-xs font-semibold uppercase tracking-widest text-zinc-500">
          Ranking Weights
        </h2>
        <span className="text-[10px] text-zinc-600">Total: {total}</span>
      </div>

      <div className="space-y-4">
        {FACTORS.map(({ key, label, description }) => {
          const w = weights[key];
          const pct = total > 0 ? Math.round((w / total) * 100) : 0;
          return (
            <div key={key} className="space-y-1.5">
              <div className="flex items-center justify-between text-xs">
                <div>
                  <span className="text-zinc-300 font-medium">{label}</span>
                  <span className="text-zinc-600 ml-1.5 text-[10px]">{description}</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <span className="text-zinc-500 text-[10px]">{pct}%</span>
                  <span className="font-semibold text-amber-400 w-5 text-right">{w}</span>
                </div>
              </div>
              <input
                type="range"
                min={0}
                max={50}
                step={1}
                value={w}
                onChange={(e) => handleChange(key, Number(e.target.value))}
                className="w-full h-1.5 rounded-full bg-zinc-700 appearance-none cursor-pointer"
              />
            </div>
          );
        })}
      </div>

      <button
        onClick={() =>
          onChange({ price: 35, mileage: 25, year: 20, condition: 12, serviceHistory: 8 })
        }
        className="w-full rounded-lg border border-zinc-700 py-1.5 text-xs text-zinc-400 hover:border-amber-500/50 hover:text-amber-400 transition-colors"
      >
        Reset to defaults
      </button>
    </div>
  );
}
