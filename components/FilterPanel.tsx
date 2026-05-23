"use client";

import type { SearchFilters } from "@/lib/types";

interface FilterPanelProps {
  filters: SearchFilters;
  onChange: (f: Partial<SearchFilters>) => void;
}

const PROVINCES = [
  "All provinces",
  "Gauteng",
  "Western Cape",
  "KwaZulu-Natal",
  "Eastern Cape",
  "Free State",
  "Limpopo",
  "Mpumalanga",
  "North West",
  "Northern Cape",
];

function RangeRow({
  label, value, min, max, step, format, onChange,
}: {
  label: string; value: number | undefined; min: number; max: number; step: number;
  format: (v: number) => string; onChange: (v: number | undefined) => void;
}) {
  const val = value ?? max;
  return (
    <div className="space-y-1.5">
      <div className="flex items-center justify-between text-xs">
        <span className="text-zinc-400">{label}</span>
        <span className="font-semibold text-zinc-200">{value ? format(value) : "Any"}</span>
      </div>
      <input
        type="range" min={min} max={max} step={step} value={val}
        onChange={(e) => {
          const v = Number(e.target.value);
          onChange(v >= max ? undefined : v);
        }}
        className="w-full h-1.5 rounded-full bg-zinc-700 appearance-none cursor-pointer"
      />
      <div className="flex justify-between text-[10px] text-zinc-600">
        <span>{format(min)}</span>
        <span>{format(max)}</span>
      </div>
    </div>
  );
}

export default function FilterPanel({ filters, onChange }: FilterPanelProps) {
  return (
    <div className="space-y-5">
      <h2 className="text-xs font-semibold uppercase tracking-widest text-zinc-500">Filters</h2>

      <RangeRow
        label="Max price" value={filters.maxPrice}
        min={50000} max={1500000} step={10000}
        format={(v) => `R${(v / 1000).toFixed(0)}k`}
        onChange={(v) => onChange({ maxPrice: v })}
      />

      <RangeRow
        label="Max mileage" value={filters.maxMileage}
        min={0} max={250000} step={5000}
        format={(v) => `${(v / 1000).toFixed(0)}k km`}
        onChange={(v) => onChange({ maxMileage: v })}
      />

      <RangeRow
        label="Min year" value={filters.minYear}
        min={2010} max={2025} step={1}
        format={(v) => String(v)}
        onChange={(v) => onChange({ minYear: v })}
      />

      <div className="space-y-2">
        <p className="text-xs text-zinc-400">Condition</p>
        {(["excellent", "good", "fair", "poor"] as const).map((c) => (
          <label key={c} className="flex items-center gap-2 cursor-pointer">
            <input
              type="radio"
              name="condition"
              checked={filters.condition === c}
              onChange={() => onChange({ condition: filters.condition === c ? undefined : c })}
              className="accent-amber-500"
            />
            <span className="text-sm capitalize text-zinc-300">{c}</span>
          </label>
        ))}
        {filters.condition && (
          <button
            onClick={() => onChange({ condition: undefined })}
            className="text-[11px] text-zinc-500 hover:text-amber-400 transition-colors"
          >
            Clear condition
          </button>
        )}
      </div>

      <label className="flex items-center gap-2 cursor-pointer">
        <input
          type="checkbox"
          checked={!!filters.serviceHistoryOnly}
          onChange={(e) => onChange({ serviceHistoryOnly: e.target.checked || undefined })}
          className="accent-amber-500 h-3.5 w-3.5"
        />
        <span className="text-sm text-zinc-300">Full service history only</span>
      </label>

      <div className="space-y-1.5">
        <p className="text-xs text-zinc-400">Province</p>
        <select
          value={filters.province ?? ""}
          onChange={(e) => onChange({ province: e.target.value || undefined })}
          className="w-full rounded-lg border border-zinc-700 bg-zinc-900 px-3 py-2 text-sm text-zinc-200 focus:border-amber-500 focus:outline-none"
        >
          {PROVINCES.map((p) => (
            <option key={p} value={p === "All provinces" ? "" : p}>
              {p}
            </option>
          ))}
        </select>
      </div>

      <div className="space-y-1.5">
        <p className="text-xs text-zinc-400">Transmission</p>
        <div className="flex gap-2">
          {(["manual", "automatic"] as const).map((t) => (
            <button
              key={t}
              onClick={() => onChange({ transmission: filters.transmission === t ? undefined : t })}
              className={`flex-1 rounded-lg border py-1.5 text-xs font-medium capitalize transition-colors ${
                filters.transmission === t
                  ? "border-amber-500 bg-amber-500/10 text-amber-400"
                  : "border-zinc-700 text-zinc-400 hover:border-zinc-600"
              }`}
            >
              {t}
            </button>
          ))}
        </div>
      </div>

      <div className="space-y-1.5">
        <p className="text-xs text-zinc-400">Fuel type</p>
        <div className="grid grid-cols-2 gap-2">
          {(["petrol", "diesel", "hybrid", "electric"] as const).map((f) => (
            <button
              key={f}
              onClick={() => onChange({ fuel: filters.fuel === f ? undefined : f })}
              className={`rounded-lg border py-1.5 text-xs font-medium capitalize transition-colors ${
                filters.fuel === f
                  ? "border-amber-500 bg-amber-500/10 text-amber-400"
                  : "border-zinc-700 text-zinc-400 hover:border-zinc-600"
              }`}
            >
              {f}
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}
