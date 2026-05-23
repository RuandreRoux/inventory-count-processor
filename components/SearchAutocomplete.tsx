"use client";

import { useState, useRef, useEffect, useCallback } from "react";
import { useRouter } from "next/navigation";
import { getSuggestions, type CarSuggestion } from "@/lib/makes-models";

interface Props {
  inputClassName?: string;
  placeholder?: string;
  defaultValue?: string;
  size?: "sm" | "lg";
}

export default function SearchAutocomplete({
  inputClassName,
  placeholder = "Search make, model…",
  defaultValue = "",
  size = "sm",
}: Props) {
  const router = useRouter();
  const [q, setQ] = useState(defaultValue);
  const [suggestions, setSuggestions] = useState<CarSuggestion[]>([]);
  const [activeIndex, setActiveIndex] = useState(-1);
  const [open, setOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleChange = useCallback((value: string) => {
    setQ(value);
    const results = getSuggestions(value);
    setSuggestions(results);
    setActiveIndex(-1);
    setOpen(results.length > 0);
  }, []);

  const navigate = useCallback((query: string) => {
    if (query.trim()) {
      setOpen(false);
      router.push(`/search?q=${encodeURIComponent(query.trim())}`);
    }
  }, [router]);

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (!open) return;
    if (e.key === "ArrowDown") {
      e.preventDefault();
      setActiveIndex(i => Math.min(i + 1, suggestions.length - 1));
    } else if (e.key === "ArrowUp") {
      e.preventDefault();
      setActiveIndex(i => Math.max(i - 1, -1));
    } else if (e.key === "Enter") {
      if (activeIndex >= 0) {
        e.preventDefault();
        navigate(suggestions[activeIndex].label);
      }
    } else if (e.key === "Escape") {
      setOpen(false);
      setActiveIndex(-1);
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    navigate(activeIndex >= 0 ? suggestions[activeIndex].label : q);
  };

  // Close on outside click
  useEffect(() => {
    const handler = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  const isLg = size === "lg";

  const defaultInputClass = isLg
    ? "h-14 w-full max-w-md rounded-xl border border-zinc-700 bg-zinc-900 px-5 text-base text-zinc-100 placeholder-zinc-500 focus:border-amber-500 focus:outline-none focus:ring-1 focus:ring-amber-500/30 transition"
    : "h-9 flex-1 rounded-lg border border-zinc-700 bg-zinc-900 px-3 text-sm text-zinc-100 placeholder-zinc-500 focus:border-amber-500 focus:outline-none transition-colors";

  return (
    <div ref={containerRef} className="relative flex-1">
      <form onSubmit={handleSubmit} className="flex items-center gap-2">
        <input
          ref={inputRef}
          type="text"
          value={q}
          onChange={e => handleChange(e.target.value)}
          onFocus={() => q && setSuggestions(s => s.length ? s : getSuggestions(q)) || setOpen(getSuggestions(q).length > 0)}
          onKeyDown={handleKeyDown}
          placeholder={placeholder}
          autoComplete="off"
          className={inputClassName ?? defaultInputClass}
        />
        <button
          type="submit"
          className={
            isLg
              ? "h-14 rounded-xl bg-gradient-to-r from-amber-500 to-orange-500 px-8 text-base font-bold text-zinc-900 hover:from-amber-400 hover:to-orange-400 transition-all shadow-[0_0_24px_rgba(245,158,11,0.25)] hover:shadow-[0_0_32px_rgba(245,158,11,0.4)] shrink-0"
              : "h-9 rounded-lg bg-amber-500 px-4 text-sm font-semibold text-zinc-900 hover:bg-amber-400 transition-colors"
          }
        >
          {isLg ? "Search Cars" : "Search"}
        </button>
      </form>

      {open && suggestions.length > 0 && (
        <ul className="absolute left-0 top-full z-50 mt-1.5 w-full rounded-xl border border-zinc-700 bg-zinc-900 shadow-xl overflow-hidden">
          {suggestions.map((s, i) => (
            <li key={s.label}>
              <button
                type="button"
                onMouseDown={e => { e.preventDefault(); navigate(s.label); }}
                onMouseEnter={() => setActiveIndex(i)}
                className={`w-full px-4 py-2.5 text-left text-sm flex items-center gap-3 transition-colors ${
                  i === activeIndex ? "bg-zinc-800 text-amber-400" : "text-zinc-200 hover:bg-zinc-800/60"
                }`}
              >
                <svg className="h-3.5 w-3.5 shrink-0 text-zinc-500" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M21 21l-4.35-4.35M17 11A6 6 0 1 1 5 11a6 6 0 0 1 12 0z" />
                </svg>
                <span>
                  <span className="font-medium">{s.make}</span>
                  {s.model && <span className="text-zinc-400"> {s.model}</span>}
                </span>
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
