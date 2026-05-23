"use client";

import Link from "next/link";
import SearchAutocomplete from "@/components/SearchAutocomplete";

const POPULAR = [
  "Toyota Hilux", "VW Polo", "Ford Ranger", "BMW 3 Series",
  "Toyota Fortuner", "Mercedes C-Class", "Hyundai Tucson", "Isuzu D-Max",
];

const STATS = [
  { value: "50+", label: "Live listings" },
  { value: "6", label: "Marketplaces" },
  { value: "5", label: "Ranking factors" },
  { value: "100%", label: "Free to use" },
];

const HOW_IT_WORKS = [
  {
    step: "01",
    title: "Search any vehicle",
    body: "Type a make, model or variant. We search across AutoTrader SA, Cars.co.za, OLX, Gumtree, Facebook and IOL simultaneously.",
  },
  {
    step: "02",
    title: "Apply your filters",
    body: "Narrow by price, mileage, year, condition, service history, province and fuel type.",
  },
  {
    step: "03",
    title: "Adjust the ranking",
    body: "Move the weight sliders to match your priorities. Care most about price? Slide it up. Mileage obsessive? Same deal.",
  },
];

export default function HomePage() {
  return (
    <div className="min-h-screen flex flex-col">
      <header className="flex items-center justify-between px-6 py-4 border-b border-zinc-800/60">
        <span className="text-2xl font-black tracking-tight">
          <span className="text-amber-400">Dream</span>
          <span className="text-white">Car</span>
          <span className="ml-2 rounded bg-amber-500/10 px-1.5 py-0.5 text-[10px] font-semibold uppercase tracking-widest text-amber-400 border border-amber-500/20 align-middle">
            ZA
          </span>
        </span>
        <Link
          href="/search?q="
          className="text-sm text-zinc-400 hover:text-amber-400 transition-colors hidden sm:block"
        >
          Browse all →
        </Link>
      </header>

      <section className="relative flex flex-1 flex-col items-center justify-center overflow-hidden px-4 py-24 text-center">
        <div className="pointer-events-none absolute inset-0 overflow-hidden">
          <div className="absolute left-1/2 top-1/3 h-96 w-96 -translate-x-1/2 -translate-y-1/2 rounded-full bg-amber-500/5 blur-3xl" />
          <div className="absolute left-1/4 bottom-1/4 h-64 w-64 rounded-full bg-amber-600/[0.04] blur-3xl" />
          <div className="absolute right-1/4 top-1/4 h-64 w-64 rounded-full bg-orange-500/[0.04] blur-3xl" />
        </div>

        <div className="relative z-10 max-w-3xl space-y-6">
          <div className="inline-flex items-center gap-2 rounded-full border border-amber-500/20 bg-amber-500/5 px-4 py-1.5 text-xs font-medium text-amber-400">
            <span className="h-1.5 w-1.5 rounded-full bg-amber-400 animate-pulse" />
            Searching 6 SA marketplaces simultaneously
          </div>

          <h1 className="text-4xl font-black leading-tight tracking-tight text-white sm:text-5xl lg:text-6xl">
            Find Your{" "}
            <span className="bg-gradient-to-r from-amber-400 to-orange-500 bg-clip-text text-transparent">
              Dream Car
            </span>{" "}
            in South Africa
          </h1>

          <p className="text-lg text-zinc-400 max-w-xl mx-auto">
            We aggregate listings from AutoTrader SA, Cars.co.za, OLX, Gumtree and more — then rank
            them by what matters <em className="text-zinc-300 not-italic">to you</em>.
          </p>

          <div className="flex justify-center">
            <SearchAutocomplete
              size="lg"
              placeholder="Try: Toyota Hilux, VW Polo, BMW 3 Series..."
            />
          </div>

          <div className="flex flex-wrap justify-center gap-2 pt-2">
            {POPULAR.map((s) => (
              <Link
                key={s}
                href={`/search?q=${encodeURIComponent(s)}`}
                className="rounded-full border border-zinc-700 bg-zinc-900/60 px-3 py-1.5 text-xs text-zinc-400 hover:border-amber-500/50 hover:text-amber-400 transition-colors"
              >
                {s}
              </Link>
            ))}
          </div>
        </div>
      </section>

      <section className="border-y border-zinc-800 bg-zinc-900/30">
        <div className="mx-auto grid max-w-4xl grid-cols-2 divide-x divide-zinc-800 sm:grid-cols-4">
          {STATS.map(({ value, label }) => (
            <div key={label} className="px-6 py-8 text-center">
              <div className="text-3xl font-black text-amber-400">{value}</div>
              <div className="mt-1 text-sm text-zinc-500">{label}</div>
            </div>
          ))}
        </div>
      </section>

      <section className="mx-auto max-w-5xl px-4 py-20">
        <h2 className="mb-12 text-center text-2xl font-black text-white">How DreamCar works</h2>
        <div className="grid gap-8 sm:grid-cols-3">
          {HOW_IT_WORKS.map(({ step, title, body }) => (
            <div key={step} className="rounded-xl border border-zinc-800 bg-[#111118] p-6 space-y-3">
              <div className="text-4xl font-black text-zinc-700">{step}</div>
              <h3 className="font-bold text-white">{title}</h3>
              <p className="text-sm text-zinc-400 leading-relaxed">{body}</p>
            </div>
          ))}
        </div>
      </section>

      <section className="border-t border-zinc-800 bg-zinc-900/20 px-4 py-12">
        <p className="text-center text-xs text-zinc-600 mb-6 uppercase tracking-widest font-semibold">
          Aggregating listings from
        </p>
        <div className="flex flex-wrap justify-center gap-4">
          {[
            { name: "AutoTrader SA", color: "text-blue-400" },
            { name: "Cars.co.za", color: "text-red-400" },
            { name: "OLX", color: "text-purple-400" },
            { name: "Gumtree", color: "text-green-400" },
            { name: "Facebook Marketplace", color: "text-indigo-400" },
            { name: "IOL Motoring", color: "text-orange-400" },
          ].map(({ name, color }) => (
            <span
              key={name}
              className={`rounded-full border border-zinc-800 bg-zinc-900 px-4 py-2 text-sm font-medium ${color}`}
            >
              {name}
            </span>
          ))}
        </div>
        <p className="text-center text-[11px] text-zinc-700 mt-6">
          v1 uses representative mock data. Real feed integrations via dealer APIs and commercial
          partnerships coming soon.
        </p>
      </section>

      <footer className="border-t border-zinc-800 px-6 py-6 text-center text-xs text-zinc-600">
        DreamCar ZA — South African Automotive Aggregator · Built with Next.js 16
      </footer>
    </div>
  );
}
