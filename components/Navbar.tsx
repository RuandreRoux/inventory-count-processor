"use client";

import Link from "next/link";
import { useRouter } from "next/navigation";
import { useState } from "react";

export default function Navbar() {
  const router = useRouter();
  const [q, setQ] = useState("");

  const handleSearch = (e: React.FormEvent) => {
    e.preventDefault();
    if (q.trim()) router.push(`/search?q=${encodeURIComponent(q.trim())}`);
  };

  return (
    <header className="sticky top-0 z-50 border-b border-zinc-800 bg-[#09090f]/90 backdrop-blur-md">
      <div className="mx-auto flex max-w-7xl items-center gap-4 px-4 py-3">
        <Link href="/" className="flex items-center gap-2 shrink-0">
          <span className="text-xl font-black tracking-tight">
            <span className="text-amber-400">Dream</span>
            <span className="text-white">Car</span>
          </span>
          <span className="rounded bg-amber-500/10 px-1.5 py-0.5 text-[10px] font-semibold uppercase tracking-widest text-amber-400 border border-amber-500/20">
            ZA
          </span>
        </Link>

        <form onSubmit={handleSearch} className="flex flex-1 items-center gap-2">
          <input
            type="text"
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder="Search make, model…"
            className="h-9 flex-1 rounded-lg border border-zinc-700 bg-zinc-900 px-3 text-sm text-zinc-100 placeholder-zinc-500 focus:border-amber-500 focus:outline-none transition-colors"
          />
          <button
            type="submit"
            className="h-9 rounded-lg bg-amber-500 px-4 text-sm font-semibold text-zinc-900 hover:bg-amber-400 transition-colors"
          >
            Search
          </button>
        </form>
      </div>
    </header>
  );
}
