"use client";

import Link from "next/link";
import SearchAutocomplete from "./SearchAutocomplete";

export default function Navbar() {
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

        <SearchAutocomplete size="sm" placeholder="Search make, model…" />
      </div>
    </header>
  );
}
