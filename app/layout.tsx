import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({ variable: "--font-geist-sans", subsets: ["latin"] });
const geistMono = Geist_Mono({ variable: "--font-geist-mono", subsets: ["latin"] });

export const metadata: Metadata = {
  title: "DreamCar — Find Your Perfect Car in South Africa",
  description:
    "Search and compare car listings across AutoTrader SA, Cars.co.za, OLX, Gumtree and more. Ranked by what matters to you.",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en" className={`${geistSans.variable} ${geistMono.variable} h-full`}>
      <body className="min-h-full bg-[#09090f] text-zinc-100 antialiased" suppressHydrationWarning>{children}</body>
    </html>
  );
}
