import type { Source } from "@/lib/types";
import { SOURCE_COLORS, SOURCE_LABELS } from "@/lib/types";

interface SourceBadgeProps {
  source: Source;
  className?: string;
}

export default function SourceBadge({ source, className = "" }: SourceBadgeProps) {
  return (
    <span
      className={`inline-flex items-center rounded-full border px-2 py-0.5 text-[10px] font-semibold uppercase tracking-wider ${SOURCE_COLORS[source]} ${className}`}
    >
      {SOURCE_LABELS[source]}
    </span>
  );
}
