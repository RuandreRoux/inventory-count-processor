interface ScoreBadgeProps {
  score: number;
  size?: "sm" | "md" | "lg";
}

function getColor(score: number) {
  if (score >= 80) return { ring: "#f59e0b", text: "#f59e0b", bg: "rgba(245,158,11,0.12)" };
  if (score >= 60) return { ring: "#84cc16", text: "#84cc16", bg: "rgba(132,204,22,0.12)" };
  if (score >= 40) return { ring: "#60a5fa", text: "#60a5fa", bg: "rgba(96,165,250,0.12)" };
  return { ring: "#71717a", text: "#a1a1aa", bg: "rgba(113,113,122,0.1)" };
}

export default function ScoreBadge({ score, size = "md" }: ScoreBadgeProps) {
  const { ring, text, bg } = getColor(score);

  const dim = size === "lg" ? 64 : size === "md" ? 52 : 40;
  const stroke = size === "lg" ? 5 : 4;
  const r = (dim - stroke * 2) / 2;
  const circ = 2 * Math.PI * r;
  const dash = (score / 100) * circ;
  const fontSize = size === "lg" ? 16 : size === "md" ? 13 : 10;

  return (
    <div
      className="relative inline-flex items-center justify-center shrink-0"
      style={{ width: dim, height: dim }}
    >
      <svg width={dim} height={dim} style={{ transform: "rotate(-90deg)" }}>
        <circle
          cx={dim / 2} cy={dim / 2} r={r}
          fill="none" stroke="#27272a" strokeWidth={stroke}
        />
        <circle
          cx={dim / 2} cy={dim / 2} r={r}
          fill="none" stroke={ring} strokeWidth={stroke}
          strokeDasharray={`${dash} ${circ}`}
          strokeLinecap="round"
          style={{ transition: "stroke-dasharray 0.6s ease" }}
        />
      </svg>
      <span
        className="absolute font-black tabular-nums"
        style={{ fontSize, color: text, background: bg, borderRadius: "50%", padding: "2px" }}
      >
        {score}
      </span>
    </div>
  );
}
