import type { RiskLevel } from "../../data/types";
import { riskText } from "../../utils/format";

export function RiskBadge({ level }: { level: RiskLevel }) {
  return <span className={`badge ${level}`}>{riskText[level] ?? level}</span>;
}
