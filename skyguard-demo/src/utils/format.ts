import type { RiskLevel } from "../data/types";

export const riskText: Record<RiskLevel, string> = {
  low: "低风险",
  medium: "中风险",
  high: "高风险",
  critical: "紧急风险"
};

export function pct(value: number) {
  return `${Math.round(value * 100)}%`;
}

export function compact(value: number) {
  return new Intl.NumberFormat("zh-CN").format(Math.round(value));
}
