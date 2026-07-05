import type { ReactNode } from "react";

export function KpiCard({ label, value, hint, icon }: { label: string; value: string | number; hint?: string; icon?: ReactNode }) {
  return (
    <div className="kpi-card">
      <div>
        <div className="kpi-label">{label}</div>
        <div className="kpi-value">{value}</div>
        {hint ? <div className="kpi-hint">{hint}</div> : null}
      </div>
      {icon ? <div className="kpi-icon" aria-hidden="true">{icon}</div> : null}
    </div>
  );
}
