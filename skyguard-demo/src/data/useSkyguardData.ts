import { useEffect, useState } from "react";
import type { DataAsset, DroneTarget, FlightPlan, Geofence, ReportSummary, RiskEvent, Sensor } from "./types";

type DataState = {
  targets: DroneTarget[];
  plans: FlightPlan[];
  events: RiskEvent[];
  geofences: Geofence[];
  sensors: Sensor[];
  summary: ReportSummary | null;
  assets: DataAsset[];
  loading: boolean;
};

async function loadJson<T>(path: string): Promise<T> {
  const res = await fetch(path);
  if (!res.ok) throw new Error(`Failed to load ${path}`);
  return res.json();
}

export function useSkyguardData(): DataState {
  const [state, setState] = useState<DataState>({
    targets: [],
    plans: [],
    events: [],
    geofences: [],
    sensors: [],
    summary: null,
    assets: [],
    loading: true
  });

  useEffect(() => {
    let alive = true;
    Promise.all([
      loadJson<DroneTarget[]>("/mock/drone_targets.json"),
      loadJson<FlightPlan[]>("/mock/flight_plans.json"),
      loadJson<RiskEvent[]>("/mock/risk_events.json"),
      loadJson<Geofence[]>("/mock/geofences.json"),
      loadJson<Sensor[]>("/mock/sensor_status.json"),
      loadJson<ReportSummary>("/mock/report_summary.json"),
      loadJson<DataAsset[]>("/mock/data_assets.json")
    ]).then(([targets, plans, events, geofences, sensors, summary, assets]) => {
      if (!alive) return;
      setState({ targets, plans, events, geofences, sensors, summary, assets, loading: false });
    }).catch(() => {
      if (!alive) return;
      setState((prev) => ({ ...prev, loading: false }));
    });
    return () => { alive = false; };
  }, []);

  return state;
}
