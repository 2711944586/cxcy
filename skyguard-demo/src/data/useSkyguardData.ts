import { useEffect, useState } from "react";
import { appPath } from "../utils/navigation";
import type { DataAsset, DroneTarget, FlightPlan, FlightTrackPoint, Geofence, ReportSummary, RiskEvent, RiskModelSample, Sensor, SourceManifest, VisualAsset } from "./types";

type DataState = {
  targets: DroneTarget[];
  plans: FlightPlan[];
  tracks: FlightTrackPoint[];
  events: RiskEvent[];
  geofences: Geofence[];
  sensors: Sensor[];
  summary: ReportSummary | null;
  assets: DataAsset[];
  visualAssets: VisualAsset[];
  riskSamples: RiskModelSample[];
  sourceManifest: SourceManifest | null;
  loading: boolean;
  error: string | null;
};

async function loadJson<T>(path: string): Promise<T> {
  const res = await fetch(appPath(path));
  if (!res.ok) throw new Error(`Failed to load ${path}`);
  return res.json();
}

async function loadOptionalJson<T>(path: string): Promise<T | null> {
  try {
    return await loadJson<T>(path);
  } catch {
    return null;
  }
}

export function useSkyguardData(includeHeavy = false): DataState {
  const [state, setState] = useState<DataState>({
    targets: [],
    plans: [],
    tracks: [],
    events: [],
    geofences: [],
    sensors: [],
    summary: null,
    assets: [],
    visualAssets: [],
    riskSamples: [],
    sourceManifest: null,
    loading: true,
    error: null
  });

  useEffect(() => {
    let alive = true;
    const heavyTracks = includeHeavy ? loadJson<FlightTrackPoint[]>("/mock/flight_tracks.json") : Promise.resolve<FlightTrackPoint[]>([]);
    const heavyRiskSamples = includeHeavy ? loadJson<RiskModelSample[]>("/mock/risk_model_samples.json") : Promise.resolve<RiskModelSample[]>([]);

    Promise.all([
      loadJson<DroneTarget[]>("/mock/drone_targets.json"),
      loadJson<FlightPlan[]>("/mock/flight_plans.json"),
      heavyTracks,
      loadJson<RiskEvent[]>("/mock/risk_events.json"),
      loadJson<Geofence[]>("/mock/geofences.json"),
      loadJson<Sensor[]>("/mock/sensor_status.json"),
      loadJson<ReportSummary>("/mock/report_summary.json"),
      loadJson<DataAsset[]>("/mock/data_assets.json"),
      loadJson<VisualAsset[]>("/assets/asset_manifest.json"),
      heavyRiskSamples,
      loadOptionalJson<SourceManifest>("/source-data/manifest.json")
    ]).then(([targets, plans, tracks, events, geofences, sensors, summary, assets, visualAssets, riskSamples, sourceManifest]) => {
      if (!alive) return;
      setState({ targets, plans, tracks, events, geofences, sensors, summary, assets, visualAssets, riskSamples, sourceManifest, loading: false, error: null });
    }).catch((error: Error) => {
      if (!alive) return;
      setState((prev) => ({ ...prev, loading: false, error: error.message }));
    });
    return () => { alive = false; };
  }, [includeHeavy]);

  return state;
}
