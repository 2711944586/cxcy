export type RiskLevel = "low" | "medium" | "high" | "critical";

export type DroneTarget = {
  drone_id: string;
  target_type: string;
  confidence: number;
  current_status: string;
  risk_score: number;
  risk_level: RiskLevel;
  lat: number;
  lng: number;
  altitude: number;
  speed: number;
  heading: number;
  last_seen_time: string;
};

export type FlightPlan = {
  plan_id: string;
  operator_id: string;
  aircraft_id: string;
  aircraft_type: string;
  city: string;
  scenario: string;
  planned_takeoff_time: string;
  planned_landing_time: string;
  approved_altitude_max: number;
  planned_distance_km: number;
  approval_status: string;
  risk_score: number;
  risk_level: string;
};

export type RiskEvent = {
  event_id: string;
  drone_id: string;
  event_type: string;
  risk_score: number;
  risk_level: RiskLevel;
  event_time: string;
  city: string;
  lat: number;
  lng: number;
  status: string;
  response_time_min: number;
  closed_loop: boolean;
  source: string;
};

export type Geofence = {
  fence_id: string;
  fence_name: string;
  fence_type: string;
  risk_weight: number;
  active: boolean;
  center: { lat: number; lng: number };
  radius_m: number;
};

export type Sensor = {
  sensor_id: string;
  sensor_type: string;
  location: string;
  online_status: string;
  accuracy: number;
  coverage_radius_m: number;
  last_heartbeat: string;
};

export type ReportSummary = {
  kpis: {
    today_flights: number;
    active_targets: number;
    open_events: number;
    sensor_online_rate: number;
    avg_response_min: number;
  };
  risk_level_distribution: Record<string, number>;
  event_type_top: [string, number][];
  city_risk_top: [string, number][];
  daily_trend: { date: string; flights: number; alerts: number; closed: number }[];
};

export type DataAsset = {
  name: string;
  category: string;
  size_kb: number;
  use: string;
};
