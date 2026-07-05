import type { DroneTarget, Geofence, RiskEvent } from "../../data/types";

function project(lat: number, lng: number) {
  const x = ((lng - 103) / 20) * 100;
  const y = 100 - ((lat - 21) / 21) * 100;
  return { x: Math.max(3, Math.min(97, x)), y: Math.max(5, Math.min(95, y)) };
}

export function LowAltitudeMap({ targets, geofences, events }: { targets: DroneTarget[]; geofences: Geofence[]; events: RiskEvent[] }) {
  const shownTargets = targets.slice(0, 42);
  const shownFences = geofences.slice(0, 10);
  const shownEvents = events.slice(0, 16);
  const criticalCount = shownTargets.filter((target) => target.risk_level === "critical").length + shownEvents.filter((event) => event.risk_level === "critical").length;
  const openEvents = events.filter((event) => !event.closed_loop).length;
  const latestEvents = shownEvents.slice(0, 3);
  return (
    <div className="map" aria-label="低空态势模拟地图">
      <div className="map-scan" aria-hidden="true" />
      <svg viewBox="0 0 100 100" preserveAspectRatio="none">
        <defs>
          <pattern id="grid" width="8" height="8" patternUnits="userSpaceOnUse">
            <path d="M 8 0 L 0 0 0 8" fill="none" stroke="rgba(116,150,162,.16)" strokeWidth=".35" />
          </pattern>
        </defs>
        <rect width="100" height="100" fill="url(#grid)" />
        <path d="M8,62 C22,45 32,54 45,34 S68,32 88,18" fill="none" stroke="rgba(87,166,255,.38)" strokeWidth=".6" />
        <path d="M12,74 C30,76 46,58 64,64 S82,72 94,58" fill="none" stroke="rgba(79,209,197,.34)" strokeWidth=".5" />
        {shownTargets.slice(0, 8).map((t, i) => {
          const p = project(t.lat, t.lng);
          const p2 = { x: Math.min(98, p.x + 6 + i % 4), y: Math.max(4, p.y - 7 + i % 3) };
          return <path key={t.drone_id} d={`M${p.x},${p.y} C${p.x+4},${p.y-4} ${p2.x-4},${p2.y+4} ${p2.x},${p2.y}`} stroke="rgba(79,209,197,.55)" strokeWidth=".35" fill="none" />;
        })}
      </svg>
      {shownFences.map((f) => {
        const p = project(f.center.lat, f.center.lng);
        const size = Math.max(42, Math.min(120, f.radius_m / 18));
        return <div key={f.fence_id} className={`fence ${f.fence_type}`} style={{ left: `${p.x}%`, top: `${p.y}%`, width: size, height: size }} title={f.fence_name} />;
      })}
      {shownTargets.map((t) => {
        const p = project(t.lat, t.lng);
        return <div key={t.drone_id} className={`target-dot ${t.risk_level}`} style={{ left: `${p.x}%`, top: `${p.y}%` }} title={`${t.drone_id} ${t.risk_score}`} />;
      })}
      {shownEvents.map((e) => {
        const p = project(e.lat, e.lng);
        return <div key={e.event_id} className={`target-dot ${e.risk_level}`} style={{ left: `${p.x}%`, top: `${p.y}%`, width: 7, height: 7 }} title={e.event_type} />;
      })}
      <span className="city-label" style={{ left: "14%", top: "60%" }}>成都</span>
      <span className="city-label" style={{ left: "70%", top: "40%" }}>杭州</span>
      <span className="city-label" style={{ left: "82%", top: "66%" }}>广州</span>
      <div className="map-hud">
        <span>运行域</span>
        <strong>滨江会展核心区</strong>
        <small>{shownTargets.length} 个活动目标 · {openEvents} 件开放事件</small>
      </div>
      <div className="map-layer-switch" aria-label="地图图层状态">
        <span><i />轨迹</span>
        <span><i />围栏</span>
        <span><i />事件</span>
      </div>
      <div className="map-legend" aria-label="风险图例">
        <span><i className="legend-low" />低/中</span>
        <span><i className="legend-high" />高</span>
        <span><i className="legend-critical" />严重 {criticalCount}</span>
      </div>
      <div className="map-event-stack" aria-label="实时事件队列">
        {latestEvents.map((event) => (
          <div key={event.event_id}>
            <span>{event.event_type}</span>
            <strong>{event.city} · {event.response_time_min} 分</strong>
          </div>
        ))}
      </div>
    </div>
  );
}
