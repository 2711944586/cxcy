import {
  Area,
  AreaChart,
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  ComposedChart,
  Line,
  LineChart,
  Pie,
  PieChart,
  PolarAngleAxis,
  PolarGrid,
  Radar,
  RadarChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis
} from "recharts";

const colors = ["#0b6f68", "#24567d", "#b1843c", "#9b3f3a", "#2f6b4f", "#6d716c"];
const grid = "rgba(97, 111, 128, .16)";
const axis = "#667085";
const tooltip = {
  background: "#ffffff",
  border: "1px solid rgba(121, 134, 153, .24)",
  color: "#172033",
  boxShadow: "0 16px 40px rgba(23, 32, 51, .12)",
  borderRadius: 8
};
const xTick = { fontSize: 11, fill: axis };
const yTick = { fontSize: 11, fill: axis };

export function TrendChart({ data }: { data: { date: string; flights: number; alerts: number; closed: number }[] }) {
  return (
    <ResponsiveContainer width="100%" height={260}>
      <LineChart data={data} margin={{ left: 4, right: 12, top: 8, bottom: 4 }}>
        <CartesianGrid stroke={grid} vertical={false} strokeDasharray="3 6" />
        <XAxis dataKey="date" stroke={axis} tick={xTick} minTickGap={22} tickMargin={8} />
        <YAxis stroke={axis} tick={yTick} tickMargin={8} width={38} />
        <Tooltip contentStyle={tooltip} />
        <Line type="monotone" dataKey="flights" stroke="#0b6f68" dot={false} strokeWidth={2.6} />
        <Line type="monotone" dataKey="alerts" stroke="#b1843c" dot={false} strokeWidth={2.4} />
        <Line type="monotone" dataKey="closed" stroke="#24567d" dot={false} strokeWidth={2.4} />
      </LineChart>
    </ResponsiveContainer>
  );
}

export function SimpleBarChart({ data, nameKey = "name", valueKey = "value" }: { data: Record<string, string | number>[]; nameKey?: string; valueKey?: string }) {
  return (
    <ResponsiveContainer width="100%" height={260}>
      <BarChart data={data} margin={{ left: 4, right: 12, top: 8, bottom: 4 }}>
        <CartesianGrid stroke={grid} vertical={false} strokeDasharray="3 6" />
        <XAxis dataKey={nameKey} stroke={axis} tick={xTick} minTickGap={12} tickMargin={8} />
        <YAxis stroke={axis} tick={yTick} tickMargin={8} width={38} />
        <Tooltip contentStyle={tooltip} />
        <Bar dataKey={valueKey} fill="#0b6f68" radius={[5, 5, 0, 0]} maxBarSize={38} />
      </BarChart>
    </ResponsiveContainer>
  );
}

export function HorizontalBarChart({ data, nameKey = "name", valueKey = "value" }: { data: Record<string, string | number>[]; nameKey?: string; valueKey?: string }) {
  return (
    <ResponsiveContainer width="100%" height={260}>
      <BarChart data={data} layout="vertical" margin={{ left: 28, right: 14, top: 8, bottom: 8 }}>
        <CartesianGrid stroke={grid} horizontal={false} strokeDasharray="3 6" />
        <XAxis type="number" stroke={axis} tick={xTick} tickMargin={8} />
        <YAxis type="category" dataKey={nameKey} stroke={axis} tick={yTick} width={88} />
        <Tooltip contentStyle={tooltip} />
        <Bar dataKey={valueKey} fill="#24567d" radius={[0, 5, 5, 0]} maxBarSize={24} />
      </BarChart>
    </ResponsiveContainer>
  );
}

export function AreaTrendChart({ data }: { data: { date: string; flights: number; alerts: number; closed: number }[] }) {
  return (
    <ResponsiveContainer width="100%" height={260}>
      <AreaChart data={data} margin={{ left: 4, right: 12, top: 8, bottom: 4 }}>
        <CartesianGrid stroke={grid} vertical={false} strokeDasharray="3 6" />
        <XAxis dataKey="date" stroke={axis} tick={xTick} minTickGap={22} tickMargin={8} />
        <YAxis stroke={axis} tick={yTick} tickMargin={8} width={38} />
        <Tooltip contentStyle={tooltip} />
        <Area type="monotone" dataKey="flights" stroke="#0b6f68" fill="#0b6f68" fillOpacity={0.16} />
        <Area type="monotone" dataKey="alerts" stroke="#b1843c" fill="#b1843c" fillOpacity={0.14} />
      </AreaChart>
    </ResponsiveContainer>
  );
}

export function MixedOpsChart({ data }: { data: { date: string; flights: number; alerts: number; closed: number }[] }) {
  return (
    <ResponsiveContainer width="100%" height={260}>
      <ComposedChart data={data} margin={{ left: 4, right: 12, top: 8, bottom: 4 }}>
        <CartesianGrid stroke={grid} vertical={false} strokeDasharray="3 6" />
        <XAxis dataKey="date" stroke={axis} tick={xTick} minTickGap={22} tickMargin={8} />
        <YAxis stroke={axis} tick={yTick} tickMargin={8} width={38} />
        <Tooltip contentStyle={tooltip} />
        <Bar dataKey="alerts" fill="#b1843c" radius={[5, 5, 0, 0]} maxBarSize={28} />
        <Line type="monotone" dataKey="closed" stroke="#0b6f68" strokeWidth={2.6} dot={false} />
      </ComposedChart>
    </ResponsiveContainer>
  );
}

export function RadarWeightChart({ data }: { data: { name: string; value: number }[] }) {
  return (
    <ResponsiveContainer width="100%" height={260}>
      <RadarChart data={data}>
        <PolarGrid stroke="rgba(97, 111, 128, .22)" />
        <PolarAngleAxis dataKey="name" stroke="#4a5568" tick={{ fontSize: 11, fill: "#4a5568" }} />
        <Radar dataKey="value" stroke="#0b6f68" fill="#0b6f68" fillOpacity={0.18} />
        <Tooltip contentStyle={tooltip} />
      </RadarChart>
    </ResponsiveContainer>
  );
}

export function DonutChart({ data }: { data: { name: string; value: number }[] }) {
  const total = data.reduce((sum, item) => sum + Number(item.value || 0), 0) || 1;
  let cursor = 0;
  const segments = data.map((item, index) => {
    const start = cursor;
    const end = cursor + (Number(item.value || 0) / total) * 100;
    cursor = end;
    return `${colors[index % colors.length]} ${start}% ${end}%`;
  }).join(", ");
  return (
    <div className="donut-wrap">
      <div className="donut-ring" style={{ background: `conic-gradient(${segments})` }}>
        <div>
          <strong>{total}</strong>
          <span>事件</span>
        </div>
      </div>
      <div className="donut-legend">
        {data.map((item, index) => (
          <div key={item.name}>
            <i style={{ background: colors[index % colors.length] }} />
            <span>{item.name}</span>
            <strong>{item.value}</strong>
          </div>
        ))}
      </div>
    </div>
  );
}
