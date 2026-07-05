import { Activity, AlertTriangle, BarChart3, Bell, Building2, CheckCircle2, Clock, Database, FileDown, FileText, Filter, Layers3, MapPinned, Plane, Radar, Radio, Route, Shield, SlidersHorizontal, Smartphone, Upload } from "lucide-react";
import { AreaTrendChart, DonutChart, HorizontalBarChart, MixedOpsChart, RadarWeightChart, SimpleBarChart, TrendChart } from "../components/charts/Charts";
import { LowAltitudeMap } from "../components/map/LowAltitudeMap";
import { KpiCard } from "../components/ui/KpiCard";
import { RiskBadge } from "../components/ui/RiskBadge";
import type { RiskEvent } from "../data/types";
import { useSkyguardData } from "../data/useSkyguardData";
import { compact, pct } from "../utils/format";
import { appPath } from "../utils/navigation";
import { useMemo, useState, type ReactNode } from "react";

function Loading() {
  return (
    <div className="page">
      <div className="skeleton-board" aria-label="正在载入演示数据">
        <div />
        <div />
        <div />
        <div />
      </div>
    </div>
  );
}

function PageTitle({ eyebrow, title, desc, action }: { eyebrow?: string; title: string; desc: string; action?: ReactNode }) {
  return (
    <header className="page-title">
      <div>
        {eyebrow ? <span className="eyebrow">{eyebrow}</span> : null}
        <h2>{title}</h2>
        <p>{desc}</p>
      </div>
      {action ? <div className="title-actions">{action}</div> : null}
    </header>
  );
}

function Toolbar({ children }: { children: ReactNode }) {
  return <div className="toolbar"><Filter size={15} />{children}</div>;
}

function OpsBrief() {
  return (
    <div className="ops-brief">
      <div><span>值班域</span><strong>滨江会展核心区</strong></div>
      <div><span>规则包</span><strong>活动保障 / 临时限飞 / 白名单计划</strong></div>
      <div><span>当前班次</span><strong>09:00-17:00 · 双人复核</strong></div>
    </div>
  );
}

export function Home() {
  const data = useSkyguardData();
  if (data.loading || !data.summary) return <Loading />;
  const topEvent = data.events.find((event) => event.risk_level === "critical") ?? data.events[0];
  return (
    <div className="page ops-home">
      <section className="command-hero">
        <div className="hero-copy">
          <span className="eyebrow">SkyGuard 低空智眼</span>
          <h2>重点区域低空运行监管工作台</h2>
          <p>把飞行计划、实时目标、电子围栏、风险事件和运行报表放在同一条业务链路中。演示数据用于展示流程，行业数据和政策来源在计划书中单独列明。</p>
          <div className="hero-actions">
            <a className="btn primary" href={appPath("/dashboard")}><Radar size={16} />进入综合态势</a>
            <a className="btn" href={appPath("/demo-center")}><Activity size={16} />载入演示脚本</a>
            <a className="btn" href={appPath("/dashboard/reports")}><FileDown size={16} />查看报表</a>
          </div>
        </div>
        <div className="hero-photo">
          <img src={appPath("/assets/asset_01.jpg")} alt="低空飞行服务平台真实界面素材" />
          <div className="photo-label">参考素材：低空飞行服务平台界面</div>
        </div>
      </section>
      <OpsBrief />
      <div className="kpi-strip">
        <KpiCard label="今日飞行计划" value={compact(data.summary.kpis.today_flights)} hint="已提交和已审批计划" icon={<Route size={18} />} />
        <KpiCard label="活动目标" value={compact(data.targets.length)} hint="当前态势图内目标样本" icon={<Radar size={18} />} />
        <KpiCard label="开放事件" value={compact(data.summary.kpis.open_events)} hint="待确认、处理中、待归档" icon={<Bell size={18} />} />
        <KpiCard label="设备在线率" value={pct(data.summary.kpis.sensor_online_rate)} hint="雷达、光电、通信感知接入" icon={<Radio size={18} />} />
      </div>
      <div className="workspace-grid">
        <LowAltitudeMap targets={data.targets} geofences={data.geofences} events={data.events} />
        <EventDetail event={topEvent} />
      </div>
      <OpsConsole events={data.events} />
      <div className="report-grid compact">
        <section className="panel"><h3>近30日运行趋势</h3><TrendChart data={data.summary.daily_trend} /></section>
        <section className="panel"><h3>风险等级结构</h3><DonutChart data={Object.entries(data.summary.risk_level_distribution).map(([name, value]) => ({ name, value }))} /></section>
      </div>
    </div>
  );
}

const productRows = [
  ["综合态势", "目标、围栏、计划、告警同图层展示", "查看态势、定位目标、进入事件"],
  ["计划核验", "飞行计划、白名单、航线和高度约束", "核验计划、比对轨迹、归档结果"],
  ["围栏规则", "禁飞、限飞、重点保护区和临时管制", "配置阈值、触发告警、回写工单"],
  ["事件中心", "确认、派单、处置、复盘和日报", "缩短响应时间，保留证据链"],
  ["报表中心", "日报、月报、重点区域风险评估", "提交验收材料，支持续费复盘"]
];

export function Product() {
  return (
    <div className="page">
      <PageTitle eyebrow="产品方案" title="产品方案" desc="产品不以“大屏效果”作为终点，而以可运行、可追溯、可复盘的低空监管流程作为交付对象。" action={<a className="btn primary" href={appPath("/dashboard")}>打开后台</a>} />
      <div className="two-column">
        <section className="panel image-panel"><img src={appPath("/assets/asset_10.jpg")} alt="低空平台界面素材" /></section>
        <section className="panel">
          <h3>核心模块与客户动作</h3>
          <TableMini headers={["模块", "能力", "客户动作"]} rows={productRows} />
        </section>
      </div>
    </div>
  );
}

export function Technology() {
  const layers = [
    ["数据接入", "飞行计划、目标轨迹、围栏规则、设备告警、人工复核"],
    ["时空索引", "轨迹点、围栏面、起降点、敏感区域统一进入地图坐标"],
    ["风险解释", "计划匹配、身份可信、轨迹偏离、围栏权重、响应紧迫度"],
    ["事件闭环", "自动提示、人工确认、派单处置、报表归档"],
    ["审计安全", "权限分级、导出留痕、操作日志、数据脱敏"]
  ];
  return (
    <div className="page">
      <PageTitle eyebrow="技术架构" title="技术架构" desc="规则先行、模型辅助、人工确认。系统强调可解释和可回放，不把关键处置完全交给黑箱评分。" action={<a className="btn primary" href={appPath("/dashboard/risk-model")}>风险解释器</a>} />
      <div className="architecture">
        {layers.map(([name, body]) => <div className="layer" key={name}><strong>{name}</strong><span>{body}</span></div>)}
      </div>
      <div className="two-column">
        <section className="panel"><img src={appPath("/assets/asset_06.png")} alt="低空感知系统展台素材" /></section>
        <section className="panel"><h3>工程边界</h3><p className="reading">SkyGuard 提供辅助感知、风险预警、事件协同和报表复盘，不做干扰、捕获、打击，也不替代审批或执法。早期试点先接入计划、围栏、轨迹和工单，硬件通过伙伴逐步扩展。</p></section>
      </div>
    </div>
  );
}

export function Scenarios() {
  const scenarios = [
    ["景区核心区", "游客航拍、核心保护区、节假日客流叠加，优先使用围栏和移动处置。"],
    ["园区巡检", "固定航线和设备巡检频繁，适合订阅、运维和报表复盘。"],
    ["大型活动", "会展、赛事、演唱会需要临时限飞和短周期日报。"],
    ["物流航线", "航线评估、计划核验、起降点容量和风险报告是关键交付。"]
  ];
  return (
    <div className="page">
      <PageTitle eyebrow="应用场景" title="应用场景" desc="从责任主体清楚、边界清楚、可验收的区域切入，再进入区县复制。" />
      <div className="scenario-board">
        {scenarios.map(([name, body], index) => <article key={name}><span>0{index + 1}</span><h3>{name}</h3><p>{body}</p></article>)}
      </div>
    </div>
  );
}

export function DataPage() {
  const data = useSkyguardData();
  if (data.loading) return <Loading />;
  return (
    <div className="page">
      <PageTitle eyebrow="数据证据" title="数据证据" desc="官方统计、地方政策、公开数据和演示样本分层管理，避免把演示样本写成真实客户业务。" />
      <div className="kpi-strip">
        <KpiCard label="数据文件" value={data.assets.length} icon={<Database size={18} />} />
        <KpiCard label="目标样本" value={compact(data.targets.length)} icon={<Radar size={18} />} />
        <KpiCard label="飞行计划" value={compact(data.plans.length)} icon={<Route size={18} />} />
        <KpiCard label="事件记录" value={compact(data.events.length)} icon={<Bell size={18} />} />
      </div>
      <section className="panel"><TableMini headers={["文件", "类别", "大小", "用途"]} rows={data.assets.slice(0, 32).map((a) => [a.name, a.category, `${a.size_kb} KB`, a.use])} /></section>
    </div>
  );
}

export function Business() {
  const rows = [
    ["Lite SaaS", "8-15万元/年", "单景区、园区", "账号、围栏、告警、基础报表"],
    ["Pro 部署", "30-80万元/项目", "多区域试点", "数据接入、规则配置、培训验收"],
    ["活动保障", "5-30万元/次", "会展、赛事、节庆", "临时围栏、移动告警、日报复盘"],
    ["航线评估", "3000-20000元/条", "物流、巡检", "路线风险、容量、合规建议"],
    ["运维服务", "合同额10%-20%/年", "持续客户", "规则调优、报表会议、设备巡检"]
  ];
  return (
    <div className="page">
      <PageTitle eyebrow="商业测算" title="商业模式" desc="收入来自订阅、部署、活动保障、航线评估、报告服务和运维续费，核心是把一次试点转成持续服务。" action={<a className="btn primary" href={appPath("/dashboard/reports")}>经营报表</a>} />
      <section className="panel"><TableMini headers={["收入项", "价格口径", "适用客户", "交付内容"]} rows={rows} /></section>
    </div>
  );
}

export function CaseStudy() {
  return (
    <div className="page">
      <PageTitle eyebrow="试点案例" title="重点区域试点案例" desc="以会展核心区为样例，展示部署、运行、处置和复盘四个阶段。" action={<a className="btn primary" href={appPath("/demo-center")}>载入脚本</a>} />
      <div className="case-timeline">
        {["划定重点区域与临时限飞规则", "导入白名单计划和感知设备", "未知目标触发高风险工单", "现场确认后归档并生成日报"].map((step, i) => (
          <div className="case-step" key={step}><span>{i + 1}</span><strong>{step}</strong></div>
        ))}
      </div>
      <div className="two-column"><section className="panel"><img src={appPath("/assets/asset_03.jpg")} alt="低空综合监管平台素材" /></section><section className="panel"><h3>验收口径</h3><p className="reading">试点验收关注计划匹配率、告警确认时长、工单闭环率、设备在线率和日报输出质量。系统不承诺替代执法，只提供辅助监管和复盘材料。</p></section></div>
    </div>
  );
}

export function DemoCenter() {
  const [active, setActive] = useState(0);
  const scripts = [
    ["未知目标进入重点保护区", ["目标出现", "比对计划库", "围栏触发", "派发工单", "移动端确认"]],
    ["合规物流航线运行", ["计划审批", "航线匹配", "轨迹跟踪", "自动归档", "生成日报"]],
    ["大型活动临时限飞", ["创建临时围栏", "导入名单", "现场告警", "处置记录", "复盘报表"]],
    ["设备离线降级处置", ["心跳异常", "置信度下降", "人工复核", "设备工单", "恢复确认"]],
    ["月度经营复盘", ["汇总飞行", "统计风险", "分析响应", "检查续费", "导出报告"]]
  ];
  return (
    <div className="page">
      <PageTitle eyebrow="演示脚本" title="演示脚本" desc="按真实业务路径组织演示：态势、计划、围栏、事件、移动端和报表连续点击。" action={<a className="btn primary" href={appPath("/dashboard")}>打开态势</a>} />
      <div className="script-list">
        {scripts.map(([title, steps], i) => (
          <button className={`script-card ${active === i ? "active" : ""}`} key={title as string} onClick={() => setActive(i)}>
            <span>脚本 {i + 1}</span>
            <h3>{title as string}</h3>
            <p>{(steps as string[]).slice(0, 3).join(" / ")}</p>
          </button>
        ))}
      </div>
      <section className="panel demo-runner">
        <div><span className="badge high">当前脚本</span><h2>{scripts[active][0] as string}</h2></div>
        <div className="timeline">{(scripts[active][1] as string[]).map((step, i) => <div className="timeline-step" key={step}><span>{i + 1}</span><strong>{step}</strong></div>)}</div>
        <div className="flow"><a className="btn primary" href={appPath("/dashboard")}>综合态势</a><a className="btn" href={appPath("/dashboard/incidents")}>事件工单</a><a className="btn" href={appPath("/mobile-alerts")}>移动处置</a><a className="btn" href={appPath("/dashboard/reports")}>复盘报表</a></div>
      </section>
    </div>
  );
}

export function About() {
  const rows = [
    ["产品与行业", "场景定义、客户访谈、试点方案、商业材料"],
    ["数据与算法", "样本治理、风险字段、评分解释、模型校验"],
    ["前后端研发", "工作台、地图、表格、报表和部署包"],
    ["交付与伙伴", "设备接入、现场配置、培训、客户成功"]
  ];
  return <div className="page"><PageTitle eyebrow="团队组织" title="团队组织" desc="团队围绕演示、试点和可交付材料配置，不把早期项目写成臃肿组织。" /><section className="panel"><TableMini headers={["角色", "职责"]} rows={rows} /></section></div>;
}

export function Dashboard() {
  const data = useSkyguardData();
  if (data.loading || !data.summary) return <Loading />;
  return (
    <div className="page">
      <OpsBrief />
      <div className="kpi-strip">
        <KpiCard label="今日飞行" value={compact(data.summary.kpis.today_flights)} icon={<Plane size={18} />} />
        <KpiCard label="活动目标" value={compact(data.targets.length)} icon={<Radar size={18} />} />
        <KpiCard label="开放事件" value={compact(data.summary.kpis.open_events)} icon={<Bell size={18} />} />
        <KpiCard label="平均响应" value={`${data.summary.kpis.avg_response_min} 分`} icon={<CheckCircle2 size={18} />} />
      </div>
      <div className="workspace-grid"><LowAltitudeMap targets={data.targets} geofences={data.geofences} events={data.events} /><EventPanel /></div>
      <OpsConsole events={data.events} />
      <div className="report-grid compact"><section className="panel"><h3>30 日运行趋势</h3><TrendChart data={data.summary.daily_trend} /></section><section className="panel"><h3>风险等级分布</h3><DonutChart data={Object.entries(data.summary.risk_level_distribution).map(([name, value]) => ({ name, value }))} /></section></div>
    </div>
  );
}

export function LiveTracking() {
  const data = useSkyguardData();
  const [level, setLevel] = useState("全部");
  if (data.loading) return <Loading />;
  const rows = data.targets.filter((t) => level === "全部" || t.risk_level === level).slice(0, 58).map((t) => [t.drone_id, t.target_type, `${t.altitude}m`, `${t.speed}m/s`, t.current_status, <RiskBadge level={t.risk_level} />, t.last_seen_time]);
  return <TablePage title="目标监测" desc="目标编号、类型、高度、速度、状态、风险等级和最近更新时间。" toolbar={<Toolbar><button className={level === "全部" ? "seg active" : "seg"} onClick={() => setLevel("全部")}>全部</button><button className={level === "high" ? "seg active" : "seg"} onClick={() => setLevel("high")}>高风险</button><button className={level === "critical" ? "seg active" : "seg"} onClick={() => setLevel("critical")}>严重</button></Toolbar>} rows={rows} headers={["目标", "类型", "高度", "速度", "状态", "风险", "最后出现"]} />;
}

export function FlightPlans() {
  const data = useSkyguardData();
  if (data.loading) return <Loading />;
  return <TablePage title="计划审批" desc="展示飞行计划审批、任务场景、计划高度、距离和风险评分。" rows={data.plans.slice(0, 64).map((p) => [p.plan_id, p.city, p.scenario, p.approval_status, `${p.approved_altitude_max}m`, `${p.planned_distance_km}km`, p.risk_score])} headers={["计划", "城市", "场景", "审批", "高度", "距离", "风险分"]} />;
}

export function RecognitionReview() {
  const [fileName, setFileName] = useState("低空平台样例图");
  return (
    <div className="page">
      <PageTitle eyebrow="智能复核" title="识别复核" desc="上传图片后显示模拟检测框、置信度和风险建议。结果用于演示，不作为真实执法结论。" action={<label className="btn primary"><Upload size={16} />上传图片<input type="file" hidden onChange={(e) => setFileName(e.target.files?.[0]?.name ?? "未选择文件")} /></label>} />
      <div className="two-column">
        <section className="panel detection-frame"><img src={appPath("/assets/asset_03.jpg")} alt="目标识别样例" /><div className="detection-box" /><span className="detect-tag">无人机 0.91</span></section>
        <section className="panel"><h3>复核建议</h3><p className="muted">输入文件：{fileName}</p><p><span className="badge high">需人工确认</span></p><p className="reading">目标疑似低空航拍无人机，建议与飞行计划库、Remote ID、电子围栏和现场人员确认结果进行比对。系统只给出辅助建议，最终处置保留人工确认。</p><div className="flow"><span className="flow-step">图像输入</span><span className="flow-step">检测框</span><span className="flow-step">风险建议</span><span className="flow-step">人工确认</span></div></section>
      </div>
    </div>
  );
}

export function GeoFence() {
  const data = useSkyguardData();
  const [triggered, setTriggered] = useState(false);
  if (data.loading) return <Loading />;
  return (
    <div className="page">
      <PageTitle eyebrow="电子围栏" title="围栏规则" desc="管理禁飞区、限飞区、重点保护区和普通监测区，并演示目标进入围栏后的事件生成。" action={<button className="btn primary" onClick={() => setTriggered(true)}><MapPinned size={16} />触发模拟告警</button>} />
      <div className="workspace-grid">
        <LowAltitudeMap targets={data.targets.slice(10)} geofences={data.geofences} events={data.events.slice(0, 8)} />
        <div className="stack">{triggered && <div className="alert-card"><AlertTriangle size={18} /><div><strong>重点保护区越界</strong><p>UAV-037 进入临时限飞区，系统已生成待确认工单。</p></div></div>}<TableMini rows={data.geofences.slice(0, 14).map((g) => [g.fence_name, g.fence_type, g.risk_weight, g.active ? "启用" : "停用"])} headers={["围栏", "类型", "权重", "状态"]} /></div>
      </div>
    </div>
  );
}

export function Incidents() {
  const data = useSkyguardData();
  const [selected, setSelected] = useState<RiskEvent | null>(null);
  if (data.loading) return <Loading />;
  return (
    <div className="page">
      <PageTitle eyebrow="事件工单" title="事件工单" desc="待确认、已确认、处理中、已处置、已归档的闭环事件。" />
      <div className="incident-layout">
        <section className="panel"><TableMini headers={["事件", "城市", "类型", "等级", "状态", "响应"]} rows={data.events.slice(0, 66).map((e) => [<button className="linklike" onClick={() => setSelected(e)}>{e.event_id}</button>, e.city, e.event_type, <RiskBadge level={e.risk_level} />, e.status, `${e.response_time_min}分`])} /></section>
        <EventDetail event={selected ?? data.events[0]} />
      </div>
    </div>
  );
}

export function Reports() {
  const data = useSkyguardData();
  const [exportedAt, setExportedAt] = useState<string | null>(null);
  if (data.loading || !data.summary) return <Loading />;
  const cityData = data.summary.city_risk_top.slice(0, 8).map(([name, value]) => ({ name, value }));
  const eventData = data.summary.event_type_top.slice(0, 8).map(([name, value]) => ({ name, value }));
  const weightData = [{ name: "区域", value: 82 }, { name: "合规", value: 76 }, { name: "轨迹", value: 69 }, { name: "身份", value: 61 }, { name: "环境", value: 48 }, { name: "响应", value: 74 }];
  const handleExport = () => {
    setExportedAt(new Date().toLocaleString("zh-CN", { hour12: false }));
    window.setTimeout(() => window.print(), 80);
  };
  return (
    <div className="page">
      <PageTitle eyebrow="运行报表" title="运行报表" desc="日报、周报、月报、重点区域报告、航线风险评估和设备运维报告共享同一套字段口径。" action={<button type="button" className="btn primary" onClick={handleExport}><FileDown size={16} />导出报告</button>} />
      <section className="report-summary" aria-live="polite">
        <div><FileText size={17} /><span>报表口径</span><strong>日报 / 周报 / 月报</strong></div>
        <div><CheckCircle2 size={17} /><span>平均响应</span><strong>{data.summary.kpis.avg_response_min} 分钟</strong></div>
        <div><AlertTriangle size={17} /><span>开放事件</span><strong>{compact(data.summary.kpis.open_events)} 件</strong></div>
        {exportedAt ? <div className="report-status"><FileDown size={17} /><span>导出记录</span><strong>{exportedAt}</strong></div> : null}
      </section>
      <div className="report-grid">
        <section className="panel"><h3>运行趋势</h3><TrendChart data={data.summary.daily_trend} /></section>
        <section className="panel"><h3>风险解释雷达</h3><RadarWeightChart data={weightData} /></section>
        <section className="panel"><h3>飞行与告警面积</h3><AreaTrendChart data={data.summary.daily_trend} /></section>
        <section className="panel"><h3>区域风险排行</h3><HorizontalBarChart data={cityData} /></section>
        <section className="panel"><h3>告警类型排行</h3><SimpleBarChart data={eventData} /></section>
        <section className="panel"><h3>处置闭环组合图</h3><MixedOpsChart data={data.summary.daily_trend} /></section>
      </div>
    </div>
  );
}

export function Sensors() {
  const data = useSkyguardData();
  if (data.loading) return <Loading />;
  return <TablePage title="感知设备" desc="传感器列表、在线状态、心跳时间、覆盖范围和识别准确率。" rows={data.sensors.slice(0, 70).map((s) => [s.sensor_id, s.sensor_type, s.location, s.online_status, pct(s.accuracy), `${s.coverage_radius_m}m`, s.last_heartbeat])} headers={["设备", "类型", "位置", "状态", "准确率", "覆盖", "心跳"]} />;
}

export function DataAssetsDashboard() { return <DataPage />; }

export function RiskModel() {
  const [area, setArea] = useState(25);
  const [track, setTrack] = useState(20);
  const factors = useMemo(() => [{ name: "区域敏感度", value: area }, { name: "飞行合规性", value: 20 }, { name: "轨迹异常度", value: track }, { name: "目标可信度", value: 15 }, { name: "环境影响", value: 10 }, { name: "响应紧迫度", value: 10 }], [area, track]);
  const score = Math.round(area * 0.92 + track * 1.15 + 36);
  return (
    <div className="page">
      <PageTitle eyebrow="风险解释" title="风险解释" desc="展示六维权重、样例目标评分过程和分级处置策略，保留人工确认入口。" />
      <div className="two-column">
        <section className="panel"><h3>六维权重</h3><SimpleBarChart data={factors} /></section>
        <section className="panel"><h3>样例目标 UAV-037</h3><p className="reading">区域敏感度高、无有效飞行计划、靠近重点保护区。调整权重后，系统重算风险分并保留确认记录。</p><Control label="区域敏感度" value={area} setValue={setArea} /><Control label="轨迹异常度" value={track} setValue={setTrack} /><div className="score-card"><Clock size={18} /><span>当前综合评分</span><strong>{score}</strong><RiskBadge level={score > 80 ? "critical" : "high"} /></div></section>
      </div>
    </div>
  );
}

function Control({ label, value, setValue }: { label: string; value: number; setValue: (value: number) => void }) {
  return <div className="control-row"><SlidersHorizontal size={16} /><label>{label}<input type="range" min="10" max="35" value={value} onChange={(e) => setValue(Number(e.target.value))} /></label><strong>{value}%</strong></div>;
}

export function Settings() {
  const rows = [["角色权限", "管理员、值班员、企业用户、审计角色分离"], ["区域规则", "围栏、阈值、告警等级和通知人按场景包维护"], ["数据安全", "日志、脱敏、导出审批和操作留痕"], ["部署方式", "本地静态部署、内网演示、云端预览均可"]];
  return <div className="page"><PageTitle eyebrow="系统配置" title="系统配置" desc="权限、区域、告警、脱敏、审计和部署能力用于支撑可试点性。" /><section className="panel"><TableMini headers={["配置项", "说明"]} rows={rows} /></section></div>;
}

export function MobileAlerts() {
  const data = useSkyguardData();
  if (data.loading) return <Loading />;
  return (
    <div className="page mobile-page">
      <PageTitle eyebrow="移动处置" title="移动处置" desc="现场人员接收、确认和处置低空风险事件，适合展示现场协同。" action={<a className="btn primary" href={appPath("/dashboard/incidents")}><Smartphone size={16} />事件中心</a>} />
      <div className="phone-shell"><div className="phone-header">滨江会展核心区</div><TableMini rows={data.events.slice(0, 18).map((e) => [e.event_type, <RiskBadge level={e.risk_level} />, e.status])} headers={["事件", "等级", "状态"]} /></div>
    </div>
  );
}

export function CommandScreen() {
  const data = useSkyguardData();
  if (data.loading || !data.summary) return <Loading />;
  return (
    <div className="page command-page">
      <PageTitle eyebrow="指挥大屏" title="低空指挥大屏" desc="适合全屏展示的态势图、关键指标和实时告警。" />
      <div className="kpi-strip command-kpis"><KpiCard label="飞行计划" value={data.summary.kpis.today_flights} icon={<Route size={18} />} /><KpiCard label="风险事件" value={data.events.length} icon={<Bell size={18} />} /><KpiCard label="围栏数量" value={data.geofences.length} icon={<Layers3 size={18} />} /><KpiCard label="传感设备" value={data.sensors.length} icon={<Radio size={18} />} /></div>
      <div className="command-screen-grid">
        <LowAltitudeMap targets={data.targets} geofences={data.geofences} events={data.events} />
        <CommandEventRail events={data.events} />
      </div>
      <div className="command-ticker" aria-label="运行状态滚动条">
        <span>双人复核开启</span>
        <span>临时限飞规则 14 条</span>
        <span>高风险事件优先派单</span>
        <span>日报字段已同步</span>
      </div>
    </div>
  );
}

function OpsConsole({ events }: { events: RiskEvent[] }) {
  const steps = [
    ["发现", events[0]?.event_time?.slice(11, 16) ?? "09:12", "目标进入重点区域"],
    ["核验", events[1]?.event_time?.slice(11, 16) ?? "09:18", "计划库与围栏规则比对"],
    ["派单", events[2]?.event_time?.slice(11, 16) ?? "09:26", "现场人员接收工单"],
    ["归档", events[3]?.event_time?.slice(11, 16) ?? "09:43", "事件进入日报口径"]
  ];
  const queue = events.slice(0, 5);
  return (
    <section className="ops-console" aria-label="值班节奏和处置队列">
      <div className="ops-timeline">
        {steps.map(([label, time, detail], index) => (
          <div className="ops-timeline-step" key={label}>
            <span>{String(index + 1).padStart(2, "0")}</span>
            <strong>{label}</strong>
            <small>{time} · {detail}</small>
          </div>
        ))}
      </div>
      <div className="triage-queue">
        <div className="queue-head"><span className="eyebrow">待处理</span><strong>待处理队列</strong></div>
        {queue.map((event) => (
          <div className="queue-row" key={event.event_id}>
            <div><strong>{event.event_type}</strong><span>{event.city} · {event.status}</span></div>
            <RiskBadge level={event.risk_level} />
          </div>
        ))}
      </div>
    </section>
  );
}

function CommandEventRail({ events }: { events: RiskEvent[] }) {
  return (
    <aside className="command-rail" aria-label="指挥屏事件侧栏">
      <div className="command-rail-head">
        <span className="eyebrow">实时队列</span>
        <strong>{events.filter((event) => !event.closed_loop).length} open</strong>
      </div>
      {events.slice(0, 8).map((event) => (
        <div className="command-rail-row" key={event.event_id}>
          <RiskBadge level={event.risk_level} />
          <div>
            <strong>{event.event_type}</strong>
            <span>{event.city} · {event.response_time_min} 分钟 · {event.status}</span>
          </div>
        </div>
      ))}
    </aside>
  );
}

function EventPanel() {
  const data = useSkyguardData();
  return <section className="panel event-list"><h3>实时告警</h3>{data.events.slice(0, 12).map((e) => <div key={e.event_id} className="event-row"><div><strong>{e.event_type}</strong><span>{e.city} · {e.status} · {e.response_time_min} 分</span></div><RiskBadge level={e.risk_level} /></div>)}</section>;
}

function EventDetail({ event }: { event?: RiskEvent }) {
  if (!event) return <section className="panel"><h3>事件详情</h3><p className="muted">暂无事件。</p></section>;
  return (
    <aside className="panel event-detail">
      <div className="detail-head"><span className="eyebrow">事件详情</span><RiskBadge level={event.risk_level} /></div>
      <h3>{event.event_type}</h3>
      <dl><dt>事件编号</dt><dd>{event.event_id}</dd><dt>城市</dt><dd>{event.city}</dd><dt>来源</dt><dd>{event.source}</dd><dt>状态</dt><dd>{event.status}</dd><dt>响应时间</dt><dd>{event.response_time_min} 分钟</dd></dl>
      <div className="flow"><span className="flow-step">发现</span><span className="flow-step">核验</span><span className="flow-step">派单</span><span className="flow-step">归档</span></div>
    </aside>
  );
}

function TableMini({ headers, rows }: { headers: string[]; rows: ReactNode[][] | string[][] }) {
  return <div className="table-wrap"><table><thead><tr>{headers.map((h) => <th key={h}>{h}</th>)}</tr></thead><tbody>{rows.map((row, i) => <tr key={i}>{row.map((cell, j) => <td key={j}>{cell}</td>)}</tr>)}</tbody></table></div>;
}

function TablePage({ title, desc, headers, rows, toolbar }: { title: string; desc: string; headers: string[]; rows: ReactNode[][]; toolbar?: ReactNode }) {
  return <div className="page"><PageTitle eyebrow="运行工作台" title={title} desc={desc} action={toolbar} /><section className="panel"><TableMini headers={headers} rows={rows} /></section></div>;
}
