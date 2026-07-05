import { Activity, BarChart3, Bell, Binary, Building2, Database, FileText, Gauge, Home, Map, Plane, Radar, Route, Search, Settings, Shield, Smartphone, Users } from "lucide-react";
import type { ReactNode } from "react";
import { NavLink } from "./routing";
import { appPath } from "../../utils/navigation";

const nav = [
  { group: "计划书", items: [
    ["/", "项目总览", Home],
    ["/product", "产品方案", Shield],
    ["/technology", "技术架构", Binary],
    ["/scenarios", "场景策略", Building2],
    ["/data", "数据证据", Database],
    ["/business", "商业测算", BarChart3],
    ["/case-study", "试点案例", FileText],
    ["/demo-center", "演示脚本", Activity],
    ["/about", "团队组织", Users]
  ]},
  { group: "运行工作台", items: [
    ["/dashboard", "综合态势", Gauge],
    ["/dashboard/live", "目标监测", Radar],
    ["/dashboard/flight-plans", "计划审批", Route],
    ["/dashboard/recognition-review", "识别复核", Plane],
    ["/dashboard/geofence", "围栏规则", Map],
    ["/dashboard/incidents", "事件工单", Bell],
    ["/dashboard/reports", "运行报表", BarChart3],
    ["/dashboard/sensors", "感知设备", Radar],
    ["/dashboard/data-assets", "数据资产", Database],
    ["/dashboard/risk-model", "风险解释", Activity],
    ["/dashboard/settings", "系统配置", Settings]
  ]},
  { group: "现场端", items: [
    ["/mobile-alerts", "移动处置", Smartphone],
    ["/command-screen", "指挥大屏", Shield]
  ]}
] as const;

const quickLinks = [
  ["/dashboard", "态势", Gauge],
  ["/dashboard/incidents", "工单", Bell],
  ["/dashboard/reports", "报表", BarChart3],
  ["/mobile-alerts", "现场端", Smartphone]
] as const;

export function Shell({ children, title }: { children: ReactNode; title: string }) {
  const now = new Date();
  const time = now.toLocaleString("zh-CN", { hour12: false, month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });

  return (
    <div className="app-shell">
      <a className="skip-link" href="#main-content">跳到主内容</a>
      <aside className="sidebar" aria-label="SkyGuard 主导航">
        <div className="brand">
          <div className="brand-mark"><Shield size={18} /></div>
          <div>
            <div className="brand-title">SkyGuard</div>
            <div className="brand-sub">低空智眼运行平台</div>
          </div>
        </div>
        <div className="ops-card">
          <span>当前运行域</span>
          <strong>滨江会展核心区</strong>
          <small>计划核验 · 围栏规则 · 工单闭环 · 报表复盘</small>
          <div className="ops-meter" aria-hidden="true"><i /></div>
        </div>
        {nav.map((group) => (
          <nav className="nav-group" key={group.group} aria-label={group.group}>
            <div className="nav-title">{group.group}</div>
            {group.items.map(([href, label, Icon]) => (
              <NavLink key={href} href={href}><Icon size={15} />{label}</NavLink>
            ))}
          </nav>
        ))}
      </aside>
      <main className="main" id="main-content">
        <div className="topbar">
          <div>
            <span className="topbar-label">低空运行值班台</span>
            <h1>{title}</h1>
          </div>
          <div className="topbar-actions" aria-label="系统状态">
            <span className="status-dot" aria-hidden="true" />
            <span className="topbar-meta">演示数据已同步 · {time}</span>
            <a className="btn primary" href={appPath("/demo-center")}><Activity size={16} />载入脚本</a>
          </div>
        </div>
        <div className="command-strip" aria-label="快速运行入口">
          <div className="command-search">
            <Search size={15} />
            <span>滨江会展核心区 / UAV-037 / 临时限飞 / 白名单计划</span>
          </div>
          <div className="quick-links">
            {quickLinks.map(([href, label, Icon]) => (
              <a key={href} href={appPath(href)}><Icon size={14} />{label}</a>
            ))}
          </div>
          <div className="system-chips" aria-label="运行状态">
            <span>双人复核开启</span>
            <span>设备在线 91.6%</span>
            <span>规则包 v0.9.4</span>
          </div>
        </div>
        {children}
      </main>
    </div>
  );
}
