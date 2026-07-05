import { Shell } from "./components/layout/Shell";
import { About, Business, CaseStudy, CommandScreen, Dashboard, DataAssetsDashboard, DataPage, DemoCenter, FlightPlans, GeoFence, Home, Incidents, LiveTracking, MobileAlerts, Product, RecognitionReview, Reports, RiskModel, Scenarios, Sensors, Settings, Technology } from "./routes/pages";
import { currentAppPath } from "./utils/navigation";
import type { ReactNode } from "react";

const routes: Record<string, { title: string; element: ReactNode }> = {
  "/": { title: "首页", element: <Home /> },
  "/product": { title: "产品方案", element: <Product /> },
  "/technology": { title: "技术能力", element: <Technology /> },
  "/scenarios": { title: "应用场景", element: <Scenarios /> },
  "/data": { title: "数据资产", element: <DataPage /> },
  "/business": { title: "商业模式", element: <Business /> },
  "/case-study": { title: "模拟案例", element: <CaseStudy /> },
  "/demo-center": { title: "演示中心", element: <DemoCenter /> },
  "/about": { title: "团队合作", element: <About /> },
  "/dashboard": { title: "低空态势总览", element: <Dashboard /> },
  "/dashboard/live": { title: "实时目标监测", element: <LiveTracking /> },
  "/dashboard/flight-plans": { title: "飞行计划管理", element: <FlightPlans /> },
  "/dashboard/recognition-review": { title: "智能识别复核", element: <RecognitionReview /> },
  "/dashboard/geofence": { title: "电子围栏", element: <GeoFence /> },
  "/dashboard/incidents": { title: "事件处置中心", element: <Incidents /> },
  "/dashboard/reports": { title: "报表中心", element: <Reports /> },
  "/dashboard/sensors": { title: "设备管理", element: <Sensors /> },
  "/dashboard/data-assets": { title: "数据资产中心", element: <DataAssetsDashboard /> },
  "/dashboard/risk-model": { title: "风险模型解释器", element: <RiskModel /> },
  "/dashboard/settings": { title: "系统设置", element: <Settings /> },
  "/mobile-alerts": { title: "移动处置", element: <MobileAlerts /> },
  "/command-screen": { title: "指挥大屏", element: <CommandScreen /> }
};

export function App() {
  const route = routes[currentAppPath()] ?? routes["/"];
  return <Shell title={route.title}>{route.element}</Shell>;
}
