import { chromium } from "playwright";
import fs from "node:fs/promises";
import path from "node:path";

const root = path.resolve("..");
const out = path.join(root, "deliverables", "SkyGuard_WebDemo_截图包");
await fs.mkdir(out, { recursive: true });
for (const file of await fs.readdir(out).catch(() => [])) {
  if (file.endsWith(".png")) {
    await fs.rm(path.join(out, file), { force: true });
  }
}

const pages = [
  ["/", "01_home.png"],
  ["/product", "02_product.png"],
  ["/technology", "03_technology.png"],
  ["/scenarios", "04_scenarios.png"],
  ["/data", "05_data.png"],
  ["/business", "06_business.png"],
  ["/case-study", "07_case_study.png"],
  ["/demo-center", "08_demo_center.png"],
  ["/dashboard", "09_dashboard.png"],
  ["/dashboard/live", "10_live_tracking.png"],
  ["/dashboard/flight-plans", "11_flight_plans.png"],
  ["/dashboard/recognition-review", "12_recognition_review.png"],
  ["/dashboard/geofence", "13_geofence.png"],
  ["/dashboard/incidents", "14_incidents.png"],
  ["/dashboard/reports", "15_reports.png"],
  ["/dashboard/sensors", "16_sensors.png"],
  ["/dashboard/data-assets", "17_data_assets.png"],
  ["/dashboard/risk-model", "18_risk_model.png"],
  ["/dashboard/settings", "19_settings.png"],
  ["/command-screen", "20_command_screen.png"],
  ["/mobile-alerts", "21_mobile.png", { width: 390, height: 900 }]
];

const browser = await chromium.launch();
for (const [url, file, viewport] of pages) {
  const page = await browser.newPage({ viewport: viewport ?? { width: 1440, height: 1000 } });
  await page.goto(`http://localhost:5173${url}`, { waitUntil: "networkidle" });
  await page.screenshot({ path: path.join(out, file), fullPage: true });
  await page.close();
}
await browser.close();
console.log(out);
