import { chromium } from "playwright";
import fs from "node:fs/promises";
import path from "node:path";

const base = process.env.SKYGUARD_BASE_URL ?? "http://127.0.0.1:5173";
const root = path.resolve("..");
const outDir = path.join(root, "deliverables");
const outFile = path.join(outDir, "SkyGuard_WebDemo_路由巡检报告.json");

const routes = [
  "/",
  "/product",
  "/technology",
  "/scenarios",
  "/data",
  "/business",
  "/case-study",
  "/demo-center",
  "/about",
  "/dashboard",
  "/dashboard/live",
  "/dashboard/flight-plans",
  "/dashboard/recognition-review",
  "/dashboard/geofence",
  "/dashboard/incidents",
  "/dashboard/reports",
  "/dashboard/sensors",
  "/dashboard/data-assets",
  "/dashboard/risk-model",
  "/dashboard/settings",
  "/command-screen",
  "/mobile-alerts"
];

function isIgnorableRequest(url) {
  return url.startsWith("data:") || url.startsWith("blob:");
}

const browser = await chromium.launch();
const results = [];

for (const route of routes) {
  const page = await browser.newPage({
    viewport: route === "/mobile-alerts" ? { width: 390, height: 900 } : { width: 1440, height: 1000 }
  });
  const messages = [];
  const failedRequests = [];
  const pageErrors = [];

  page.on("console", (message) => {
    if (["error", "warning"].includes(message.type())) {
      messages.push({ type: message.type(), text: message.text() });
    }
  });
  page.on("pageerror", (error) => {
    pageErrors.push(error.message);
  });
  page.on("requestfailed", (request) => {
    const url = request.url();
    if (!isIgnorableRequest(url)) {
      failedRequests.push({ url, failure: request.failure()?.errorText ?? "unknown" });
    }
  });

  const response = await page.goto(`${base}${route}`, { waitUntil: "networkidle" });
  const title = await page.locator("h1, h2").first().textContent().catch(() => "");
  const bodyText = await page.locator("body").innerText().catch(() => "");
  const status = response?.status() ?? 0;

  results.push({
    route,
    status,
    title: title?.trim() ?? "",
    bodyLength: bodyText.length,
    consoleIssues: messages,
    pageErrors,
    failedRequests,
    ok: status >= 200 && status < 400 && bodyText.length > 200 && messages.length === 0 && pageErrors.length === 0 && failedRequests.length === 0
  });

  await page.close();
}

await browser.close();

const summary = {
  generatedAt: new Date().toISOString(),
  base,
  routeCount: routes.length,
  passed: results.filter((item) => item.ok).length,
  failed: results.filter((item) => !item.ok).length,
  results
};

await fs.mkdir(outDir, { recursive: true });
await fs.writeFile(outFile, JSON.stringify(summary, null, 2), "utf8");
console.log(JSON.stringify(summary, null, 2));

if (summary.failed > 0) {
  process.exitCode = 1;
}
