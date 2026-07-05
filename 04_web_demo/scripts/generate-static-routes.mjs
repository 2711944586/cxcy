import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const dist = path.resolve(here, "..", "dist");

const routes = [
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

const indexHtml = await fs.readFile(path.join(dist, "index.html"), "utf8");

for (const route of routes) {
  const routeDir = path.join(dist, route.replace(/^\/+/, ""));
  await fs.mkdir(routeDir, { recursive: true });
  await fs.writeFile(path.join(routeDir, "index.html"), indexHtml, "utf8");
}

await fs.writeFile(path.join(dist, "404.html"), indexHtml, "utf8");
console.log(`static route indexes generated: ${routes.length} routes + 404.html`);
