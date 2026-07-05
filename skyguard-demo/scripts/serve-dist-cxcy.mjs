import http from "node:http";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const dist = path.resolve(here, "..", "dist");
const port = Number(process.env.PORT ?? 4174);
const base = "/cxcy";

const mime = new Map([
  [".html", "text/html; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".png", "image/png"],
  [".jpg", "image/jpeg"],
  [".jpeg", "image/jpeg"],
  [".webp", "image/webp"],
  [".csv", "text/csv; charset=utf-8"],
  [".txt", "text/plain; charset=utf-8"],
  [".md", "text/markdown; charset=utf-8"],
  [".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"],
  [".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"],
  [".gz", "application/gzip"]
]);

function safeJoin(root, requestPath) {
  const target = path.resolve(root, requestPath.replace(/^\/+/, ""));
  if (!target.startsWith(root)) {
    return null;
  }
  return target;
}

async function sendFile(res, file) {
  const data = await fs.readFile(file);
  res.writeHead(200, { "content-type": mime.get(path.extname(file).toLowerCase()) ?? "application/octet-stream" });
  res.end(data);
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url ?? "/", `http://${req.headers.host ?? "127.0.0.1"}`);
    if (!url.pathname.startsWith(base)) {
      res.writeHead(302, { location: `${base}/` });
      res.end();
      return;
    }
    const relative = decodeURIComponent(url.pathname.slice(base.length)) || "/index.html";
    const candidate = safeJoin(dist, relative === "/" ? "/index.html" : relative);
    if (candidate) {
      try {
        const stat = await fs.stat(candidate);
        if (stat.isFile()) {
          await sendFile(res, candidate);
          return;
        }
      } catch {
        // Fall through to SPA fallback.
      }
    }
    await sendFile(res, path.join(dist, "index.html"));
  } catch (error) {
    res.writeHead(500, { "content-type": "text/plain; charset=utf-8" });
    res.end(error instanceof Error ? error.message : "server error");
  }
});

server.listen(port, "127.0.0.1", () => {
  console.log(`serving ${dist} at http://127.0.0.1:${port}${base}/`);
});
