import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const demoRoot = path.resolve(here, "..");
const root = path.resolve(demoRoot, "..");
const dist = path.join(demoRoot, "dist");

const sources = [
  {
    type: "data",
    from: path.join(root, "02_source_assets", "_整理", "数据"),
    to: path.join(dist, "source-data", "data")
  },
  {
    type: "image",
    from: path.join(root, "02_source_assets", "_整理", "图"),
    to: path.join(dist, "source-data", "images")
  },
  {
    type: "reference",
    from: path.join(root, "01_references"),
    to: path.join(dist, "source-data", "references")
  }
];

async function exists(target) {
  try {
    await fs.access(target);
    return true;
  } catch {
    return false;
  }
}

async function copyRecursive(from, to) {
  await fs.mkdir(to, { recursive: true });
  const entries = await fs.readdir(from, { withFileTypes: true });
  const copied = [];
  for (const entry of entries) {
    const src = path.join(from, entry.name);
    const dest = path.join(to, entry.name);
    if (entry.isDirectory()) {
      copied.push(...await copyRecursive(src, dest));
    } else if (entry.isFile()) {
      await fs.copyFile(src, dest);
      const stat = await fs.stat(src);
      copied.push({ name: entry.name, path: path.relative(to, dest).replaceAll(path.sep, "/"), size_kb: Number((stat.size / 1024).toFixed(1)) });
    }
  }
  return copied;
}

const manifest = {
  generated_at: new Date().toISOString(),
  data: [],
  images: [],
  references: []
};

for (const source of sources) {
  if (!await exists(source.from)) {
    console.warn(`skip missing source: ${source.from}`);
    continue;
  }
  await fs.rm(source.to, { recursive: true, force: true });
  const files = await copyRecursive(source.from, source.to);
  manifest[source.type === "image" ? "images" : source.type === "reference" ? "references" : "data"] = files.map((file) => ({
    ...file,
    type: source.type
  }));
  console.log(`${source.type}: copied ${files.length} files`);
}

await fs.mkdir(path.join(dist, "source-data"), { recursive: true });
await fs.writeFile(path.join(dist, "source-data", "manifest.json"), JSON.stringify(manifest, null, 2), "utf8");
console.log("source-data manifest generated");
