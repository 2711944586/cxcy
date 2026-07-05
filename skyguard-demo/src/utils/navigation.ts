const rawBase = ((import.meta as ImportMeta & { env: { BASE_URL?: string } }).env.BASE_URL ?? "/").replace(/\/$/, "");

export const basePath = rawBase === "/" ? "" : rawBase;

export function appPath(path: string) {
  if (path.startsWith("#") || /^https?:\/\//.test(path)) {
    return path;
  }
  const normalized = path.startsWith("/") ? path : `/${path}`;
  return `${basePath}${normalized}` || normalized;
}

export function currentAppPath() {
  const { pathname } = window.location;
  if (basePath && pathname === basePath) {
    return "/";
  }
  if (basePath && pathname.startsWith(`${basePath}/`)) {
    return pathname.slice(basePath.length) || "/";
  }
  return pathname || "/";
}

