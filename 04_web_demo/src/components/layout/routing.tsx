import type { ReactNode } from "react";
import { appPath, currentAppPath } from "../../utils/navigation";

export function NavLink({ href, children }: { href: string; children: ReactNode }) {
  const active = currentAppPath() === href;
  return <a className={`nav-link ${active ? "active" : ""}`} href={appPath(href)}>{children}</a>;
}
