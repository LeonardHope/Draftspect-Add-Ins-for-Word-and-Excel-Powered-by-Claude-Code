// Auto-sideload the Claude Code add-in manifests into Word and Excel.
//
// macOS: drop manifests into the per-host wef/ folder inside the Office
// app's container. Word/Excel scan that folder on launch and surface
// every manifest under Insert → Office Add-ins → SHARED FOLDER.
//
//   ~/Library/Containers/com.microsoft.Word/Data/Documents/wef/word.xml
//   ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/excel.xml
//
// Windows: copy manifests to a folder under %APPDATA%, then register
// that folder as a Trusted Catalog in the Office Trust Center via the
// registry. Office scans trusted-catalog folders for manifests.
//
//   %APPDATA%\Claude Code for Office\manifests\{word,excel}.xml
//   HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\<our-guid>
//
// All operations are idempotent: re-running install is safe, uninstall
// silently ignores already-removed pieces.

import { copyFile, mkdir, unlink, readdir } from "node:fs/promises";
import { existsSync } from "node:fs";
import { join, dirname } from "node:path";
import { homedir } from "node:os";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = join(__dirname, "..");
const MANIFESTS_DIR = join(PROJECT_ROOT, "manifests");

// One row per supported Office host. `mac_container` is the bundle ID whose
// sandbox we drop the manifest into on macOS. `guid` is purely informational
// here (the manifest carries its own Id element).
const HOSTS = [
  { host: "Word",  file: "word.xml",  mac_container: "com.microsoft.Word"  },
  { host: "Excel", file: "excel.xml", mac_container: "com.microsoft.Excel" },
];

// ---------------------------------------------------------------------------
// macOS
// ---------------------------------------------------------------------------
function macWefDir(container) {
  return join(homedir(), "Library", "Containers", container, "Data", "Documents", "wef");
}
function macWefPath(host) {
  return join(macWefDir(host.mac_container), host.file);
}

async function macInstall() {
  const installed = [];
  for (const host of HOSTS) {
    const wef = macWefDir(host.mac_container);
    await mkdir(wef, { recursive: true });
    const src = join(MANIFESTS_DIR, host.file);
    const dst = join(wef, host.file);
    await copyFile(src, dst);
    installed.push({ host: host.host, path: dst });
  }
  return { installed };
}

async function macUninstall() {
  const removed = [];
  for (const host of HOSTS) {
    const p = macWefPath(host);
    try {
      await unlink(p);
      removed.push({ host: host.host, path: p });
    } catch (e) {
      if (e.code !== "ENOENT") throw e;
    }
  }
  return { removed };
}

function macIsInstalled() {
  return HOSTS.every(host => existsSync(macWefPath(host)));
}

// ---------------------------------------------------------------------------
// Windows
// ---------------------------------------------------------------------------
// Catalog id is arbitrary — it's the registry key name under TrustedCatalogs.
// Using a stable string (not a fresh UUID per install) so a second install
// updates the entry in place rather than piling up duplicates.
const WIN_CATALOG_ID = "claude-code-office-trusted-catalog";
const WIN_CATALOG_KEY = `HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs\\${WIN_CATALOG_ID}`;

function winCatalogDir() {
  return join(homedir(), "AppData", "Roaming", "Claude Code for Office", "manifests");
}

function regCommand(args) {
  return new Promise((resolve, reject) => {
    const p = spawn("reg.exe", args, { windowsHide: true });
    let stderr = "";
    p.stderr.on("data", (d) => { stderr += d.toString(); });
    p.on("exit", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`reg.exe ${args.join(" ")} exited ${code}: ${stderr.trim()}`));
    });
    p.on("error", reject);
  });
}

async function regAddString(key, name, value) {
  return regCommand(["add", key, "/v", name, "/t", "REG_SZ", "/d", value, "/f"]);
}
async function regAddDword(key, name, value) {
  return regCommand(["add", key, "/v", name, "/t", "REG_DWORD", "/d", String(value), "/f"]);
}
async function regDelete(key) {
  return regCommand(["delete", key, "/f"]);
}
function regQuery(key) {
  return new Promise((resolve) => {
    const p = spawn("reg.exe", ["query", key], { windowsHide: true });
    p.on("exit", (code) => resolve(code === 0));
    p.on("error", () => resolve(false));
  });
}

async function winInstall() {
  const catalogDir = winCatalogDir();
  await mkdir(catalogDir, { recursive: true });
  const copied = [];
  for (const host of HOSTS) {
    const dst = join(catalogDir, host.file);
    await copyFile(join(MANIFESTS_DIR, host.file), dst);
    copied.push({ host: host.host, path: dst });
  }
  await regAddString(WIN_CATALOG_KEY, "Id",           WIN_CATALOG_ID);
  await regAddString(WIN_CATALOG_KEY, "Url",          catalogDir);
  await regAddString(WIN_CATALOG_KEY, "FriendlyName", "Claude Code for Office");
  // Flags=1 means "Show in menu" — equivalent to the checkbox in the Trust
  // Center UI; without this the catalog is registered but invisible.
  await regAddDword (WIN_CATALOG_KEY, "Flags",        1);
  return { installed: copied, catalog: catalogDir, registry: WIN_CATALOG_KEY };
}

async function winUninstall() {
  const removed = [];
  // Remove registered catalog
  try { await regDelete(WIN_CATALOG_KEY); } catch { /* not present */ }
  // Best-effort clean up the manifest files
  try {
    const dir = winCatalogDir();
    for (const f of await readdir(dir)) {
      try { await unlink(join(dir, f)); removed.push(join(dir, f)); } catch {}
    }
  } catch { /* dir missing — fine */ }
  return { removed, registry: WIN_CATALOG_KEY };
}

async function winIsInstalled() {
  return await regQuery(WIN_CATALOG_KEY);
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------
export async function isAddinInstalled() {
  if (process.platform === "darwin") return macIsInstalled();
  if (process.platform === "win32")  return await winIsInstalled();
  return false;
}

export async function installAddin() {
  if (process.platform === "darwin") return macInstall();
  if (process.platform === "win32")  return winInstall();
  throw new Error(`Auto-sideload not supported on ${process.platform}; see README for manual instructions.`);
}

export async function uninstallAddin() {
  if (process.platform === "darwin") return macUninstall();
  if (process.platform === "win32")  return winUninstall();
  throw new Error(`Auto-sideload not supported on ${process.platform}.`);
}
