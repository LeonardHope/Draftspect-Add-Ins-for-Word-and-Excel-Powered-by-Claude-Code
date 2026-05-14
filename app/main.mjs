// Office Claude — Electron menu bar shell.
//
// Wraps the daemon as a child process, exposes a tray icon with status and
// controls, hides the dock icon (we're a background-only app), and restarts
// the daemon if it crashes. The daemon code itself is untouched.

import { app, Tray, Menu, shell, dialog, nativeImage } from "electron";
import { spawn } from "node:child_process";
import { readFile, stat } from "node:fs/promises";
import fs from "node:fs";
import { homedir } from "node:os";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { isAddinInstalled, installAddin, uninstallAddin } from "./sideload.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = join(__dirname, "..");
const DAEMON_ENTRY = join(PROJECT_ROOT, "daemon", "index.mjs");
const SESSIONS_FILE = join(homedir(), ".claude", "office-claude", "sessions.json");
const LOG_FILE = join(homedir(), ".claude", "office-claude", "daemon.log");

// Don't launch the daemon with one of these as the initial cwd, even if a
// stale sessions.json says so. Matches the set in daemon/workspace.mjs.
// macOS-only: Windows has no equivalent OS-managed children of $HOME that
// could be silently confused with a real workspace folder.
const HOME = homedir();
const SYSTEM_HOME_CHILDREN = process.platform === "darwin"
  ? new Set([
      join(HOME, "Library"),
      join(HOME, "Movies"),
      join(HOME, "Music"),
      join(HOME, "Pictures"),
      join(HOME, "Public"),
    ])
  : new Set();

let tray = null;
let daemonProcess = null;
let logStream = null;
let daemonStatus = "starting"; // "starting" | "running" | "crashed" | "stopped"
let currentWorkspace = null;
let restartAttempts = 0;
const MAX_RESTART = 3;

// --------------------------------------------------------------------------
// Daemon lifecycle
// --------------------------------------------------------------------------
async function findInitialWorkspace() {
  try {
    const state = JSON.parse(await readFile(SESSIONS_FILE, "utf8"));
    const folders = Object.entries(state.folders || {})
      .filter(([cwd]) => !SYSTEM_HOME_CHILDREN.has(cwd));
    if (folders.length === 0) return null;
    folders.sort(([, a], [, b]) =>
      (b.last_used || "").localeCompare(a.last_used || "")
    );
    return folders[0][0];
  } catch {
    return null;
  }
}

function openLogStream() {
  fs.mkdirSync(dirname(LOG_FILE), { recursive: true });
  logStream = fs.createWriteStream(LOG_FILE, { flags: "a" });
  logStream.write(`\n=== ${new Date().toISOString()} app start ===\n`);
}

async function startDaemon() {
  const workspace = await findInitialWorkspace();
  currentWorkspace = workspace;
  daemonStatus = "starting";
  updateTray();

  const args = [DAEMON_ENTRY];
  if (workspace) args.push(workspace);

  daemonProcess = spawn(process.execPath, args, {
    cwd: PROJECT_ROOT,
    // fd 3 is an IPC channel for control messages (e.g. the native folder
    // picker). The daemon uses process.send / process.on('message') over this
    // channel; stdout/stderr remain plain log streams.
    stdio: ["ignore", "pipe", "pipe", "ipc"],
    env: { ...process.env, ELECTRON_RUN_AS_NODE: "1" },
  });

  daemonProcess.stdout.on("data", (chunk) => {
    logStream?.write(chunk);
    const text = chunk.toString();
    if (text.includes("Starting agent loop") || text.includes("Starting session for")) {
      daemonStatus = "running";
      restartAttempts = 0;
      updateTray();
    }
    const m = /Starting session for (.+?) \(/.exec(text);
    if (m) {
      currentWorkspace = m[1];
      updateTray();
    }
  });

  daemonProcess.stderr.on("data", (chunk) => {
    logStream?.write(chunk);
  });

  daemonProcess.on("message", handleDaemonMessage);

  daemonProcess.on("exit", (code, signal) => {
    const wasIntentional = signal === "SIGTERM" || daemonStatus === "stopped";
    daemonProcess = null;
    if (wasIntentional) {
      daemonStatus = "stopped";
    } else {
      daemonStatus = "crashed";
      logStream?.write(`\n[app] daemon exited unexpectedly (code=${code}, signal=${signal})\n`);
      if (restartAttempts < MAX_RESTART) {
        restartAttempts++;
        logStream?.write(`[app] auto-restarting (attempt ${restartAttempts}/${MAX_RESTART})\n`);
        setTimeout(() => startDaemon(), 1000);
      } else {
        logStream?.write(`[app] giving up after ${MAX_RESTART} restart attempts\n`);
      }
    }
    updateTray();
  });
}

function stopDaemon() {
  if (daemonProcess) {
    daemonStatus = "stopped";
    daemonProcess.kill("SIGTERM");
  }
}

function restartDaemon() {
  if (daemonProcess) {
    const proc = daemonProcess;
    daemonStatus = "starting";
    updateTray();
    proc.once("exit", () => {
      restartAttempts = 0;
      startDaemon();
    });
    proc.kill("SIGTERM");
  } else {
    restartAttempts = 0;
    startDaemon();
  }
}

// --------------------------------------------------------------------------
// IPC: handle control requests from the daemon. Currently just the native
// folder/file picker — the taskpane forwards picker requests over WebSocket
// to the daemon, which proxies them up here so we can show a real macOS
// open panel (NSOpenPanel, with sidebar shortcuts, Google Drive, iCloud,
// recent items — everything a synthetic in-page modal can't do).
// --------------------------------------------------------------------------
async function handleDaemonMessage(msg) {
  if (!msg || msg.type !== "pick_path") return;
  const reply = (payload) => {
    try { daemonProcess?.send({ type: "pick_path_result", id: msg.id, ...payload }); }
    catch (err) { logStream?.write(`[app] failed to reply to pick_path: ${err.message}\n`); }
  };
  try {
    const properties = msg.include_files
      ? ["openFile", "openDirectory", "createDirectory"]
      : ["openDirectory", "createDirectory"];
    // On macOS we're tray-only (LSUIElement), so the open panel surfaces
    // behind whatever is frontmost (Word/Excel) by default. Force-focus the
    // app first; with the dock hidden the user only sees the panel come
    // forward, not the app itself. On Windows, force-focus behaves
    // differently (taskbar flashing) and isn't typically needed because
    // tray apps aren't backgrounded the same way — skip it.
    if (process.platform === "darwin") app.focus({ steal: true });
    const result = await dialog.showOpenDialog({
      title: msg.title || (msg.include_files ? "Choose a folder or file" : "Choose a folder"),
      buttonLabel: msg.button_label || "Use this",
      defaultPath: msg.default_path || undefined,
      properties,
    });
    if (result.canceled || result.filePaths.length === 0) {
      reply({ ok: true, canceled: true });
      return;
    }
    const picked = result.filePaths[0];
    const s = await stat(picked);
    reply({ ok: true, path: picked, kind: s.isDirectory() ? "directory" : "file" });
  } catch (err) {
    reply({ ok: false, error: err.message });
  }
}

// --------------------------------------------------------------------------
// Tray UI
// --------------------------------------------------------------------------
function statusLabel() {
  switch (daemonStatus) {
    case "starting": return "● Starting…";
    case "running":  return "● Ready";
    case "crashed":  return restartAttempts >= MAX_RESTART
                       ? "● Crashed (won't restart)"
                       : `● Crashed (restart ${restartAttempts}/${MAX_RESTART})`;
    case "stopped":  return "● Stopped";
    default:         return "● ?";
  }
}

function workspaceLabel() {
  if (!currentWorkspace) return "Workspace: (none)";
  const name = currentWorkspace.split(/[\\/]/).filter(Boolean).pop();
  return `Workspace: ${name}`;
}

function buildMenu() {
  return Menu.buildFromTemplate([
    { label: statusLabel(), enabled: false },
    { label: workspaceLabel(), enabled: false },
    { type: "separator" },
    { label: "Open logs", click: () => shell.openPath(LOG_FILE) },
    {
      label: daemonProcess ? "Restart daemon" : "Start daemon",
      click: restartDaemon,
    },
    ...(daemonProcess ? [{ label: "Stop daemon", click: stopDaemon }] : []),
    {
      label: "Show in Finder",
      submenu: [
        { label: "Project folder", click: () => shell.openPath(PROJECT_ROOT) },
        ...(currentWorkspace
          ? [{ label: "Current workspace", click: () => shell.openPath(currentWorkspace) }]
          : []),
        { label: "Log file", click: () => shell.showItemInFolder(LOG_FILE) },
      ],
    },
    { type: "separator" },
    { label: "Open Microsoft Word", click: () => shell.openExternal("ms-word:") },
    { label: "Open Microsoft Excel", click: () => shell.openExternal("ms-excel:") },
    { type: "separator" },
    {
      label: addinInstalled ? "Reinstall add-in in Word + Excel" : "Install add-in in Word + Excel…",
      click: () => runInstall({ interactive: true }),
    },
    ...(addinInstalled
      ? [{ label: "Uninstall add-in", click: () => runUninstall({ interactive: true }) }]
      : []),
    { type: "separator" },
    { label: "Quit Office Claude", click: () => app.quit() },
  ]);
}

function updateTray() {
  if (!tray) return;
  tray.setToolTip(`Office Claude — ${statusLabel().replace(/^●\s*/, "")}`);
  tray.setContextMenu(buildMenu());
}

// --------------------------------------------------------------------------
// Add-in sideload — install / uninstall the Word + Excel manifests so the
// user never has to manually drop files into wef/ or touch Trust Center.
// --------------------------------------------------------------------------
let addinInstalled = false;

async function refreshAddinInstalled() {
  try { addinInstalled = await isAddinInstalled(); }
  catch { addinInstalled = false; }
  updateTray();
}

async function runInstall({ interactive }) {
  try {
    const result = await installAddin();
    addinInstalled = true;
    updateTray();
    if (interactive) {
      dialog.showMessageBox({
        type: "info",
        title: "Office Claude installed",
        message: "The add-in is now registered with Word and Excel.",
        detail:
          "Quit and reopen Word / Excel (if they're already running), then look for " +
          "Office Claude under Insert → Office Add-ins → Shared Folder.\n\n" +
          (process.platform === "win32"
            ? `Trusted catalog registered at:\n${result.catalog}`
            : `Manifests copied to each app's wef/ folder.`),
        buttons: ["OK"],
      }).catch(() => {});
    }
    return result;
  } catch (err) {
    logStream?.write(`[app] install failed: ${err.message}\n`);
    if (interactive) {
      dialog.showMessageBox({
        type: "error",
        title: "Couldn't install the add-in",
        message: err.message,
        detail:
          "You can sideload manually instead — see the README's Sideload section. " +
          "The shortest path: in Word/Excel, Insert → My Add-ins → Upload My Add-in → " +
          `pick the appropriate file from ${join(PROJECT_ROOT, "manifests")}.`,
        buttons: ["OK"],
      }).catch(() => {});
    }
  }
}

async function runUninstall({ interactive }) {
  try {
    await uninstallAddin();
    addinInstalled = false;
    updateTray();
    if (interactive) {
      dialog.showMessageBox({
        type: "info",
        title: "Add-in uninstalled",
        message: "Office Claude is no longer registered with Word or Excel.",
        detail: "The daemon is still running. Quit Office Claude from the tray menu to stop it entirely.",
        buttons: ["OK"],
      }).catch(() => {});
    }
  } catch (err) {
    logStream?.write(`[app] uninstall failed: ${err.message}\n`);
    if (interactive) {
      dialog.showMessageBox({
        type: "error",
        title: "Couldn't uninstall the add-in",
        message: err.message,
        buttons: ["OK"],
      }).catch(() => {});
    }
  }
}

// First-run prompt — invoked once at app start if the add-in isn't already
// registered. Skipped silently on unsupported platforms (Linux, etc.).
async function offerFirstRunInstall() {
  if (process.platform !== "darwin" && process.platform !== "win32") return;
  if (addinInstalled) return;
  const { response } = await dialog.showMessageBox({
    type: "question",
    title: "Install Office Claude in Word + Excel?",
    message: "Office Claude can install itself in Word and Excel automatically — no manifest copying or registry editing needed.",
    detail:
      "Click Install to register the add-in now. You can install later from the tray menu " +
      "if you'd prefer. After installing, open Word/Excel and find Office Claude under " +
      "Insert → Office Add-ins → Shared Folder.",
    buttons: ["Install", "Not now"],
    defaultId: 0,
    cancelId: 1,
  }).catch(() => ({ response: 1 }));
  if (response === 0) await runInstall({ interactive: true });
}

// --------------------------------------------------------------------------
// App lifecycle
// --------------------------------------------------------------------------
app.whenReady().then(async () => {
  // Background-only app — no dock icon on macOS.
  if (process.platform === "darwin") app.dock?.hide();

  openLogStream();

  // Tray icon. macOS expects a template (monochrome with transparency, auto-
  // recolored per theme); Windows expects a full-color .ico (or PNG fallback).
  // Convention: tray-icon.png is the macOS template; tray-icon-win.ico is the
  // Windows full-color version. Falls back to the PNG if .ico is missing.
  const macIconPath = join(__dirname, "tray-icon.png");
  const winIconPath = join(__dirname, "tray-icon-win.ico");
  const iconPath = process.platform === "win32" && fs.existsSync(winIconPath)
    ? winIconPath
    : macIconPath;
  const icon = nativeImage.createFromPath(iconPath);
  if (process.platform === "darwin") icon.setTemplateImage(true);
  tray = new Tray(icon);
  tray.setToolTip("Office Claude");

  // Check whether the add-in is already registered and refresh the tray
  // menu, then offer to install on first run.
  await refreshAddinInstalled();
  updateTray();
  startDaemon();
  offerFirstRunInstall().catch(err => logStream?.write(`[app] first-run prompt failed: ${err.message}\n`));
});

// Don't quit on "all windows closed" — we have no windows; tray is the UI.
app.on("window-all-closed", (e) => {
  if (e && e.preventDefault) e.preventDefault();
});

app.on("before-quit", () => {
  stopDaemon();
  logStream?.write(`=== ${new Date().toISOString()} app quit ===\n`);
  logStream?.end();
});
