// Patspect — Electron menu bar shell.
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

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = join(__dirname, "..");
const DAEMON_ENTRY = join(PROJECT_ROOT, "daemon", "index.mjs");
const SESSIONS_FILE = join(homedir(), ".claude", "word-addin", "sessions.json");
const LOG_FILE = join(homedir(), ".claude", "word-addin", "daemon.log");

// Don't launch the daemon with one of these as the initial cwd, even if a
// stale sessions.json says so. Matches the set in daemon/sessions.mjs.
const HOME = homedir();
const SYSTEM_HOME_CHILDREN = new Set([
  join(HOME, "Library"),
  join(HOME, "Movies"),
  join(HOME, "Music"),
  join(HOME, "Pictures"),
  join(HOME, "Public"),
]);

let tray = null;
let daemonProcess = null;
let logStream = null;
let daemonStatus = "starting"; // "starting" | "running" | "crashed" | "stopped"
let currentMatter = null;
let restartAttempts = 0;
const MAX_RESTART = 3;

// --------------------------------------------------------------------------
// Daemon lifecycle
// --------------------------------------------------------------------------
async function findInitialMatter() {
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
  const matter = await findInitialMatter();
  currentMatter = matter;
  daemonStatus = "starting";
  updateTray();

  const args = [DAEMON_ENTRY];
  if (matter) args.push(matter);

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
      currentMatter = m[1];
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
    // We're a tray-only (LSUIElement) app, so showOpenDialog will surface
    // the panel *behind* whatever is frontmost (Word) by default. Force the
    // app to the foreground first. The dock icon is hidden, so the user
    // doesn't see Patspect "appear" — only the panel comes forward.
    app.focus({ steal: true });
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

function matterLabel() {
  if (!currentMatter) return "Matter: (none)";
  const name = currentMatter.split(/[\\/]/).filter(Boolean).pop();
  return `Matter: ${name}`;
}

function buildMenu() {
  return Menu.buildFromTemplate([
    { label: statusLabel(), enabled: false },
    { label: matterLabel(), enabled: false },
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
        ...(currentMatter
          ? [{ label: "Current matter", click: () => shell.openPath(currentMatter) }]
          : []),
        { label: "Log file", click: () => shell.showItemInFolder(LOG_FILE) },
      ],
    },
    { type: "separator" },
    { label: "Open Microsoft Word", click: () => shell.openExternal("ms-word:") },
    { type: "separator" },
    { label: "Quit Patspect", click: () => app.quit() },
  ]);
}

function updateTray() {
  if (!tray) return;
  tray.setToolTip(`Patspect — ${statusLabel().replace(/^●\s*/, "")}`);
  tray.setContextMenu(buildMenu());
}

// --------------------------------------------------------------------------
// App lifecycle
// --------------------------------------------------------------------------
app.whenReady().then(() => {
  // Background-only app — no dock icon on macOS.
  if (process.platform === "darwin") app.dock?.hide();

  openLogStream();

  // Tray icon (template image — macOS recolors per theme).
  const iconPath = join(__dirname, "tray-icon.png");
  const icon = nativeImage.createFromPath(iconPath);
  icon.setTemplateImage(true);
  tray = new Tray(icon);
  tray.setToolTip("Patspect");

  updateTray();
  startDaemon();
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
