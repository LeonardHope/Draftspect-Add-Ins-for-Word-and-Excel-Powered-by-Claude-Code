// Per-folder Agent SDK session bookkeeping. The SDK itself persists each
// session's transcript to ~/.claude/projects/<hash>/*.jsonl, where <hash> is
// derived from the cwd. To resume a session via `query({ options: { resume } })`
// we need the session_id string, which we record here per cwd along with a
// last-used timestamp.
//
// Stored at ~/.claude/office-addins/sessions.json:
//
//   {
//     "version": 1,
//     "folders": {
//       "/Users/leonard/.../DEMO-2026-001": {
//         "session_id": "uuid…",
//         "last_used": "2026-05-14T08:45:12.000Z",
//         "display_name": "DEMO-2026-001"
//       },
//       ...
//     }
//   }

import { readFile, writeFile, mkdir } from "node:fs/promises";
import { homedir } from "node:os";
import { join, dirname, basename } from "node:path";
import { isSystemHomeChild } from "./system-paths.mjs";

const FILE = join(homedir(), ".claude", "office-addins", "sessions.json");

// Don't persist or surface OS-managed $HOME children (e.g. ~/Library) as
// a recent workspace. Canonical set lives in ./system-paths.mjs — this
// was previously a drifted unconditional copy.
function isAllowedMatterPath(cwd) {
  return !isSystemHomeChild(cwd);
}

async function readState() {
  try {
    const raw = await readFile(FILE, "utf8");
    const parsed = JSON.parse(raw);
    if (!parsed || parsed.version !== 1 || typeof parsed.folders !== "object") {
      return { version: 1, folders: {} };
    }
    return parsed;
  } catch (e) {
    if (e.code === "ENOENT") return { version: 1, folders: {} };
    throw e;
  }
}

async function writeState(state) {
  await mkdir(dirname(FILE), { recursive: true });
  await writeFile(FILE, JSON.stringify(state, null, 2) + "\n", "utf8");
}

export async function getSessionForFolder(cwd) {
  const state = await readState();
  return state.folders[cwd] ?? null;
}

export async function saveSessionForFolder(cwd, sessionId) {
  const state = await readState();
  state.folders[cwd] = {
    session_id: sessionId,
    last_used: new Date().toISOString(),
    display_name: basename(cwd),
  };
  await writeState(state);
}

export async function touchFolder(cwd) {
  if (!isAllowedMatterPath(cwd)) return;
  // Update last_used without changing the session_id.
  const state = await readState();
  const existing = state.folders[cwd];
  if (existing) {
    existing.last_used = new Date().toISOString();
  } else {
    state.folders[cwd] = {
      session_id: null,
      last_used: new Date().toISOString(),
      display_name: basename(cwd),
    };
  }
  await writeState(state);
}

export async function getRecentFolders(limit = 10) {
  const state = await readState();
  return Object.entries(state.folders)
    .filter(([cwd]) => isAllowedMatterPath(cwd))
    .map(([cwd, info]) => ({ cwd, ...info }))
    .sort((a, b) => (b.last_used || "").localeCompare(a.last_used || ""))
    .slice(0, limit);
}

export async function forgetFolder(cwd) {
  const state = await readState();
  delete state.folders[cwd];
  await writeState(state);
}
