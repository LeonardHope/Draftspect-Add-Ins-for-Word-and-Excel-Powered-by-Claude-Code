// Workspace-folder detection and marker management.
//
// "Workspace folder" = the folder Claude treats as cwd. The agent reads its
// CLAUDE.md, considers context files added to it, and uses it as the base
// for filesystem operations.
//
// Detection cascade for a given Word/Excel doc path:
//   1. resolveWorkspaceRoot — walk up looking for a real marker (CLAUDE.md or
//      .claude). If found, that level is the workspace. Else null.
//   2. suggestWorkspaceRoot — if no marker, use a folder-name heuristic to
//      propose a candidate; the taskpane shows a confirm banner.
//
// On explicit user pick (via the native folder picker), ensureWorkspaceMarker
// drops an empty CLAUDE.md so the next open is silent.

import { readFile, writeFile, stat } from "node:fs/promises";
import { homedir } from "node:os";
import { join, dirname, basename, resolve as resolvePath } from "node:path";

const MATTER_MARKERS = ["CLAUDE.md", ".claude"];

// Subfolder names whose presence between the doc and a likely workspace root
// indicates the workspace is one level higher. Common convention: doc in a
// "Drafts" subfolder, with related materials as siblings.
const SUBFOLDER_HINTS = new Set([
  "drafts", "draft", "specs", "spec", "specification", "specifications",
  "documents", "docs", "working", "drafting", "files", "current",
]);

// Immediate children of $HOME that are OS-managed, never workspaces.
// macOS-specific; on Windows there's no equivalent stuck-state, so the set
// is empty.
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

function normalizeDocPath(docPath) {
  if (!docPath) return null;
  let p = docPath.startsWith("file://") ? docPath.replace(/^file:\/\//, "") : docPath;
  p = decodeURIComponent(p);
  return resolvePath(p);
}

async function docDirOf(p) {
  try {
    const s = await stat(p);
    return s.isDirectory() ? p : dirname(p);
  } catch {
    return dirname(p);
  }
}

// Walk up from a file path looking for a real workspace marker. Returns the
// workspace root if a marker is found, otherwise null. Deliberately does not
// guess — see suggestWorkspaceRoot for the heuristic fallback.
//
// Markers, in priority order:
//   1. CLAUDE.md or .claude — explicit user intent.
//   2. .git — the doc lives inside a git repo; the repo root is the workspace.
//      Previous behavior returned the *child* of the .git dir, which was
//      wrong: a doc at the repo root returned null, and a doc under repo/docs
//      returned repo/docs instead of repo.
export async function resolveWorkspaceRoot(docPath) {
  const p = normalizeDocPath(docPath);
  if (!p) return null;
  const docDir = await docDirOf(p);

  let dir = docDir;

  while (true) {
    if (dir === HOME) return null;
    for (const m of MATTER_MARKERS) {
      try { await stat(join(dir, m)); return dir; } catch {}
    }
    try {
      await stat(join(dir, ".git"));
      return dir;
    } catch {}
    const up = dirname(dir);
    if (up === dir) return null;
    dir = up;
  }
}

// Best-guess workspace root when there's no marker. Heuristic: if the doc's
// parent is a "drafts-like" folder, the workspace is the grandparent;
// otherwise the workspace is the doc's parent folder. Returns { cwd,
// confidence: "marker" | "heuristic" } or null.
export async function suggestWorkspaceRoot(docPath) {
  const real = await resolveWorkspaceRoot(docPath);
  if (real) return { cwd: real, confidence: "marker" };

  const p = normalizeDocPath(docPath);
  if (!p) return null;
  const docDir = await docDirOf(p);

  if (docDir === HOME || SYSTEM_HOME_CHILDREN.has(docDir) || dirname(docDir) === docDir) {
    return null;
  }

  const parentName = basename(docDir).toLowerCase();
  if (SUBFOLDER_HINTS.has(parentName)) {
    const up = dirname(docDir);
    if (up !== docDir && up !== HOME && !SYSTEM_HOME_CHILDREN.has(up)) {
      return { cwd: up, confidence: "heuristic" };
    }
  }
  return { cwd: docDir, confidence: "heuristic" };
}

// Drop an empty CLAUDE.md in the given folder if neither marker exists.
// Called on explicit user pick so next-open auto-detect is silent.
export async function ensureWorkspaceMarker(cwd) {
  if (!cwd) return false;
  for (const m of MATTER_MARKERS) {
    try { await stat(join(cwd, m)); return false; } catch {}
  }
  const seed =
    `# ${basename(cwd)}\n\n` +
    `_Workspace for the Office add-in. Add notes, glossary entries, ` +
    `style preferences, or instructions you want the assistant to follow ` +
    `here. The file is loaded automatically every time you open a document ` +
    `in this folder._\n`;
  await writeFile(join(cwd, "CLAUDE.md"), seed, { encoding: "utf8", flag: "wx" });
  return true;
}

export { SYSTEM_HOME_CHILDREN };
