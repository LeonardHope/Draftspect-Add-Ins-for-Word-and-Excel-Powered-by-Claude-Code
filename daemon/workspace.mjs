// Workspace-folder detection and marker management.
//
// "Workspace folder" = the folder Claude treats as cwd. The agent reads its
// CLAUDE.md, considers context files added to it, and uses it as the base
// for filesystem operations.
//
// Detection model (deliberately dead simple): the workspace is the folder
// that directly contains the open Word/Excel file. No walking up the tree,
// no marker search, no folder-name heuristics. For a document-centric tool
// the document *is* the subject, so its own folder is the natural cwd —
// and "no search" means no surprises (a stray ~/Desktop/.claude can't
// hijack the pick) and a result the user can always predict. The only
// thing that overrides it is the user explicitly choosing a folder in the
// picker; ensureWorkspaceMarker then drops a CLAUDE.md there.
//
// The single guard: never auto-select $HOME itself or an OS-managed $HOME
// child (~/Library, …). A doc loose in ~ shouldn't make all of home the
// workspace; the user can pick a real folder instead.

import { writeFile, stat } from "node:fs/promises";
import { readFileSync } from "node:fs";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { homedir } from "node:os";
import { join, dirname, basename, resolve as resolvePath } from "node:path";
import { fileURLToPath } from "node:url";
import { isSystemHomeChild } from "./system-paths.mjs";

const MARKERS = ["CLAUDE.md", ".claude"];
const HOME = homedir();

// Detect WSL once at module load. When the daemon runs under WSL, the
// taskpane (in Word on Windows) sends `C:\Users\...\file.docx`-style
// paths that mean nothing to a Linux process — they must be translated
// to `/mnt/c/Users/.../file.docx` before any stat/dirname can work.
const IS_WSL = (() => {
  if (process.platform !== "linux") return false;
  try {
    return /microsoft|wsl/i.test(readFileSync("/proc/version", "utf8"));
  } catch {
    return false;
  }
})();
const execFileP = promisify(execFile);

async function winPathToWsl(p) {
  try {
    const { stdout } = await execFileP("wslpath", ["-u", p]);
    return stdout.trim() || null;
  } catch {
    return null;
  }
}

// Office.js gives us paths in a few shapes depending on platform and where
// the doc came from:
//   - macOS local:   "/Users/x/Docs/spec.docx"
//   - macOS file://: "file:///Users/x/Docs/spec.docx"
//   - Windows local: "C:\\Users\\x\\Docs\\spec.docx"
//   - Windows URL:   "file:///C:/Users/x/Docs/spec.docx"
//   - SharePoint:    "https://contoso.sharepoint.com/…/spec.docx"
// Hand the URL forms to fileURLToPath so the Windows drive letter survives
// (a manual `replace(/^file:\/\//)` turns "file:///C:/x" into "/C:/x").
// Treat non-file:// URLs (SharePoint, OneDrive cloud) as unaddressable.
async function normalizeDocPath(docPath) {
  if (!docPath) return null;
  if (/^file:\/\//i.test(docPath)) {
    try {
      return resolvePath(fileURLToPath(docPath));
    } catch {
      return null;
    }
  }
  // A Windows drive path (C:\… or C:/…) or a UNC share (\\host\share)
  // is a real filesystem path, NOT a URL — and must be recognized BEFORE
  // the scheme check below, because a bare "C:" matches the URL-scheme
  // regex and would be wrongly discarded.
  //
  // Under WSL, hand these to wslpath -u so a Windows-side Word path
  // ("C:\\Users\\me\\Docs\\foo.docx") becomes a Linux-readable
  // /mnt/c/... path the daemon can stat. Without that step,
  // resolvePath() interprets the Windows path as a relative one on
  // Linux and the daemon falls back to its launch dir.
  if (/^[a-zA-Z]:[\\/]/.test(docPath) || /^\\\\[^\\]/.test(docPath)) {
    if (IS_WSL) {
      const wslPath = await winPathToWsl(docPath);
      return wslPath ? resolvePath(wslPath) : null;
    }
    return resolvePath(docPath);
  }
  if (/^[a-z][a-z0-9+.-]*:/i.test(docPath)) {
    // Some non-file URL (https://…sharepoint…) — no filesystem path.
    return null;
  }
  return resolvePath(decodeURIComponent(docPath));
}

async function docDirOf(p) {
  try {
    const s = await stat(p);
    return s.isDirectory() ? p : dirname(p);
  } catch {
    return dirname(p);
  }
}

// The workspace for a document: the folder that directly contains it.
// Returns that folder, or null if the doc is unaddressable (cloud URL) or
// its folder isn't a sane cwd ($HOME itself / an OS-managed $HOME child /
// the filesystem root).
export async function resolveWorkspaceRoot(docPath) {
  const p = await normalizeDocPath(docPath);
  if (!p) return null;
  const docDir = await docDirOf(p);
  if (docDir === HOME || isSystemHomeChild(docDir) || dirname(docDir) === docDir) {
    return null;
  }
  return docDir;
}

// Same answer, in the shape the taskpane expects. confidence is always
// "doc" (deterministic — the doc's own folder), so the pane can auto-apply
// it without a confirm banner. null when there's nothing sane to suggest.
export async function suggestWorkspaceRoot(docPath) {
  const cwd = await resolveWorkspaceRoot(docPath);
  return cwd ? { cwd, confidence: "doc" } : null;
}

// Drop an empty CLAUDE.md in the given folder if neither marker exists.
// Called on explicit user pick so the folder is set up for next time
// (context files, workspace instructions).
export async function ensureWorkspaceMarker(cwd) {
  if (!cwd) return false;
  for (const m of MARKERS) {
    try {
      await stat(join(cwd, m));
      return false;
    } catch {}
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
