// Read/write the user's drafting setup at ~/.claude/word-addin/drafting-setup.md.
//
// This is *product-scoped* — only the Word add-in's daemon reads it, by
// appending its content to the system prompt at agent-loop start. It does NOT
// load into terminal `claude` sessions or any other CC use of the user's
// machine. That keeps the per-session token cost where it's useful (drafting)
// and avoids polluting the user's universal Claude Code config.
//
// Two distinct kinds of reference folder:
//   - "guidelines" — rules and conventions the agent must follow when drafting
//   - "samples"    — prior applications to use as style/structure references
//
// Each kind is stored in its own marker-bounded block so the user can edit
// this file manually if they want. Changes apply on the next agent session
// (daemon restart, or close-and-reopen the add-in).

import { readFile, writeFile, stat, mkdir } from "node:fs/promises";
import { homedir } from "node:os";
import { join, dirname, basename, resolve as resolvePath } from "node:path";

const MATTER_MARKERS = ["CLAUDE.md", ".claude"];

// Subfolder names whose presence between the doc and a likely matter root
// indicates the matter is one level higher. e.g. .../A088-0176US/Drafts/spec.docx
// → the matter is A088-0176US, not Drafts. Patspect convention plus a few
// common synonyms.
const DRAFTS_LIKE = new Set([
  "drafts", "draft", "specs", "spec", "specification", "specifications",
  "documents", "docs", "working", "drafting", "files",
]);

// Immediate children of $HOME that are macOS system folders, never matters.
const HOME = homedir();
const SYSTEM_HOME_CHILDREN = new Set([
  join(HOME, "Library"),
  join(HOME, "Movies"),
  join(HOME, "Music"),
  join(HOME, "Pictures"),
  join(HOME, "Public"),
]);

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

const SETUP_FILE = join(homedir(), ".claude", "word-addin", "drafting-setup.md");

const KINDS = {
  guidelines: {
    begin: "<!-- DRAFTING-GUIDELINES:BEGIN -->",
    end:   "<!-- DRAFTING-GUIDELINES:END -->",
    preamble:
      "The following folders contain drafting guidelines, style rules, and attorney preferences " +
      "I follow when drafting patent specifications. **Treat the content of these as rules to " +
      "obey** — at the start of any drafting task, read the relevant files in these folders and " +
      "honor their guidance throughout. Do not modify any files in them.",
    postamble:
      "If a guideline conflicts with a per-matter `CLAUDE.md` or an explicit user instruction, " +
      "the more specific source wins.",
  },
  samples: {
    begin: "<!-- SAMPLE-APPLICATIONS:BEGIN -->",
    end:   "<!-- SAMPLE-APPLICATIONS:END -->",
    preamble:
      "The following folders contain prior patent applications and example drafts. Use them as " +
      "**structural and stylistic references** when drafting — pattern-match against them for " +
      "tone, paragraph structure, claim style, and phrasing conventions. Do not copy substantive " +
      "content verbatim; the goal is matching style and structure, not the underlying invention.",
    postamble:
      "Read the relevant samples on demand using Read / Glob / Grep. For .docx files, the raw " +
      "Read returns OOXML — use `unzip -p <file> word/document.xml` plus a text extraction step " +
      "(grep/python) to get clean text.",
  },
};

function lineFor(ref) {
  const desc = (ref.description || "").trim();
  return desc ? `- \`${ref.path}\` — ${desc}` : `- \`${ref.path}\``;
}

function parseBlockBody(body) {
  const refs = [];
  for (const raw of body.split("\n")) {
    const line = raw.trim();
    const m = /^-\s+`([^`]+)`(?:\s*—\s*(.+))?$/.exec(line);
    if (m) refs.push({ path: m[1], description: (m[2] || "").trim() });
  }
  return refs;
}

async function readSetupFile() {
  try { return await readFile(SETUP_FILE, "utf8"); }
  catch (e) { if (e.code === "ENOENT") return ""; throw e; }
}

// Used by the daemon at startup to inject the full setup file (markers and
// all) into the agent's system prompt. The agent reads the bullet paths from
// it and consults the folders on demand.
export async function readDraftingSetupForPrompt() {
  const content = await readSetupFile();
  return content.trim();
}

function extractBlock(content, kind) {
  const cfg = KINDS[kind];
  const start = content.indexOf(cfg.begin);
  const end = content.indexOf(cfg.end);
  if (start === -1 || end === -1 || end < start) return [];
  return parseBlockBody(content.slice(start + cfg.begin.length, end));
}

function renderBlock(kind, refs) {
  const cfg = KINDS[kind];
  const lines = [cfg.begin, "", cfg.preamble, ""];
  if (refs.length === 0) {
    lines.push("_(none configured)_");
  } else {
    for (const r of refs) lines.push(lineFor(r));
  }
  lines.push("", cfg.postamble, "", cfg.end);
  return lines.join("\n");
}

function replaceBlock(content, kind, refs) {
  const cfg = KINDS[kind];
  const block = renderBlock(kind, refs);
  const start = content.indexOf(cfg.begin);
  const end = content.indexOf(cfg.end);
  if (start !== -1 && end !== -1 && end > start) {
    return content.slice(0, start) + block + content.slice(end + cfg.end.length);
  }
  const sep = content.length === 0
    ? ""
    : (content.endsWith("\n\n") ? "" : (content.endsWith("\n") ? "\n" : "\n\n"));
  return content + sep + block + "\n";
}

async function validate(refs) {
  const validated = [];
  const errors = [];
  for (const r of refs || []) {
    if (!r || typeof r.path !== "string" || !r.path.trim()) {
      errors.push({ path: r?.path ?? "", error: "Path is required" });
      continue;
    }
    const path = r.path.trim();
    try {
      const s = await stat(path);
      if (!s.isDirectory()) { errors.push({ path, error: "Not a directory" }); continue; }
    } catch (e) {
      errors.push({ path, error: e.code === "ENOENT" ? "Does not exist" : (e.message || "stat failed") });
      continue;
    }
    validated.push({ path, description: (r.description || "").trim() });
  }
  return { validated, errors };
}

export async function getDraftingSetup() {
  const content = await readSetupFile();
  return {
    guidelines: extractBlock(content, "guidelines"),
    samples: extractBlock(content, "samples"),
  };
}

export async function setDraftingSetup({ guidelines, samples }) {
  const allErrors = [];

  const g = await validate(guidelines);
  for (const e of g.errors) allErrors.push({ kind: "guidelines", ...e });

  const s = await validate(samples);
  for (const e of s.errors) allErrors.push({ kind: "samples", ...e });

  let content = await readSetupFile();
  content = replaceBlock(content, "guidelines", g.validated);
  content = replaceBlock(content, "samples", s.validated);

  await mkdir(dirname(SETUP_FILE), { recursive: true });
  await writeFile(SETUP_FILE, content, "utf8");

  return {
    saved: { guidelines: g.validated, samples: s.validated },
    errors: allErrors,
  };
}

// Walk up from a file path looking for a real matter marker (CLAUDE.md /
// .claude / .git). Returns the matter root if one is found, or null if not.
//
// We deliberately do NOT guess — no "immediate child of $HOME" or "doc's
// parent" fallback. Those produce wrong answers more often than right ones
// once iCloud / Google Drive / CloudStorage paths are in the mix (e.g.
// returning ~/Library for a doc deep inside a Google Drive mount). Callers
// that want a guess when there's no marker should call suggestMatterRoot.
export async function resolveMatterRoot(docPath) {
  const p = normalizeDocPath(docPath);
  if (!p) return null;
  const docDir = await docDirOf(p);

  let dir = docDir;
  let prevDir = null;

  while (true) {
    if (dir === HOME) return null;

    for (const m of MATTER_MARKERS) {
      try { await stat(join(dir, m)); return dir; } catch {}
    }

    // .git at this level → matter is one level below.
    try {
      await stat(join(dir, ".git"));
      return prevDir;
    } catch {}

    prevDir = dir;
    const up = dirname(dir);
    if (up === dir) return null; // fs root, no markers found
    dir = up;
  }
}

// Best-guess matter root when there's no real marker. Uses a folder-name
// heuristic: if the doc lives in a "drafts-like" subfolder (Drafts, Specs,
// Documents, …), the matter is the parent of that subfolder; otherwise the
// matter is just the doc's own folder. Returns { cwd, confidence } where
// confidence is "marker" (real marker found, guaranteed correct) or
// "heuristic" (a guess that the user should confirm).
//
// Returns null if no sensible guess can be made (doc in $HOME / system
// folder / fs root, or doc path missing).
export async function suggestMatterRoot(docPath) {
  // If a real marker exists upstream, that's the best answer.
  const real = await resolveMatterRoot(docPath);
  if (real) return { cwd: real, confidence: "marker" };

  const p = normalizeDocPath(docPath);
  if (!p) return null;
  const docDir = await docDirOf(p);

  // Refuse to suggest anything at $HOME, immediately under it (~/Library
  // etc.), or at the fs root.
  if (docDir === HOME || SYSTEM_HOME_CHILDREN.has(docDir) || dirname(docDir) === docDir) {
    return null;
  }

  // If the doc's parent looks like a "drafts" subfolder, the matter is the
  // grandparent. Otherwise it's the parent.
  const parentName = basename(docDir).toLowerCase();
  if (DRAFTS_LIKE.has(parentName)) {
    const up = dirname(docDir);
    if (up !== docDir && up !== HOME && !SYSTEM_HOME_CHILDREN.has(up)) {
      return { cwd: up, confidence: "heuristic" };
    }
  }
  return { cwd: docDir, confidence: "heuristic" };
}

// Drop an empty CLAUDE.md in the given folder if neither matter marker
// already exists. Called when the user explicitly picks a matter root via
// the picker, so the next open of any doc in this matter is silently
// auto-detected. Returns true if a file was created, false if a marker
// already existed.
export async function ensureMatterMarker(cwd) {
  if (!cwd) return false;
  for (const m of MATTER_MARKERS) {
    try { await stat(join(cwd, m)); return false; } catch {}
  }
  const seed =
    `# ${basename(cwd)}\n\n` +
    `_Matter context for the Patspect Word add-in. Add matter-specific ` +
    `instructions, glossary entries, or claim-style notes here._\n`;
  await writeFile(join(cwd, "CLAUDE.md"), seed, { encoding: "utf8", flag: "wx" });
  return true;
}
