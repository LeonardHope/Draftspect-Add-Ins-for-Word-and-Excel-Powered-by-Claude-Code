// Per-workspace context list: folders and/or files the user wants the
// assistant to consider as background material for whatever's open.
//
// Stored as a marker-bounded block inside the workspace's CLAUDE.md.
// Claude Code auto-loads each workspace's CLAUDE.md into the agent's system
// prompt at session init, so anything we write here flows into the agent's
// context automatically — no extra plumbing in the daemon. Switching
// workspaces switches which CLAUDE.md is loaded.

import { readFile, writeFile, stat } from "node:fs/promises";
import { join, resolve as resolvePath } from "node:path";

const BEGIN = "<!-- CONTEXT-FILES:BEGIN -->";
const END = "<!-- CONTEXT-FILES:END -->";
const PREAMBLE =
  "The following folders and files are background context the user has " +
  "added for this workspace. Read them on demand — when the user's request " +
  "involves a topic, term, or document name that one of these entries " +
  "looks relevant to, use `Read` / `Glob` / `Grep` to consult them before " +
  "answering.";
const POSTAMBLE =
  "Each entry below is annotated with its kind (folder or file). Folder " +
  "entries should be globbed when relevant. File entries are individual " +
  "documents to read directly. Do not modify any of these files unless " +
  "the user explicitly asks you to.";

function lineFor(entry) {
  const k = entry.kind === "file" ? "file" : "folder";
  const desc = (entry.description || "").trim();
  return desc
    ? `- (${k}) \`${entry.path}\` — ${desc}`
    : `- (${k}) \`${entry.path}\``;
}

function parseBlock(blockBody) {
  const out = [];
  for (const raw of blockBody.split("\n")) {
    const line = raw.trim();
    const m = /^-\s+\((file|folder)\)\s+`([^`]+)`(?:\s*—\s*(.+))?$/.exec(line);
    if (m) out.push({ kind: m[1], path: m[2], description: (m[3] || "").trim() });
  }
  return out;
}

function claudeMdPath(cwd) {
  return join(cwd, "CLAUDE.md");
}

async function readClaudeMd(cwd) {
  try { return await readFile(claudeMdPath(cwd), "utf8"); }
  catch (e) { if (e.code === "ENOENT") return ""; throw e; }
}

async function validate(entries) {
  const validated = [];
  const errors = [];
  for (const e of entries || []) {
    if (!e || typeof e.path !== "string" || !e.path.trim()) {
      errors.push({ path: e?.path ?? "", error: "Path is required" });
      continue;
    }
    const p = resolvePath(e.path.trim());
    try {
      const s = await stat(p);
      const detectedKind = s.isDirectory() ? "folder" : (s.isFile() ? "file" : null);
      if (!detectedKind) {
        errors.push({ path: p, error: "Not a regular file or directory" });
        continue;
      }
      validated.push({
        kind: detectedKind,
        path: p,
        description: (e.description || "").trim(),
      });
    } catch (err) {
      errors.push({
        path: p,
        error: err.code === "ENOENT" ? "Does not exist" : (err.message || "stat failed"),
      });
    }
  }
  return { validated, errors };
}

export async function getContextEntries(cwd) {
  if (!cwd) return [];
  const content = await readClaudeMd(cwd);
  const start = content.indexOf(BEGIN);
  const end = content.indexOf(END);
  if (start === -1 || end === -1 || end < start) return [];
  return parseBlock(content.slice(start + BEGIN.length, end));
}

export async function setContextEntries(cwd, entries) {
  if (!cwd) throw new Error("setContextEntries requires cwd");
  await stat(cwd); // confirm folder exists

  const { validated, errors } = await validate(entries);

  const lines = [BEGIN, "", PREAMBLE, ""];
  if (validated.length === 0) {
    lines.push("_(no context entries configured)_");
  } else {
    for (const e of validated) lines.push(lineFor(e));
  }
  lines.push("", POSTAMBLE, "", END);
  const newBlock = lines.join("\n");

  let content = await readClaudeMd(cwd);
  const start = content.indexOf(BEGIN);
  const end = content.indexOf(END);
  if (start !== -1 && end !== -1 && end > start) {
    content = content.slice(0, start) + newBlock + content.slice(end + END.length);
  } else if (content.length === 0) {
    content = `# Workspace context\n\n${newBlock}\n`;
  } else {
    const sep = content.endsWith("\n\n") ? "" : (content.endsWith("\n") ? "\n" : "\n\n");
    content = content + sep + newBlock + "\n";
  }

  await writeFile(claudeMdPath(cwd), content, "utf8");
  return { saved: validated, errors };
}
