// Per-matter disclosure-materials chooser: read/write a marker-bounded block
// in the matter's own CLAUDE.md.
//
// CC auto-loads each matter folder's CLAUDE.md into the agent's system
// prompt at session init, so anything we write here flows into the agent's
// context automatically — no extra plumbing in the daemon. Switching matters
// switches which CLAUDE.md is loaded.
//
// Entries are either folders or specific files. Each is written as a bullet
// in the marker block with its kind annotated.

import { readFile, writeFile, stat } from "node:fs/promises";
import { join, resolve as resolvePath } from "node:path";

const BEGIN = "<!-- DISCLOSURE-MATERIALS:BEGIN -->";
const END = "<!-- DISCLOSURE-MATERIALS:END -->";
const PREAMBLE =
  "The following are the disclosure materials for this specific matter — " +
  "invention disclosures, inventor interviews, technical notes, prior-art " +
  "notes, and any related primary source material. **Treat these as the " +
  "primary content sources when drafting** any section of the specification. " +
  "Read them on demand; cite the source file in a Word comment on each " +
  "drafted paragraph that derives from them.";
const POSTAMBLE =
  "Each entry below is annotated with its kind (folder or file). Folder " +
  "entries should be globbed when relevant. File entries are individual " +
  "documents to read directly. Do not modify any of these files.";

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
  try {
    return await readFile(claudeMdPath(cwd), "utf8");
  } catch (e) {
    if (e.code === "ENOENT") return "";
    throw e;
  }
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

export async function getMatterDisclosure(cwd) {
  if (!cwd) return [];
  const content = await readClaudeMd(cwd);
  const start = content.indexOf(BEGIN);
  const end = content.indexOf(END);
  if (start === -1 || end === -1 || end < start) return [];
  return parseBlock(content.slice(start + BEGIN.length, end));
}

export async function setMatterDisclosure(cwd, entries) {
  if (!cwd) throw new Error("setMatterDisclosure requires cwd");

  // Confirm the matter folder exists
  await stat(cwd);

  const { validated, errors } = await validate(entries);

  // Build the block
  const lines = [BEGIN, "", PREAMBLE, ""];
  if (validated.length === 0) {
    lines.push("_(no disclosure materials configured)_");
  } else {
    for (const e of validated) lines.push(lineFor(e));
  }
  lines.push("", POSTAMBLE, "", END);
  const newBlock = lines.join("\n");

  // Read existing CLAUDE.md (or empty if absent), splice in / replace the block.
  let content = await readClaudeMd(cwd);
  const start = content.indexOf(BEGIN);
  const end = content.indexOf(END);
  if (start !== -1 && end !== -1 && end > start) {
    content = content.slice(0, start) + newBlock + content.slice(end + END.length);
  } else {
    if (content.length === 0) {
      // Brand-new CLAUDE.md — start with a small header so a manual reader
      // immediately sees what it is.
      content = `# Matter context\n\n${newBlock}\n`;
    } else {
      const sep = content.endsWith("\n\n") ? "" : (content.endsWith("\n") ? "\n" : "\n\n");
      content = content + sep + newBlock + "\n";
    }
  }

  await writeFile(claudeMdPath(cwd), content, "utf8");
  return { saved: validated, errors };
}
