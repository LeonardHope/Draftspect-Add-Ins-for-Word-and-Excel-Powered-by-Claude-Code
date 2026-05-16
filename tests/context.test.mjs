// Unit tests for daemon/context.mjs — the per-workspace context-files block
// embedded in CLAUDE.md. Tests cover round-trips, validation of missing
// paths, and parser tolerance for hand-edited blocks.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, writeFile, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { getContextEntries, setContextEntries } from "../daemon/context.mjs";

async function makeTmpWs() {
  const dir = await mkdtemp(join(tmpdir(), "cc-office-ctx-test-"));
  return dir;
}

test("getContextEntries returns [] when CLAUDE.md is missing", async () => {
  const ws = await makeTmpWs();
  try {
    const got = await getContextEntries(ws);
    assert.deepEqual(got, []);
  } finally {
    await rm(ws, { recursive: true, force: true });
  }
});

test("getContextEntries returns [] when CLAUDE.md exists but has no block", async () => {
  const ws = await makeTmpWs();
  try {
    await writeFile(join(ws, "CLAUDE.md"), "# Workspace\n\nNothing here yet.\n");
    const got = await getContextEntries(ws);
    assert.deepEqual(got, []);
  } finally {
    await rm(ws, { recursive: true, force: true });
  }
});

test("setContextEntries + getContextEntries round-trip", async () => {
  const ws = await makeTmpWs();
  try {
    // Create real files/folders so validation passes.
    const folderPath = join(ws, "notes");
    const filePath = join(ws, "ref.md");
    await mkdir(folderPath);
    await writeFile(filePath, "# ref\n");

    const { saved, errors } = await setContextEntries(ws, [
      { path: folderPath, description: "Project notes" },
      { path: filePath, description: "Reference doc" },
    ]);
    assert.equal(errors.length, 0);
    assert.equal(saved.length, 2);
    assert.equal(saved[0].kind, "folder");
    assert.equal(saved[1].kind, "file");

    const round = await getContextEntries(ws);
    assert.equal(round.length, 2);
    assert.equal(round[0].path, folderPath);
    assert.equal(round[0].kind, "folder");
    assert.equal(round[0].description, "Project notes");
    assert.equal(round[1].path, filePath);
    assert.equal(round[1].kind, "file");
  } finally {
    await rm(ws, { recursive: true, force: true });
  }
});

test("setContextEntries reports missing-path errors instead of saving them", async () => {
  const ws = await makeTmpWs();
  try {
    const realFolder = join(ws, "real");
    await mkdir(realFolder);
    const missing = join(ws, "does-not-exist");

    const { saved, errors } = await setContextEntries(ws, [
      { path: realFolder },
      { path: missing },
    ]);
    assert.equal(saved.length, 1);
    assert.equal(saved[0].path, realFolder);
    assert.equal(errors.length, 1);
    assert.equal(errors[0].path, missing);
    assert.match(errors[0].error, /Does not exist/);
  } finally {
    await rm(ws, { recursive: true, force: true });
  }
});

test("setContextEntries replaces an existing block in place (no duplication)", async () => {
  const ws = await makeTmpWs();
  try {
    const f1 = join(ws, "one");
    const f2 = join(ws, "two");
    await mkdir(f1);
    await mkdir(f2);

    await setContextEntries(ws, [{ path: f1 }]);
    await setContextEntries(ws, [{ path: f2 }]);

    const got = await getContextEntries(ws);
    assert.equal(got.length, 1);
    assert.equal(got[0].path, f2);

    const md = await readFile(join(ws, "CLAUDE.md"), "utf8");
    const beginCount = (md.match(/CONTEXT-FILES:BEGIN/g) || []).length;
    const endCount = (md.match(/CONTEXT-FILES:END/g) || []).length;
    assert.equal(beginCount, 1, "should have exactly one BEGIN marker");
    assert.equal(endCount, 1, "should have exactly one END marker");
  } finally {
    await rm(ws, { recursive: true, force: true });
  }
});

test("setContextEntries on empty CLAUDE.md preserves no prior content but wraps the block in a heading", async () => {
  const ws = await makeTmpWs();
  try {
    const folder = join(ws, "data");
    await mkdir(folder);
    await setContextEntries(ws, [{ path: folder, description: "data files" }]);
    const md = await readFile(join(ws, "CLAUDE.md"), "utf8");
    assert.match(md, /^# /m, "should have a top-level heading");
    assert.match(md, /CONTEXT-FILES:BEGIN/);
    assert.match(md, /\(folder\) `[^`]+` — data files/);
  } finally {
    await rm(ws, { recursive: true, force: true });
  }
});

test("getContextEntries ignores the marker string when it's inline in prose / a fenced block", async () => {
  const ws = await makeTmpWs();
  try {
    // The literal BEGIN marker appears inside a fenced code block (not as
    // its own delimiter line) AND there is a real block further down. The
    // line-anchored locator must match only the real one.
    const real =
      "<!-- CONTEXT-FILES:BEGIN -->\n\nP\n\n- (folder) `/tmp` — t\n\nQ\n\n<!-- CONTEXT-FILES:END -->";
    const md =
      "# Workspace\n\n" +
      "Docs may mention the marker `<!-- CONTEXT-FILES:BEGIN -->` inline.\n\n" +
      "```\n<!-- CONTEXT-FILES:BEGIN --> not a real delimiter here\n```\n\n" +
      real +
      "\n";
    await writeFile(join(ws, "CLAUDE.md"), md);
    const got = await getContextEntries(ws);
    assert.equal(got.length, 1);
    assert.equal(got[0].path, "/tmp");
    assert.equal(got[0].kind, "folder");
  } finally {
    await rm(ws, { recursive: true, force: true });
  }
});
