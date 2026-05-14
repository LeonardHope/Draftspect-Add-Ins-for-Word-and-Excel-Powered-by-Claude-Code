// Unit tests for daemon/workspace.mjs — workspace detection cascade.
//
// resolveWorkspaceRoot is the strict, marker-driven path. suggestWorkspaceRoot
// is the heuristic fallback used when no marker exists. ensureWorkspaceMarker
// drops a seed CLAUDE.md on explicit user pick. Tests use a fresh tmpdir per
// test so they don't touch each other or the user's real ~/.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import {
  resolveWorkspaceRoot,
  suggestWorkspaceRoot,
  ensureWorkspaceMarker,
} from "../daemon/workspace.mjs";

async function makeTmp() {
  return await mkdtemp(join(tmpdir(), "office-claude-ws-test-"));
}

test("resolveWorkspaceRoot returns null when no marker anywhere up the tree", async () => {
  const root = await makeTmp();
  try {
    const sub = join(root, "a", "b");
    await mkdir(sub, { recursive: true });
    await writeFile(join(sub, "doc.docx"), "");
    const got = await resolveWorkspaceRoot(join(sub, "doc.docx"));
    // No marker, no .git → null.
    assert.equal(got, null);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot stops at the CLAUDE.md marker", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    const sub = join(ws, "drafts");
    await mkdir(sub, { recursive: true });
    await writeFile(join(ws, "CLAUDE.md"), "# workspace\n");
    await writeFile(join(sub, "doc.docx"), "");
    const got = await resolveWorkspaceRoot(join(sub, "doc.docx"));
    assert.equal(got, ws);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot stops at the .claude marker (file or dir)", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    const sub = join(ws, "drafts");
    await mkdir(sub, { recursive: true });
    await mkdir(join(ws, ".claude")); // dir variant
    await writeFile(join(sub, "doc.docx"), "");
    const got = await resolveWorkspaceRoot(join(sub, "doc.docx"));
    assert.equal(got, ws);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot returns the git root when .git is found (not its child)", async () => {
  // Regression test for the prevDir bug: doc nested in a git repo used to
  // return the directory just below .git instead of the repo root.
  const root = await makeTmp();
  try {
    const repo = join(root, "repo");
    const docDir = join(repo, "docs");
    await mkdir(docDir, { recursive: true });
    await mkdir(join(repo, ".git"));
    await writeFile(join(docDir, "spec.docx"), "");
    const got = await resolveWorkspaceRoot(join(docDir, "spec.docx"));
    assert.equal(got, repo);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot returns the git root when the doc is at the repo root", async () => {
  // Regression test: doc directly in a git repo used to return null because
  // prevDir was still null on the first iteration.
  const root = await makeTmp();
  try {
    const repo = join(root, "repo");
    await mkdir(repo);
    await mkdir(join(repo, ".git"));
    await writeFile(join(repo, "spec.docx"), "");
    const got = await resolveWorkspaceRoot(join(repo, "spec.docx"));
    assert.equal(got, repo);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot prefers CLAUDE.md over .git when both present", async () => {
  // Markers have priority. If a user dropped a CLAUDE.md in a subfolder of a
  // git repo, that subfolder is the workspace, not the repo root.
  const root = await makeTmp();
  try {
    const repo = join(root, "repo");
    const sub = join(repo, "matter");
    await mkdir(sub, { recursive: true });
    await mkdir(join(repo, ".git"));
    await writeFile(join(sub, "CLAUDE.md"), "");
    await writeFile(join(sub, "spec.docx"), "");
    const got = await resolveWorkspaceRoot(join(sub, "spec.docx"));
    assert.equal(got, sub);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot handles file:// URLs (POSIX)", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    await mkdir(ws, { recursive: true });
    await writeFile(join(ws, "CLAUDE.md"), "");
    const docPath = join(ws, "doc.docx");
    await writeFile(docPath, "");
    const fileUrl = "file://" + docPath;
    const got = await resolveWorkspaceRoot(fileUrl);
    assert.equal(got, ws);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot returns null for non-file:// URLs (cloud docs)", async () => {
  // SharePoint / OneDrive URLs are not filesystem paths.
  const got = await resolveWorkspaceRoot("https://contoso.sharepoint.com/sites/x/spec.docx");
  assert.equal(got, null);
});

test("suggestWorkspaceRoot falls back to docDir when no marker", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    await mkdir(ws, { recursive: true });
    await writeFile(join(ws, "doc.docx"), "");
    const got = await suggestWorkspaceRoot(join(ws, "doc.docx"));
    assert.deepEqual(got, { cwd: ws, confidence: "heuristic" });
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("suggestWorkspaceRoot reports marker confidence when one exists", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    await mkdir(ws, { recursive: true });
    await writeFile(join(ws, "CLAUDE.md"), "");
    await writeFile(join(ws, "doc.docx"), "");
    const got = await suggestWorkspaceRoot(join(ws, "doc.docx"));
    assert.deepEqual(got, { cwd: ws, confidence: "marker" });
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("suggestWorkspaceRoot bumps up one level when doc lives in a drafts-like folder", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "matter");
    const draftsDir = join(ws, "Drafts");
    await mkdir(draftsDir, { recursive: true });
    await writeFile(join(draftsDir, "doc.docx"), "");
    const got = await suggestWorkspaceRoot(join(draftsDir, "doc.docx"));
    assert.deepEqual(got, { cwd: ws, confidence: "heuristic" });
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("ensureWorkspaceMarker drops a CLAUDE.md when none exists", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    await mkdir(ws);
    const created = await ensureWorkspaceMarker(ws);
    assert.equal(created, true);
    const second = await ensureWorkspaceMarker(ws);
    // Second call is a no-op since the marker now exists.
    assert.equal(second, false);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("ensureWorkspaceMarker does nothing when .claude already exists", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    await mkdir(ws);
    await mkdir(join(ws, ".claude"));
    const created = await ensureWorkspaceMarker(ws);
    assert.equal(created, false);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});
