// Unit tests for daemon/workspace.mjs.
//
// New model: the workspace is simply the folder that directly contains the
// open document — no walking up the tree, no marker search, no folder-name
// heuristic. The single guard is "never $HOME / an OS-managed $HOME child /
// the filesystem root". ensureWorkspaceMarker still drops a seed CLAUDE.md
// on explicit user pick. Tests use a fresh tmpdir per test.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdir, mkdtemp, writeFile, rm, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import {
  resolveWorkspaceRoot,
  suggestWorkspaceRoot,
  ensureWorkspaceMarker,
} from "../daemon/workspace.mjs";

async function makeTmp() {
  return await mkdtemp(join(tmpdir(), "cc-office-ws-test-"));
}

test("resolveWorkspaceRoot returns the document's own folder", async () => {
  const root = await makeTmp();
  try {
    const dir = join(root, "a", "b");
    await mkdir(dir, { recursive: true });
    await writeFile(join(dir, "doc.docx"), "");
    assert.equal(await resolveWorkspaceRoot(join(dir, "doc.docx")), dir);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot does NOT walk up to ancestor markers (footgun fix)", async () => {
  // The ~/Desktop/.claude footgun: a stray marker in an ancestor used to
  // hijack the pick. Now the doc's own folder always wins.
  const root = await makeTmp();
  try {
    const sub = join(root, "Test Folder");
    await mkdir(sub, { recursive: true });
    await writeFile(join(root, "CLAUDE.md"), "# stray\n");
    await mkdir(join(root, ".claude"));
    await mkdir(join(root, ".git"));
    await writeFile(join(sub, "doc.docx"), "");
    // Returns the doc's folder, not `root` (where the markers are).
    assert.equal(await resolveWorkspaceRoot(join(sub, "doc.docx")), sub);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot returns the directory itself when handed a dir path", async () => {
  const root = await makeTmp();
  try {
    const dir = join(root, "ws");
    await mkdir(dir, { recursive: true });
    assert.equal(await resolveWorkspaceRoot(dir), dir);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot handles file:// URLs (POSIX)", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    await mkdir(ws, { recursive: true });
    const docPath = join(ws, "doc.docx");
    await writeFile(docPath, "");
    assert.equal(await resolveWorkspaceRoot("file://" + docPath), ws);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("resolveWorkspaceRoot returns null for non-file:// URLs (cloud docs)", async () => {
  const got = await resolveWorkspaceRoot("https://contoso.sharepoint.com/sites/x/spec.docx");
  assert.equal(got, null);
});

test("resolveWorkspaceRoot does NOT treat a Windows drive path as a URL (regression)", async () => {
  // Regression: a bare "C:" matched the URL-scheme regex, so every
  // Windows document path was discarded and the workspace fell back to
  // the daemon's launch dir. A drive path must resolve to a folder, not
  // null. (On non-Windows CI, path.resolve keeps the backslashes as
  // filename chars — that's fine; the point is it's NOT null/URL. On
  // WSL, wslpath translates the path to /mnt/c/... — also not null.)
  assert.notEqual(await resolveWorkspaceRoot("C:\\Users\\me\\Docs\\spec.docx"), null);

  // A UNC share-path verifies the same regex regression for the
  // backslash-backslash form. On WSL the daemon can't reach an arbitrary
  // SMB share without a mount, so wslpath fails and resolveWorkspaceRoot
  // legitimately returns null — that's the correct "we can't address
  // this" answer for a Linux-side daemon, not the URL-misclassification
  // bug this test guards against. Skip on WSL.
  const isWsl =
    process.platform === "linux" &&
    /microsoft|wsl/i.test(
      await readFile("/proc/version", "utf8").catch(() => ""),
    );
  if (!isWsl) {
    assert.notEqual(await resolveWorkspaceRoot("\\\\fileserver\\share\\spec.docx"), null);
  }
});

test("resolveWorkspaceRoot returns null when the doc sits in an OS-managed $HOME child", async () => {
  if (process.platform !== "darwin") return; // deny-list is macOS-specific
  const fakeHome = await mkdtemp(join(tmpdir(), "cc-office-ws-home-"));
  const prevHome = process.env.HOME;
  const prevUserProfile = process.env.USERPROFILE;
  process.env.HOME = fakeHome;
  process.env.USERPROFILE = fakeHome;
  try {
    const lib = join(fakeHome, "Library");
    await mkdir(lib, { recursive: true });
    await writeFile(join(lib, "doc.docx"), "");
    // Re-import with the faked HOME (workspace.mjs reads homedir() at load).
    const mod = await import("../daemon/workspace.mjs?home=" + Date.now() + Math.random());
    assert.equal(await mod.resolveWorkspaceRoot(join(lib, "doc.docx")), null);
  } finally {
    if (prevHome === undefined) delete process.env.HOME;
    else process.env.HOME = prevHome;
    if (prevUserProfile === undefined) delete process.env.USERPROFILE;
    else process.env.USERPROFILE = prevUserProfile;
    await rm(fakeHome, { recursive: true, force: true });
  }
});

test("suggestWorkspaceRoot returns the doc folder with confidence 'doc'", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    await mkdir(ws, { recursive: true });
    await writeFile(join(ws, "doc.docx"), "");
    assert.deepEqual(await suggestWorkspaceRoot(join(ws, "doc.docx")), {
      cwd: ws,
      confidence: "doc",
    });
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});

test("suggestWorkspaceRoot returns null for cloud URLs", async () => {
  assert.equal(
    await suggestWorkspaceRoot("https://contoso.sharepoint.com/sites/x/spec.docx"),
    null,
  );
});

test("ensureWorkspaceMarker drops a CLAUDE.md when none exists, then no-ops", async () => {
  const root = await makeTmp();
  try {
    const ws = join(root, "ws");
    await mkdir(ws);
    assert.equal(await ensureWorkspaceMarker(ws), true);
    assert.equal(await ensureWorkspaceMarker(ws), false);
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
    assert.equal(await ensureWorkspaceMarker(ws), false);
  } finally {
    await rm(root, { recursive: true, force: true });
  }
});
