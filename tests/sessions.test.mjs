// Unit tests for daemon/sessions.mjs — per-(host, workspace) session-id
// bookkeeping. Tests override $HOME via the environment so they touch a
// fresh tmpdir instead of ~/.claude/. sessions.mjs reads homedir() at
// import time, so the override must happen before the import — we use
// dynamic import() inside the test to do that cleanly.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

async function withFakeHome(fn) {
  const fakeHome = await mkdtemp(join(tmpdir(), "cc-office-sess-test-"));
  const prevHome = process.env.HOME;
  const prevUserProfile = process.env.USERPROFILE;
  process.env.HOME = fakeHome;
  process.env.USERPROFILE = fakeHome;
  try {
    // Cache-bust the dynamic import so each test gets a freshly-evaluated
    // module reading the new $HOME.
    const cacheBuster = "?t=" + Date.now() + Math.random();
    const mod = await import("../daemon/sessions.mjs" + cacheBuster);
    await fn(mod, fakeHome);
  } finally {
    if (prevHome === undefined) delete process.env.HOME;
    else process.env.HOME = prevHome;
    if (prevUserProfile === undefined) delete process.env.USERPROFILE;
    else process.env.USERPROFILE = prevUserProfile;
    await rm(fakeHome, { recursive: true, force: true });
  }
}

test("getSessionId returns null when no state file exists", async () => {
  await withFakeHome(async ({ getSessionId }) => {
    assert.equal(await getSessionId("word", "/tmp/whatever"), null);
  });
});

test("save → get round-trip is keyed by host AND cwd", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId, getRecentFolders }) => {
    await saveSessionId("word", "/tmp/folderA", "word-sid");
    await saveSessionId("excel", "/tmp/folderA", "excel-sid");

    // Same folder, different host → independent session ids.
    assert.equal(await getSessionId("word", "/tmp/folderA"), "word-sid");
    assert.equal(await getSessionId("excel", "/tmp/folderA"), "excel-sid");
    // A host with no saved session here → null.
    assert.equal(await getSessionId("word", "/tmp/folderB"), null);

    // The folder shows once in recents (host-agnostic switcher list).
    const recent = await getRecentFolders(10);
    assert.equal(recent.length, 1);
    assert.equal(recent[0].cwd, "/tmp/folderA");
    assert.equal(recent[0].display_name, "folderA");
    assert.match(recent[0].last_used, /^\d{4}-\d{2}-\d{2}T/);
  });
});

test("normalizeHost: a bad host is a no-op / null", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }) => {
    await saveSessionId("powerpoint", "/tmp/x", "nope");
    assert.equal(await getSessionId("powerpoint", "/tmp/x"), null);
    assert.equal(await getSessionId("word", "/tmp/x"), null);
  });
});

test("touchFolder creates a recents entry without any session; save keeps it", async () => {
  await withFakeHome(async ({ touchFolder, getRecentFolders, saveSessionId, getSessionId }) => {
    await touchFolder("/tmp/folderB");
    let recent = await getRecentFolders(10);
    assert.equal(recent.length, 1);
    assert.equal(recent[0].cwd, "/tmp/folderB");
    assert.equal(await getSessionId("word", "/tmp/folderB"), null);
    const firstTouched = recent[0].last_used;

    await new Promise((r) => setTimeout(r, 10));

    await saveSessionId("word", "/tmp/folderB", "w1");
    await touchFolder("/tmp/folderB");
    assert.equal(
      await getSessionId("word", "/tmp/folderB"),
      "w1",
      "touch must not blow away a saved session id",
    );
    recent = await getRecentFolders(10);
    assert.notEqual(recent[0].last_used, firstTouched, "touch bumps last_used");
  });
});

test("getRecentFolders is newest-first, cwd-deduped, capped", async () => {
  await withFakeHome(async ({ touchFolder, saveSessionId, getRecentFolders }) => {
    await touchFolder("/tmp/older");
    await new Promise((r) => setTimeout(r, 10));
    // Two hosts in the same newer folder → still ONE recents row.
    await saveSessionId("word", "/tmp/newer", "w");
    await saveSessionId("excel", "/tmp/newer", "e");
    const recent = await getRecentFolders(10);
    assert.equal(recent.length, 2);
    assert.equal(recent[0].cwd, "/tmp/newer", "newest first");
    assert.equal(recent[1].cwd, "/tmp/older");
  });
});

test("forgetFolder removes the folder and all its per-host sessions", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId, forgetFolder, getRecentFolders }) => {
    await saveSessionId("word", "/tmp/forgetme", "w");
    await saveSessionId("excel", "/tmp/forgetme", "e");
    await forgetFolder("/tmp/forgetme");
    assert.equal(await getSessionId("word", "/tmp/forgetme"), null);
    assert.equal(await getSessionId("excel", "/tmp/forgetme"), null);
    assert.equal((await getRecentFolders(10)).length, 0);
  });
});

test("touchFolder silently drops disallowed system home children (macOS only)", async () => {
  if (process.platform !== "darwin") return; // The deny-list is macOS-specific.
  await withFakeHome(async ({ touchFolder, getRecentFolders }, fakeHome) => {
    await touchFolder(join(fakeHome, "Library"));
    assert.equal((await getRecentFolders(10)).length, 0, "Library never in recents");
  });
});

test("migrates a v1 sessions.json: folder kept in recents, old id dropped", async () => {
  await withFakeHome(async ({ getRecentFolders, getSessionId }, fakeHome) => {
    const dir = join(fakeHome, ".claude", "office-addins");
    await mkdir(dir, { recursive: true });
    const v1 = {
      version: 1,
      folders: {
        "/tmp/legacy": {
          session_id: "old-uuid",
          last_used: "2026-05-01T00:00:00.000Z",
          display_name: "legacy",
        },
      },
    };
    await writeFile(join(dir, "sessions.json"), JSON.stringify(v1));
    const recent = await getRecentFolders(10);
    assert.equal(recent.length, 1);
    assert.equal(recent[0].cwd, "/tmp/legacy");
    assert.equal(recent[0].display_name, "legacy");
    // The un-hostable v1 id is intentionally not resumable post-migration.
    assert.equal(await getSessionId("word", "/tmp/legacy"), null);
    assert.equal(await getSessionId("excel", "/tmp/legacy"), null);
  });
});
