// Unit tests for daemon/sessions.mjs — per-folder Agent SDK session-id
// bookkeeping. Tests override $HOME via the environment so they touch a
// fresh tmpdir instead of ~/.claude/. sessions.mjs reads homedir() at
// import time, so the override must happen before the import — we use
// dynamic import() inside the test to do that cleanly.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
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

test("getSessionForFolder returns null when no state file exists", async () => {
  await withFakeHome(async ({ getSessionForFolder }) => {
    const got = await getSessionForFolder("/tmp/whatever");
    assert.equal(got, null);
  });
});

test("save → get round-trip preserves session_id and updates last_used", async () => {
  await withFakeHome(async ({ saveSessionForFolder, getSessionForFolder }) => {
    await saveSessionForFolder("/tmp/folderA", "session-uuid-1");
    const got = await getSessionForFolder("/tmp/folderA");
    assert.equal(got.session_id, "session-uuid-1");
    assert.equal(got.display_name, "folderA");
    assert.match(got.last_used, /^\d{4}-\d{2}-\d{2}T/);
  });
});

test("touchFolder upserts: creates a new entry with no session_id, updates timestamp on existing", async () => {
  await withFakeHome(async ({ touchFolder, getSessionForFolder, saveSessionForFolder }) => {
    await touchFolder("/tmp/folderB");
    let got = await getSessionForFolder("/tmp/folderB");
    assert.equal(got.session_id, null);
    const firstTouched = got.last_used;

    // Tiny wait so the ISO string differs.
    await new Promise((r) => setTimeout(r, 10));

    await saveSessionForFolder("/tmp/folderB", "session-uuid-2");
    await touchFolder("/tmp/folderB");
    got = await getSessionForFolder("/tmp/folderB");
    assert.equal(got.session_id, "session-uuid-2", "touch should not blow away session_id");
    assert.notEqual(got.last_used, firstTouched, "touch should bump last_used");
  });
});

test("getRecentFolders returns most-recent-first, capped to the limit", async () => {
  await withFakeHome(async ({ touchFolder, getRecentFolders }) => {
    await touchFolder("/tmp/older");
    await new Promise((r) => setTimeout(r, 10));
    await touchFolder("/tmp/newer");
    const recent = await getRecentFolders(10);
    assert.equal(recent.length, 2);
    assert.equal(recent[0].cwd, "/tmp/newer", "newest first");
    assert.equal(recent[1].cwd, "/tmp/older");
  });
});

test("forgetFolder removes the entry", async () => {
  await withFakeHome(async ({ touchFolder, getSessionForFolder, forgetFolder }) => {
    await touchFolder("/tmp/forgetme");
    assert.notEqual(await getSessionForFolder("/tmp/forgetme"), null);
    await forgetFolder("/tmp/forgetme");
    assert.equal(await getSessionForFolder("/tmp/forgetme"), null);
  });
});

test("touchFolder silently drops disallowed system home children (macOS only)", async () => {
  if (process.platform !== "darwin") return; // The deny-list is macOS-specific.
  await withFakeHome(async ({ touchFolder, getRecentFolders }, fakeHome) => {
    const libraryPath = join(fakeHome, "Library");
    await touchFolder(libraryPath);
    const recent = await getRecentFolders(10);
    assert.equal(recent.length, 0, "Library should never appear in recent");
  });
});
