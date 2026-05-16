// Unit tests for daemon/sessions.mjs — per-(host, workspace) session-id
// bookkeeping. Tests override $HOME via the environment so they touch a
// fresh tmpdir instead of ~/.claude/. sessions.mjs reads homedir() at
// import time, so the override must happen before the import — we use
// dynamic import() inside the test to do that cleanly.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile, readFile, rm } from "node:fs/promises";
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

async function readStateFile(fakeHome) {
  try {
    const raw = await readFile(join(fakeHome, ".claude", "office-addins", "sessions.json"), "utf8");
    return JSON.parse(raw);
  } catch (e) {
    if (e.code === "ENOENT") return null;
    throw e;
  }
}

test("getSessionId returns null when no state file exists", async () => {
  await withFakeHome(async ({ getSessionId }) => {
    assert.equal(await getSessionId("word", "/tmp/whatever"), null);
  });
});

test("save → get round-trip is keyed by host AND cwd", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }) => {
    await saveSessionId("word", "/tmp/folderA", "word-sid");
    await saveSessionId("excel", "/tmp/folderA", "excel-sid");

    // Same folder, different host → independent session ids.
    assert.equal(await getSessionId("word", "/tmp/folderA"), "word-sid");
    assert.equal(await getSessionId("excel", "/tmp/folderA"), "excel-sid");
    // A host with no saved session here → null.
    assert.equal(await getSessionId("word", "/tmp/folderB"), null);
  });
});

test("normalizeHost: a bad host is a no-op / null", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }) => {
    await saveSessionId("powerpoint", "/tmp/x", "nope");
    assert.equal(await getSessionId("powerpoint", "/tmp/x"), null);
    assert.equal(await getSessionId("word", "/tmp/x"), null);
  });
});

test("touchFolder does not clobber a saved session id", async () => {
  await withFakeHome(async ({ touchFolder, saveSessionId, getSessionId }) => {
    await touchFolder("/tmp/folderB");
    assert.equal(await getSessionId("word", "/tmp/folderB"), null);

    await saveSessionId("word", "/tmp/folderB", "w1");
    await touchFolder("/tmp/folderB");
    assert.equal(
      await getSessionId("word", "/tmp/folderB"),
      "w1",
      "touch must not blow away a saved session id",
    );
  });
});

test("touchFolder silently drops disallowed system home children (macOS only)", async () => {
  if (process.platform !== "darwin") return; // The deny-list is macOS-specific.
  await withFakeHome(async ({ touchFolder }, fakeHome) => {
    await touchFolder(join(fakeHome, "Library"));
    const state = await readStateFile(fakeHome);
    // touchFolder returns early for $HOME children — nothing persisted.
    assert.ok(
      state === null || !state.folders || Object.keys(state.folders).length === 0,
      "an OS-managed $HOME child must never be persisted",
    );
  });
});

test("migrates a v1 sessions.json: the un-hostable old id is dropped", async () => {
  await withFakeHome(async ({ getSessionId }, fakeHome) => {
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
    // The v1 id can't be attributed to a host, so it's intentionally not
    // resumable post-migration.
    assert.equal(await getSessionId("word", "/tmp/legacy"), null);
    assert.equal(await getSessionId("excel", "/tmp/legacy"), null);
  });
});
