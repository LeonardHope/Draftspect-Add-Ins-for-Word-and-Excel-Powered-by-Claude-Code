// Unit test for daemon/bridge.mjs network binding.
//
// Regression guard for the WS loopback-bind security fix: the `ws`
// library binds 0.0.0.0 when constructed with only `port`, exposing the
// agent bridge to the local network. createBridge must pass
// host: "127.0.0.1" so the bridge is loopback-only.

import { test } from "node:test";
import assert from "node:assert/strict";

import { createBridge } from "../daemon/bridge.mjs";

// Wait for the ws server to finish binding (it binds asynchronously).
async function waitForListening(bridge, timeoutMs = 3000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (bridge.address()) return;
    await new Promise((r) => setTimeout(r, 20));
  }
  throw new Error("bridge did not start listening within timeout");
}

test("WS bridge binds to loopback (127.0.0.1), not all interfaces", async () => {
  // port: 0 → ephemeral free port, so the test never collides with a
  // running daemon or another test.
  const bridge = createBridge({ port: 0, token: "test-token", allowedOrigins: [] });
  try {
    await waitForListening(bridge);
    const addr = bridge.address();
    assert.ok(addr && typeof addr === "object", "address() should return the bound socket info");
    assert.equal(addr.address, "127.0.0.1", "WS bridge must bind loopback, not 0.0.0.0");
  } finally {
    await bridge.close();
  }
});

test("createBridge requires a token", () => {
  assert.throws(() => createBridge({ port: 0 }), /requires a token/);
});
