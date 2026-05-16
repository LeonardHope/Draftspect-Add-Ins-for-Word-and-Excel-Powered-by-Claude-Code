import { WebSocketServer } from "ws";
import { randomUUID } from "node:crypto";

// 60s, not 30s. Some Office.js ops legitimately run long on big docs —
// office_clear_highlights / a wide office_replace_section over a large
// document, batched office_highlight, a whole-sheet excel_read_range.
// 30s was tripping on real-sized documents. (Per-tool timeouts would be
// finer-grained but aren't worth the complexity yet.)
const TOOL_TIMEOUT_MS = 60_000;

const HOSTS = ["word", "excel"];
function normalizeHost(h) {
  return h === "word" || h === "excel" ? h : null;
}

// Tool names encode their host family: office_* → Word, excel_* → Excel.
// Used to route a tool call to the right pane when the caller doesn't
// pass an explicit host.
function hostForTool(name) {
  if (typeof name !== "string") return null;
  if (name.startsWith("office_")) return "word";
  if (name.startsWith("excel_")) return "excel";
  return null;
}

export function createBridge({
  port,
  extraHandlers = {},
  token,
  allowedOrigins = [],
  onHello,
  onUserMessage,
}) {
  if (!token) throw new Error("createBridge requires a token");
  const wss = new WebSocketServer({
    port,
    // Bind to loopback only. The `ws` library defaults to 0.0.0.0 when
    // given just `port`, which exposes the bridge to the local network.
    // This is a local-only IPC channel (taskpane <-> daemon on the same
    // machine), so it should never be reachable off-host regardless of
    // the token + origin gates below. Matches the HTTP server in
    // index.mjs, which already binds 127.0.0.1.
    host: "127.0.0.1",
    // First gate: only allow upgrades from known origins (the taskpane's
    // origin = our own HTTP server's). Browsers honor the Origin header on
    // WebSocket upgrades; rejecting unknown origins blocks malicious local
    // web pages from driving the agent.
    verifyClient: (info, cb) => {
      const origin = info.req.headers.origin || "";
      if (allowedOrigins.length === 0 || allowedOrigins.includes(origin)) {
        cb(true);
      } else {
        console.warn(`[bridge] Rejecting WS upgrade from origin: ${origin || "(none)"}`);
        cb(false, 403, "Forbidden origin");
      }
    },
  });

  // One pane per Office host. Word and Excel are independent surfaces and
  // a user routinely has both open at once, so they get their own WS pane
  // (and their own context + message queue). A connection is "unbound"
  // until its hello names a host; a same-host reconnect replaces the
  // prior pane (Office reloaded the panel), a different host coexists.
  // (The old single-`activeWs` model made the two hosts fight: each
  // hello closed the other's pane, which auto-reconnected and closed
  // this one — an endless connect/disconnect ping-pong.)
  const panes = new Map(); // host -> { ws, context }

  function emptyContext(host = null) {
    return { activeDoc: null, selection: null, trackChangesMode: "always", host };
  }

  function paneWs(host) {
    const p = panes.get(host);
    return p && p.ws.readyState === p.ws.OPEN ? p.ws : null;
  }

  // Tool calls awaiting their tool_result. Keyed by id; each remembers the
  // ws it was dispatched on so a stale pane's disconnect only rejects its
  // own in-flight calls.
  const pendingTools = new Map();

  // Per-host queue of pending user messages. Waiters are {resolve, reject}
  // so we can reject pending awaits when that host's session is aborted
  // (otherwise the suspended generator from the aborted session would
  // consume the next user message, starving the new session).
  const queues = new Map(); // host -> { queue: [], waiters: [] }
  function queueFor(host) {
    let q = queues.get(host);
    if (!q) {
      q = { queue: [], waiters: [] };
      queues.set(host, q);
    }
    return q;
  }

  function pushUserMessage(text, host) {
    const ctx = panes.get(host)?.context ?? emptyContext(host);
    const payload = { text, context: { ...ctx, host } };
    const q = queueFor(host);
    if (q.waiters.length > 0) {
      q.waiters.shift().resolve(payload);
    } else {
      q.queue.push(payload);
    }
  }

  function nextUserMessage(host) {
    const q = queueFor(host);
    if (q.queue.length > 0) return Promise.resolve(q.queue.shift());
    return new Promise((resolve, reject) => q.waiters.push({ resolve, reject }));
  }

  // Reject all pending waiters and drop queued messages for ONE host.
  // Called before that host's session swaps so the new session starts
  // clean with no zombie consumer — and, crucially, without touching the
  // OTHER host's queue (a message just enqueued for the host we're
  // switching TO must survive).
  function clearUserMessages(host) {
    const q = queues.get(host);
    if (!q) return;
    while (q.waiters.length > 0) {
      q.waiters.shift().reject(new Error("Session aborted"));
    }
    q.queue.length = 0;
  }

  // Merge inbound doc/selection/track-changes fields into a pane's
  // context. `hello` resets doc+selection (a fresh taskpane connection —
  // absent fields mean null), while `context_update` patches (only fields
  // present on the message change). track-changes is "set only when
  // valid" in both modes. host is fixed at bind time, never re-derived.
  function mergeContext(ctx, msg, { reset = false } = {}) {
    if (reset) {
      ctx.activeDoc = msg.active_doc ?? null;
      ctx.selection = msg.selection ?? null;
    } else {
      if (msg.active_doc !== undefined) ctx.activeDoc = msg.active_doc;
      if (msg.selection !== undefined) ctx.selection = msg.selection;
    }
    if (typeof msg.track_changes_mode === "string") {
      ctx.trackChangesMode = msg.track_changes_mode;
    }
  }

  function send(obj, host) {
    const ws = paneWs(host);
    if (!ws) {
      console.warn(`[bridge] No ${host ?? "?"} taskpane; dropping`, obj.type);
      return;
    }
    ws.send(JSON.stringify(obj));
  }

  async function callTaskpaneTool(name, args, host = null) {
    const h = normalizeHost(host) ?? hostForTool(name);
    const ws = paneWs(h);
    if (!ws) {
      throw new Error(`Cannot call tool ${name}: no ${h ?? "?"} taskpane connected`);
    }
    const id = randomUUID();
    // Capture the WS this call is dispatched on. On close we only reject
    // pending calls belonging to that specific WS — so a stale pane
    // disconnecting after a fresh one is active doesn't kill live work.
    const ownerWs = ws;
    const promise = new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        pendingTools.delete(id);
        reject(new Error(`Tool ${name} timed out after ${TOOL_TIMEOUT_MS}ms`));
      }, TOOL_TIMEOUT_MS);
      pendingTools.set(id, { ws: ownerWs, resolve, reject, timer });
    });
    ownerWs.send(JSON.stringify({ type: "tool_call", id, name, args }));
    return await promise;
  }

  function sendAssistantText(delta, host) {
    send({ type: "assistant_text", delta }, host);
  }

  function sendAssistantEvent(event, host) {
    send({ type: "assistant_event", ...event }, host);
  }

  wss.on("connection", (ws, req) => {
    console.log(
      `[bridge] WS connection from ${req.socket.remoteAddress} (origin: ${req.headers.origin || "(none)"})`,
    );

    // Second gate: first message must be a hello with the correct token.
    // Until that arrives we don't trust the connection — any other message
    // type is rejected and the socket is closed. The taskpane fetches the
    // token from /bridge-token over the same-origin HTTP server before
    // opening the WS. A connection is bound to exactly one host (named in
    // its hello) for its lifetime.
    let authed = false;
    let boundHost = null;

    // True iff this ws is still the live pane for the host it bound to.
    // A superseded (same-host reconnect) ws stays open just long enough
    // to be force-closed; drop any stale frames it sends meanwhile.
    function isLivePane() {
      return authed && panes.get(boundHost)?.ws === ws;
    }

    ws.on("message", (raw) => {
      let msg;
      try {
        msg = JSON.parse(raw.toString());
      } catch (e) {
        console.error("[bridge] Bad JSON from taskpane:", e.message);
        return;
      }

      if (!authed) {
        if (msg.type !== "hello") {
          console.warn(`[bridge] Pre-auth message ${msg.type}; closing`);
          ws.close(4001, "Authenticate first");
          return;
        }
        if (msg.token !== token) {
          console.warn("[bridge] hello with invalid token; closing");
          ws.close(4002, "Invalid token");
          return;
        }
        const host = normalizeHost(msg.host);
        if (!host) {
          console.warn(`[bridge] hello without a valid host (${msg.host}); closing`);
          ws.close(4005, "hello must name host word|excel");
          return;
        }
        // Replace only a prior pane for the SAME host (Office reloaded the
        // panel, or a second window of the same app). The other host's
        // pane is left untouched so Word and Excel coexist.
        const prior = panes.get(host);
        if (prior && prior.ws !== ws && prior.ws.readyState === prior.ws.OPEN) {
          console.log(`[bridge] New ${host} pane authed; closing prior ${host} pane`);
          try {
            prior.ws.close(4003, "Replaced by new pane");
          } catch {}
        }
        authed = true;
        boundHost = host;
        panes.set(host, { ws, context: emptyContext(host) });
      } else if (!isLivePane()) {
        // Superseded by a newer same-host pane; its close is in flight.
        // Drop stale frames rather than letting them mutate context or
        // queue a user message for a pane that's going away.
        console.warn(`[bridge] Post-auth ${msg.type} from superseded ${boundHost} pane; ignoring`);
        return;
      }

      const pane = panes.get(boundHost);

      switch (msg.type) {
        case "hello": {
          mergeContext(pane.context, msg, { reset: true });
          ws.send(
            JSON.stringify({
              type: "welcome",
              session_id: randomUUID(),
              server_version: "0.1.0",
            }),
          );
          console.log(
            `[bridge] hello received; host: ${boundHost}; active doc: ${pane.context.activeDoc}`,
          );
          // Let the daemon replay this pane's transcript (and anything
          // else it wants on connect). Fire-and-forget; never block.
          if (onHello) {
            Promise.resolve(onHello(boundHost)).catch((err) =>
              console.warn("[bridge] onHello failed:", err?.message ?? err),
            );
          }
          break;
        }
        case "user_message": {
          console.log(`[bridge] user_message (${boundHost}):`, msg.text);
          // Give the daemon a chance to (lazily) start/swap the agent
          // loop to this host BEFORE the message is queued, so the right
          // loop consumes it. Fire-and-forget; ordering holds because
          // session start is deferred to setImmediate, after this push.
          if (onUserMessage) {
            Promise.resolve(onUserMessage(boundHost)).catch((err) =>
              console.warn("[bridge] onUserMessage failed:", err?.message ?? err),
            );
          }
          pushUserMessage(msg.text, boundHost);
          break;
        }
        case "context_update": {
          mergeContext(pane.context, msg);
          break;
        }
        case "tool_result": {
          const pending = pendingTools.get(msg.id);
          if (!pending) {
            console.warn("[bridge] tool_result for unknown id:", msg.id);
            return;
          }
          clearTimeout(pending.timer);
          pendingTools.delete(msg.id);
          if (msg.ok) {
            pending.resolve(msg.result);
          } else {
            pending.reject(new Error(msg.error ?? "Unknown tool error"));
          }
          break;
        }
        case "ping": {
          ws.send(JSON.stringify({ type: "pong" }));
          break;
        }
        default: {
          // Custom handlers registered by the daemon (e.g. settings /
          // refs). They receive the message, a `reply(obj)` shortcut that
          // sends back over this same WS, and the bound host so per-pane
          // actions (set_cwd, stop_agent, …) act on the right surface.
          const handler = extraHandlers[msg.type];
          if (handler) {
            const reply = (obj) => ws.send(JSON.stringify(obj));
            Promise.resolve()
              .then(() => handler(msg, reply, boundHost))
              .catch((err) => {
                console.error(`[bridge] handler for ${msg.type} threw:`, err);
                reply({
                  type: msg.type + "_result",
                  ok: false,
                  error: err.message ?? String(err),
                  request_id: msg.request_id,
                });
              });
          } else {
            console.warn("[bridge] Unknown message type:", msg.type);
          }
        }
      }
    });

    ws.on("close", () => {
      console.log(`[bridge] Taskpane disconnected${boundHost ? ` (${boundHost})` : ""}`);
      // Only clear the pane mapping if THIS ws still owns it (a superseded
      // ws closing must not evict the fresh same-host pane).
      if (boundHost && panes.get(boundHost)?.ws === ws) {
        panes.delete(boundHost);
      }
      // Reject only the pending tool calls dispatched on THIS ws. A
      // different (newer) WS may have its own in-flight calls; leave them.
      for (const [id, p] of pendingTools) {
        if (p.ws === ws) {
          clearTimeout(p.timer);
          p.reject(new Error("Taskpane disconnected"));
          pendingTools.delete(id);
        }
      }
    });

    ws.on("error", (err) => {
      console.error("[bridge] WS error:", err.message);
    });
  });

  console.log(
    `[bridge] WebSocket server listening on ws://127.0.0.1:${port} (origin allowlist: ${allowedOrigins.join(", ") || "<empty>"})`,
  );

  return {
    nextUserMessage,
    clearUserMessages,
    callTaskpaneTool,
    sendAssistantText,
    sendAssistantEvent,
    sendToTaskpane: send,
    getContext: (host) => ({ ...(panes.get(host)?.context ?? emptyContext(host)) }),
    isTaskpaneConnected: (host) => (host ? !!paneWs(host) : HOSTS.some((h) => !!paneWs(h))),
    // Test/observability surface: the bound address (null until listening)
    // and a clean shutdown. Used by the loopback-bind unit test.
    address: () => wss.address(),
    close: () => new Promise((resolve) => wss.close(resolve)),
  };
}
