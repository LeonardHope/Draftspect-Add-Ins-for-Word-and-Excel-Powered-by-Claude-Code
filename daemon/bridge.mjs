import { WebSocketServer } from "ws";
import { randomUUID } from "node:crypto";

// 60s, not 30s. Some Office.js ops legitimately run long on big docs —
// office_clear_highlights / a wide office_replace_section over a large
// document, batched office_highlight, a whole-sheet excel_read_range.
// 30s was tripping on real-sized documents. (Per-tool timeouts would be
// finer-grained but aren't worth the complexity yet.)
const TOOL_TIMEOUT_MS = 60_000;

export function createBridge({ port, extraHandlers = {}, token, allowedOrigins = [] }) {
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

  // POC: single active taskpane connection. Last connect wins.
  let activeWs = null;
  let activeContext = {
    activeDoc: null,
    selection: null,
    trackChangesMode: "always",
    host: null, // "word" | "excel" — set from the taskpane's hello
  };

  // Tool calls awaiting their tool_result. Keyed by id.
  const pendingTools = new Map();

  // Queue of pending user messages. Waiters are {resolve, reject} so we can
  // reject pending awaits when the agent session is aborted (otherwise the
  // suspended generator from the aborted session would consume the next
  // user message, starving the new session).
  const userMessageQueue = [];
  const userMessageWaiters = [];

  function pushUserMessage(text) {
    if (userMessageWaiters.length > 0) {
      const { resolve } = userMessageWaiters.shift();
      resolve({ text, context: { ...activeContext } });
    } else {
      userMessageQueue.push({ text, context: { ...activeContext } });
    }
  }

  function nextUserMessage() {
    if (userMessageQueue.length > 0) {
      return Promise.resolve(userMessageQueue.shift());
    }
    return new Promise((resolve, reject) => userMessageWaiters.push({ resolve, reject }));
  }

  // Reject all pending waiters and drop queued messages. Called before a
  // session swap so the new session starts cleanly with no zombie consumers.
  function clearUserMessages() {
    while (userMessageWaiters.length > 0) {
      const { reject } = userMessageWaiters.shift();
      reject(new Error("Session aborted"));
    }
    userMessageQueue.length = 0;
  }

  function send(obj) {
    if (!activeWs || activeWs.readyState !== activeWs.OPEN) {
      console.warn("[bridge] No active taskpane; dropping", obj.type);
      return;
    }
    activeWs.send(JSON.stringify(obj));
  }

  async function callTaskpaneTool(name, args) {
    if (!activeWs || activeWs.readyState !== activeWs.OPEN) {
      throw new Error(`Cannot call tool ${name}: no taskpane connected`);
    }
    const id = randomUUID();
    // Capture the WS this tool call is dispatched on. On close, we only
    // reject pending calls belonging to that specific WS — so a stale
    // taskpane disconnecting after a new one is already active doesn't
    // kill the new session's in-flight work.
    const ownerWs = activeWs;
    const promise = new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        pendingTools.delete(id);
        reject(new Error(`Tool ${name} timed out after ${TOOL_TIMEOUT_MS}ms`));
      }, TOOL_TIMEOUT_MS);
      pendingTools.set(id, { ws: ownerWs, resolve, reject, timer });
    });
    send({ type: "tool_call", id, name, args });
    return await promise;
  }

  function sendAssistantText(delta) {
    send({ type: "assistant_text", delta });
  }

  function sendAssistantEvent(event) {
    send({ type: "assistant_event", ...event });
  }

  wss.on("connection", (ws, req) => {
    console.log(`[bridge] WS connection from ${req.socket.remoteAddress} (origin: ${req.headers.origin || "(none)"})`);

    // Second gate: first message must be a hello with the correct token.
    // Until that arrives we don't trust the connection — any other message
    // type is rejected and the socket is closed. The taskpane fetches the
    // token from /bridge-token over the same-origin HTTP server before
    // opening the WS.
    let authed = false;

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
        // Close the prior active pane (if any) before promoting this ws —
        // otherwise an old Word/Excel taskpane keeps feeding user_message /
        // context_update frames after a new pane connects from a different
        // document, racing with the legitimate session.
        if (activeWs && activeWs !== ws && activeWs.readyState === activeWs.OPEN) {
          console.log("[bridge] New pane authed; closing prior pane");
          try { activeWs.close(4003, "Replaced by new pane"); } catch {}
        }
        authed = true;
        activeWs = ws;
      } else if (ws !== activeWs) {
        // Should not happen — once a ws is authed and superseded, the old
        // socket gets closed above. But if a stale message lands in flight
        // between authing and the close handler, drop it rather than letting
        // it overwrite activeContext or queue a user message.
        console.warn(`[bridge] Post-auth ${msg.type} from non-active ws; ignoring`);
        return;
      }

      switch (msg.type) {
        case "hello": {
          activeContext.activeDoc = msg.active_doc ?? null;
          activeContext.selection = msg.selection ?? null;
          if (typeof msg.track_changes_mode === "string") {
            activeContext.trackChangesMode = msg.track_changes_mode;
          }
          if (msg.host === "word" || msg.host === "excel") {
            activeContext.host = msg.host;
          }
          ws.send(JSON.stringify({
            type: "welcome",
            session_id: randomUUID(),
            server_version: "0.1.0",
          }));
          console.log(`[bridge] hello received; host: ${activeContext.host ?? "?"}; active doc: ${activeContext.activeDoc}`);
          break;
        }
        case "user_message": {
          console.log("[bridge] user_message:", msg.text);
          pushUserMessage(msg.text);
          break;
        }
        case "context_update": {
          if (msg.active_doc !== undefined) activeContext.activeDoc = msg.active_doc;
          if (msg.selection !== undefined) activeContext.selection = msg.selection;
          if (typeof msg.track_changes_mode === "string") {
            activeContext.trackChangesMode = msg.track_changes_mode;
          }
          if (msg.host === "word" || msg.host === "excel") {
            activeContext.host = msg.host;
          }
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
          // Custom handlers registered by the daemon (e.g. settings / refs).
          // They receive the message and a `reply(obj)` shortcut that sends
          // back over the same WS.
          const handler = extraHandlers[msg.type];
          if (handler) {
            const reply = (obj) => ws.send(JSON.stringify(obj));
            Promise.resolve()
              .then(() => handler(msg, reply))
              .catch(err => {
                console.error(`[bridge] handler for ${msg.type} threw:`, err);
                reply({ type: msg.type + "_result", ok: false, error: err.message ?? String(err), request_id: msg.request_id });
              });
          } else {
            console.warn("[bridge] Unknown message type:", msg.type);
          }
        }
      }
    });

    ws.on("close", () => {
      console.log("[bridge] Taskpane disconnected");
      if (activeWs === ws) activeWs = null;
      // Reject only the pending tool calls that were dispatched on THIS ws.
      // A different (newer) WS may have its own in-flight calls; leave them.
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

  console.log(`[bridge] WebSocket server listening on ws://127.0.0.1:${port} (origin allowlist: ${allowedOrigins.join(", ") || "<empty>"})`);

  return {
    nextUserMessage,
    clearUserMessages,
    callTaskpaneTool,
    sendAssistantText,
    sendAssistantEvent,
    getContext: () => ({ ...activeContext }),
    isTaskpaneConnected: () => activeWs?.readyState === activeWs?.OPEN,
    // Test/observability surface: the bound address (null until listening)
    // and a clean shutdown. Used by the loopback-bind unit test.
    address: () => wss.address(),
    close: () => new Promise((resolve) => wss.close(resolve)),
  };
}
