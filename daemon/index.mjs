import { query } from "@anthropic-ai/claude-agent-sdk";
import { readFile, writeFile, mkdir, chmod } from "node:fs/promises";
import { randomBytes } from "node:crypto";
import { createServer } from "node:http";
import { fileURLToPath } from "node:url";
import { dirname, resolve, extname, join, sep } from "node:path";
import { homedir } from "node:os";
import { createBridge } from "./bridge.mjs";
import { createOfficeBridgeMcp } from "./office-tools.mjs";
import { resolveWorkspaceRoot, suggestWorkspaceRoot, ensureWorkspaceMarker } from "./workspace.mjs";
import { randomUUID } from "node:crypto";
import {
  getSessionForFolder,
  saveSessionForFolder,
  touchFolder,
  getRecentFolders,
  forgetFolder,
} from "./sessions.mjs";
import { readTranscript } from "./transcript.mjs";
import { getContextEntries, setContextEntries } from "./context.mjs";
import { stat } from "node:fs/promises";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = resolve(__dirname, "..");

const WS_PORT = 47823;
const HTTP_PORT = 47824;
const HTTP_ORIGIN = `http://127.0.0.1:${HTTP_PORT}`;
const TOKEN_FILE = join(homedir(), ".claude", "office-addins", "bridge-token");

// Bridge token — random per-daemon-start. The taskpane fetches it from the
// HTTP server's /bridge-token endpoint (same-origin, CORS-restricted) and
// includes it in the first WS hello. Any WS that doesn't present this token
// (or comes from an unknown origin) is closed.
const BRIDGE_TOKEN = randomBytes(24).toString("hex");
{
  await mkdir(dirname(TOKEN_FILE), { recursive: true });
  await writeFile(TOKEN_FILE, BRIDGE_TOKEN, { mode: 0o600 });
  try {
    await chmod(TOKEN_FILE, 0o600);
  } catch {}
  console.log(`[daemon] Bridge token written to ${TOKEN_FILE}`);
}

// ---------------------------------------------------------------------------
// CLI args
// ---------------------------------------------------------------------------
const matterFolder = process.argv[2]
  ? resolve(process.argv[2])
  : process.env.MATTER_FOLDER
    ? resolve(process.env.MATTER_FOLDER)
    : process.cwd();

console.log(`[daemon] Workspace folder (agent cwd): ${matterFolder}`);

// ---------------------------------------------------------------------------
// HTTP server: serve the taskpane assets so Word can load them.
// ---------------------------------------------------------------------------
const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".json": "application/json; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
};

const taskpaneDir = join(PROJECT_ROOT, "taskpane");

const http = createServer(async (req, res) => {
  // Restrictive CORS: only the taskpane's same origin (HTTP_ORIGIN) gets the
  // Access-Control-Allow-Origin header. Other origins (malicious web pages
  // running in a regular browser tab) hit a no-CORS-header response and the
  // browser blocks them from reading it. Same-origin requests from the
  // taskpane itself don't go through CORS at all.
  const reqOrigin = req.headers.origin || "";
  if (reqOrigin === HTTP_ORIGIN) {
    res.setHeader("Access-Control-Allow-Origin", reqOrigin);
  }
  // No-cache: Word's webview likes to cache aggressively. During dev we want
  // every reload to pick up the latest taskpane JS/CSS/HTML.
  res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0");
  res.setHeader("Pragma", "no-cache");
  res.setHeader("Expires", "0");
  if (req.method === "OPTIONS") {
    res.writeHead(204).end();
    return;
  }

  try {
    const urlPath = (req.url || "/").split("?")[0];

    // Token endpoint: serves the per-daemon bridge token to the taskpane.
    // Only same-origin (i.e. cross-origin requests get blocked by CORS).
    if (urlPath === "/bridge-token") {
      res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
      res.end(BRIDGE_TOKEN);
      return;
    }

    const relPath = urlPath === "/" ? "/index.html" : urlPath;
    const fsPath = join(taskpaneDir, relPath);
    // Containment check. `join` already normalizes `../`, so the obvious
    // traversal is blocked — but a bare startsWith(taskpaneDir) would also
    // accept a sibling like `<…>/taskpane-evil/x`. Require an exact match
    // OR a path under `taskpaneDir` + separator.
    if (fsPath !== taskpaneDir && !fsPath.startsWith(taskpaneDir + sep)) {
      res.writeHead(403).end("Forbidden");
      return;
    }
    const data = await readFile(fsPath);
    const mime = MIME[extname(fsPath).toLowerCase()] || "application/octet-stream";
    res.writeHead(200, { "Content-Type": mime });
    res.end(data);
  } catch (err) {
    if (err.code === "ENOENT") {
      res.writeHead(404).end("Not found");
    } else {
      console.error("[http]", err);
      res.writeHead(500).end("Server error");
    }
  }
});

http.listen(HTTP_PORT, "127.0.0.1", () => {
  console.log(`[daemon] HTTP server listening on http://127.0.0.1:${HTTP_PORT}/`);
});

// ---------------------------------------------------------------------------
// IPC channel back to the Electron main process. Used to ask main.mjs to
// show a native macOS open panel for folder/file picking — synthetic
// in-page modals can't navigate to Google Drive, iCloud, recent items, or
// any of the other sources macOS users expect in NSOpenPanel. The channel
// is fd 3 (added in main.mjs's spawn options), wired through Node IPC.
// ---------------------------------------------------------------------------
const pendingPicks = new Map(); // id -> { resolve, reject, timer }
const PICK_TIMEOUT_MS = 5 * 60_000; // 5 min; the user might leave the dialog open

if (process.send) {
  process.on("message", (msg) => {
    if (msg?.type !== "pick_path_result") return;
    const entry = pendingPicks.get(msg.id);
    if (!entry) return;
    pendingPicks.delete(msg.id);
    clearTimeout(entry.timer);
    entry.resolve(msg);
  });
}

function pickPathFromMain({ include_files, default_path, title, button_label }) {
  if (!process.send) {
    return Promise.reject(new Error("No IPC channel to Electron main process"));
  }
  const id = randomUUID();
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingPicks.delete(id);
      reject(new Error("Picker timed out"));
    }, PICK_TIMEOUT_MS);
    pendingPicks.set(id, { resolve, reject, timer });
    process.send({ type: "pick_path", id, include_files, default_path, title, button_label });
  });
}

// Hoisted above createBridge so bridge handlers that fire during the
// top-level awaits below (preflightHttpMcpServers, etc.) can call
// getCurrentCwd without hitting a TDZ on this binding.
let currentSession = null; // { cwd, sessionId, abortController, settled }

// Resolve the session id to replay. Prefer the live session's id (set once
// the SDK reports init); fall back to the id persisted for the current cwd
// (covers the window before the SDK has re-inited on a fresh launch).
async function resolveReplaySessionId() {
  if (currentSession?.sessionId) return currentSession.sessionId;
  const cwd = getCurrentCwd();
  if (!cwd) return null;
  try {
    const saved = await getSessionForFolder(cwd);
    return saved?.session_id ?? null;
  } catch {
    return null;
  }
}

// Reconstruct the prior conversation from the resumed session's .jsonl and
// push it to the taskpane. Sent on every taskpane hello and after each
// workspace switch (cwd_changed). Empty events => fresh chat, no divider.
async function sendTranscriptReplay() {
  try {
    const sessionId = await resolveReplaySessionId();
    const { events, truncated } = sessionId
      ? await readTranscript(sessionId, { maxEvents: 200 })
      : { events: [], truncated: false };
    bridge.sendToTaskpane({
      type: "transcript_replay",
      session_id: sessionId ?? null,
      truncated,
      events,
    });
  } catch (err) {
    console.warn("[daemon] transcript replay failed:", err?.message ?? err);
  }
}

// Called on every taskpane hello. The session may have been started before
// any pane connected (host unknown → both tool families registered) or for
// a different host (the user moved to a Word doc after an Excel one). If
// the connected host doesn't match the session's registered tool family,
// restart the session for the same cwd+sessionId narrowed to that host —
// the restart resumes the conversation and emits its own cwd_changed +
// transcript replay, so we skip the standalone replay in that case.
// Otherwise just replay the transcript for the (re)connected pane.
function maybeRenarrowForHost() {
  const host = bridge.getContext().host ?? null;
  if (host && currentSession && currentSession.toolHost !== host) {
    const { cwd, sessionId } = currentSession;
    console.log(
      `[daemon] Host is ${host}; session tools were ${currentSession.toolHost ?? "both"} — re-narrowing`,
    );
    // Defer past the bridge's message handler (matches the post-stop /
    // stream-end restart pattern).
    setImmediate(() => {
      startSessionForFolder(cwd, sessionId, { host }).catch((err) =>
        console.error("[daemon] host re-narrow restart failed:", err?.message ?? err),
      );
    });
    return;
  }
  sendTranscriptReplay().catch(() => {});
}

// ---------------------------------------------------------------------------
// WebSocket bridge.
// ---------------------------------------------------------------------------
const bridge = createBridge({
  port: WS_PORT,
  token: BRIDGE_TOKEN,
  allowedOrigins: [HTTP_ORIGIN],
  onHello: () => maybeRenarrowForHost(),
  extraHandlers: {
    pick_path: async (msg, reply) => {
      try {
        const result = await pickPathFromMain({
          include_files: !!msg.include_files,
          default_path: msg.default_path || null,
          title: msg.title || null,
          button_label: msg.button_label || null,
        });
        reply({ type: "pick_path_result", request_id: msg.request_id, ...result });
      } catch (e) {
        reply({
          type: "pick_path_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    set_cwd: async (msg, reply) => {
      try {
        let cwd;
        let explicitPick = false;
        if (msg.autodetect_from_doc) {
          const detected = await resolveWorkspaceRoot(msg.autodetect_from_doc);
          if (!detected)
            throw new Error("Could not auto-detect a workspace folder from that doc path");
          cwd = detected;
        } else if (msg.cwd) {
          cwd = msg.cwd;
          explicitPick = true;
        } else {
          throw new Error("set_cwd requires `cwd` or `autodetect_from_doc`");
        }
        const resolved = await switchFolder(cwd);
        // Drop a CLAUDE.md marker on explicit user pick so the next open of
        // any doc in this folder auto-detects silently.
        let markerCreated = false;
        if (explicitPick) {
          try {
            markerCreated = await ensureWorkspaceMarker(resolved);
          } catch (e) {
            console.warn(`[daemon] could not create CLAUDE.md in ${resolved}: ${e.message}`);
          }
        }
        reply({
          type: "set_cwd_result",
          ok: true,
          cwd: resolved,
          marker_created: markerCreated,
          request_id: msg.request_id,
        });
      } catch (e) {
        reply({ type: "set_cwd_result", ok: false, error: e.message, request_id: msg.request_id });
      }
    },
    suggest_workspace: async (msg, reply) => {
      try {
        const suggestion = await suggestWorkspaceRoot(msg.doc_path || null);
        reply({
          type: "suggest_workspace_result",
          ok: true,
          suggestion,
          request_id: msg.request_id,
        });
      } catch (e) {
        reply({
          type: "suggest_workspace_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    get_cwd_state: async (msg, reply) => {
      const recent = await getRecentFolders();
      reply({
        type: "get_cwd_state_result",
        ok: true,
        current_cwd: getCurrentCwd(),
        recent,
        request_id: msg.request_id,
      });
    },
    stop_agent: async (msg, reply) => {
      // User clicked the Stop button in the taskpane. Abort the current
      // agent loop; the query() iterator's catch path sees AbortError and
      // (because we flag the session as intentionally interrupted) emits
      // a turn_complete event with interrupted=true so the taskpane flips
      // back to Ready and auto-restarts a fresh resuming loop.
      if (currentSession) {
        currentSession.interrupted = true;
        currentSession.abortController.abort();
        reply({ type: "stop_agent_result", ok: true, request_id: msg.request_id });
      } else {
        reply({
          type: "stop_agent_result",
          ok: false,
          error: "No active agent turn",
          request_id: msg.request_id,
        });
      }
    },
    forget_folder: async (msg, reply) => {
      try {
        await forgetFolder(msg.cwd);
        reply({ type: "forget_folder_result", ok: true, request_id: msg.request_id });
      } catch (e) {
        reply({
          type: "forget_folder_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    get_context: async (msg, reply) => {
      try {
        const cwd = getCurrentCwd();
        const entries = cwd ? await getContextEntries(cwd) : [];
        reply({ type: "get_context_result", ok: true, cwd, entries, request_id: msg.request_id });
      } catch (e) {
        reply({
          type: "get_context_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    set_context: async (msg, reply) => {
      try {
        const cwd = getCurrentCwd();
        if (!cwd) throw new Error("No workspace selected");
        const { saved, errors } = await setContextEntries(cwd, msg.entries || []);
        reply({
          type: "set_context_result",
          ok: errors.length === 0,
          cwd,
          saved,
          errors,
          request_id: msg.request_id,
        });
        // Restart so the agent re-reads this workspace's CLAUDE.md and picks
        // up the updated context block on the next turn.
        restartCurrentSession({ reason: "context_changed" }).catch((err) =>
          console.warn("[daemon] restart failed:", err.message),
        );
      } catch (e) {
        reply({
          type: "set_context_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
  },
});

// ---------------------------------------------------------------------------
// Office-bridge MCP server (in-process; forwards tool calls to the taskpane).
// Built per session so the registered tool family matches the connected
// host (see startSessionForFolder / the host re-narrow on hello).
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Load the user's MCP servers from ~/.claude.json — the same file the
// interactive `claude` CLI uses. The Agent SDK doesn't read this file by
// default (it reads `~/.claude/settings.json` instead), so without this step
// the user's configured servers (visio, etc.) would be invisible to the
// daemon.
// ---------------------------------------------------------------------------
async function loadUserMcpServers() {
  const configPath = join(homedir(), ".claude.json");
  try {
    const raw = await readFile(configPath, "utf8");
    const parsed = JSON.parse(raw);
    const servers = parsed?.mcpServers ?? {};
    return servers;
  } catch (err) {
    if (err.code === "ENOENT") return {};
    console.warn(`[daemon] Could not load MCP servers from ${configPath}:`, err.message);
    return {};
  }
}

const userMcpServers = await loadUserMcpServers();
const userMcpNames = Object.keys(userMcpServers);
if (userMcpNames.length > 0) {
  console.log(
    `[daemon] Loaded ${userMcpNames.length} MCP server(s) from ~/.claude.json: ${userMcpNames.join(", ")}`,
  );
}

// Preflight HTTP MCP servers. The SDK will silently drop any server whose
// initial handshake fails, with no retry for the lifetime of the session
// (see memory: sdk-silently-drops-failed-mcp). We can't fix that here, but
// we can surface the failure in the daemon log so it's obvious why a tool
// is missing — "restart the daemon when the server is back up" instead of
// "no idea why Visio doesn't work".
async function preflightHttpMcpServers(servers) {
  const entries = Object.entries(servers).filter(([, cfg]) => cfg?.type === "http" && cfg.url);
  await Promise.all(
    entries.map(async ([name, cfg]) => {
      const body = JSON.stringify({
        jsonrpc: "2.0",
        id: 1,
        method: "initialize",
        params: {
          protocolVersion: "2024-11-05",
          capabilities: {},
          clientInfo: { name: "claude-code-office-preflight", version: "0.1" },
        },
      });
      try {
        const res = await fetch(cfg.url, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Accept: "application/json, text/event-stream",
          },
          body,
          signal: AbortSignal.timeout(5000),
        });
        if (res.ok) {
          console.log(`[daemon] MCP preflight: ${name} reachable at ${cfg.url}`);
        } else {
          console.warn(
            `[daemon] MCP preflight: ${name} returned HTTP ${res.status} — tools may be missing until daemon restart`,
          );
        }
      } catch (err) {
        const reason =
          err.name === "TimeoutError" ? "timeout after 5s" : err.cause?.code || err.message;
        console.warn(
          `[daemon] MCP preflight: ${name} unreachable (${reason}) — tools will be missing until daemon restart`,
        );
      }
    }),
  );
}
await preflightHttpMcpServers(userMcpServers);

// ---------------------------------------------------------------------------
// System prompt: Claude Code default + Office-specific append.
// ---------------------------------------------------------------------------
// Re-read system-prompt.md fresh on every session start so edits take effect
// immediately when a session restarts (no daemon restart required).
async function buildSystemPromptAppend() {
  return await readFile(join(__dirname, "system-prompt.md"), "utf8");
}

// ---------------------------------------------------------------------------
// Disallow the out-of-process word-mcp tools. The user has word-mcp configured
// in ~/.claude (it surfaces as `mcp__word-mcp__*`), but those tools drive Word
// via AppleScript/COM and cause screen flicker on every edit — disqualifying
// for the live-edit path. We replace them with the in-process office_* tools
// that go through Office.js. The SDK's disallowedTools doesn't support globs,
// so we have to enumerate.
// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------
// Permission handler. Hard rule: filesystem tools must never WRITE to .docx
// files. The .docx the user has open in Word is held with unsaved changes; a
// filesystem write clobbers their work and can corrupt the file (Word holds a
// lock, OOXML cross-file references can break). Reading is allowed — the agent
// occasionally falls back to unzipping .docx XML to extract text, which is
// safe and read-only.
// ---------------------------------------------------------------------------
// Path/extension test for the Office-managed file types. We block writes to
// these via Write/Edit/MultiEdit AND via Bash commands that mention them by
// name. Not airtight (echo-redirects, base64-decoded payloads, etc. evade),
// but it defends the obvious foot-gun: `cp draft.docx active.docx`,
// `rm -rf workspace/*.docx`, `mv old.xlsx active.xlsx`.
const OFFICE_FILE_EXT = /\.(docx?|xlsx?|docm|xlsm)\b/i;

function denyWithOfficeMessage() {
  return {
    behavior: "deny",
    message:
      "Refusing to write/move/delete a Word/Excel file via filesystem tools. These files are " +
      "managed by Office and may have unsaved changes; a filesystem mutation can corrupt the " +
      "active document. Use the host's editing tools (office_* for Word, excel_* for Excel) " +
      "to change document contents.",
  };
}

function customPermissionHandler(toolName, input) {
  if (toolName === "Write" || toolName === "Edit" || toolName === "MultiEdit") {
    const path = input?.file_path ?? input?.path;
    if (typeof path === "string" && OFFICE_FILE_EXT.test(path)) {
      return Promise.resolve(denyWithOfficeMessage());
    }
  }
  if (toolName === "Bash") {
    const cmd = input?.command;
    if (typeof cmd === "string" && OFFICE_FILE_EXT.test(cmd)) {
      return Promise.resolve(denyWithOfficeMessage());
    }
  }
  // Everything else: auto-approve.
  return Promise.resolve({ behavior: "allow", updatedInput: input ?? {} });
}

const WORD_MCP_DISALLOWED = [
  "mcp__word-mcp__word_accept_revisions",
  "mcp__word-mcp__word_add_comment",
  "mcp__word-mcp__word_apply_style",
  "mcp__word-mcp__word_begin_transaction",
  "mcp__word-mcp__word_commit_transaction",
  "mcp__word-mcp__word_delete_comment",
  "mcp__word-mcp__word_delete_paragraphs",
  "mcp__word-mcp__word_delete_snapshot",
  "mcp__word-mcp__word_diff_snapshots",
  "mcp__word-mcp__word_emergency_recover",
  "mcp__word-mcp__word_export_pdf",
  "mcp__word-mcp__word_find_text",
  "mcp__word-mcp__word_get_document_info",
  "mcp__word-mcp__word_get_outline",
  "mcp__word-mcp__word_get_paragraph",
  "mcp__word-mcp__word_get_paragraphs",
  "mcp__word-mcp__word_get_section",
  "mcp__word-mcp__word_get_selection",
  "mcp__word-mcp__word_get_styles",
  "mcp__word-mcp__word_insert_paragraphs",
  "mcp__word-mcp__word_list_comments",
  "mcp__word-mcp__word_list_open_documents",
  "mcp__word-mcp__word_list_revisions",
  "mcp__word-mcp__word_list_transactions",
  "mcp__word-mcp__word_prune_snapshots",
  "mcp__word-mcp__word_reject_revisions",
  "mcp__word-mcp__word_replace_paragraphs",
  "mcp__word-mcp__word_replace_range",
  "mcp__word-mcp__word_replace_section",
  "mcp__word-mcp__word_replace_text",
  "mcp__word-mcp__word_restore_snapshot",
  "mcp__word-mcp__word_rollback_transaction",
  "mcp__word-mcp__word_toggle_track_changes",
  "mcp__word-mcp__word_undo_last_edit",
];

// ---------------------------------------------------------------------------
// Async iterable that pulls user messages from the bridge and yields them
// to the Agent SDK in the SDKUserMessage shape.
//
// We also prepend a context header to each turn so the agent always knows the
// active doc and selection without having to call office_get_doc_info first.
// ---------------------------------------------------------------------------
async function* userMessageStream() {
  while (true) {
    let msg;
    try {
      msg = await bridge.nextUserMessage();
    } catch {
      // Bridge rejected the waiter — session was aborted. Exit cleanly so
      // the underlying query() iterator can shut down without a stray error.
      return;
    }
    const { text, context } = msg;
    const header = renderContextHeader(context);
    const content = header ? `${header}\n\n${text}` : text;
    yield {
      type: "user",
      message: { role: "user", content },
      parent_tool_use_id: null,
    };
  }
}

function renderContextHeader(ctx) {
  const parts = [];
  // Host first so the agent immediately knows which tool family to use.
  // Both office_* (Word) and excel_* tools are registered simultaneously;
  // without this hint, the agent can pick the wrong family.
  if (ctx.host === "word" || ctx.host === "excel") {
    parts.push(`Host: ${ctx.host === "word" ? "Word" : "Excel"}`);
  }
  if (ctx.activeDoc) parts.push(`Doc: ${ctx.activeDoc}`);
  if (ctx.selection) {
    const s = ctx.selection;
    if (s.text) {
      const preview = s.text.length > 80 ? s.text.slice(0, 77) + "..." : s.text;
      parts.push(`Selection: "${preview}"`);
    } else if (s.para_id) {
      parts.push(`Cursor in paragraph ${s.para_id}`);
    }
  }
  // Only surface track-changes mode when it's NOT the default ("always"). The
  // system prompt says "always" by default; only deviation needs signaling.
  if (ctx.trackChangesMode && ctx.trackChangesMode !== "always") {
    parts.push(`Track changes: ${ctx.trackChangesMode}`);
  }
  return parts.length ? `[${parts.join(" · ")}]` : "";
}

// ---------------------------------------------------------------------------
// Session management. The daemon supports runtime cwd switching: when the
// user picks a different matter via the taskpane, we abort the current
// query(), clear pending user messages, and start a new query() with the new
// cwd (resuming the prior session_id for that folder if one is on record).
// Only one session is active at a time; histories live in
// ~/.claude/projects/<hash>/*.jsonl per the SDK's normal persistence.
//
// NOTE: `currentSession` is declared at the top of the module (above
// createBridge) so bridge handlers that fire during top-level awaits
// (e.g. preflightHttpMcpServers, which can block for 5s) don't hit a TDZ
// error when calling getCurrentCwd. Reassignment still happens in
// startSessionForFolder / clearCurrentSession.
// ---------------------------------------------------------------------------

function getCurrentCwd() {
  return currentSession?.cwd ?? null;
}

async function startSessionForFolder(cwd, resumeSessionId = null, { host = null } = {}) {
  // Stop any prior session before starting a new one.
  if (currentSession) {
    currentSession.abortController.abort();
    bridge.clearUserMessages();
  }

  const abortController = new AbortController();
  const session = {
    cwd,
    sessionId: resumeSessionId,
    abortController,
    settled: false,
    toolHost: host,
  };
  currentSession = session;

  // Register only the connected host's tool family (or both when no
  // taskpane has said hello yet — nothing can run at that point anyway;
  // the first hello re-narrows via maybeRenarrowForHost).
  const officeMcp = createOfficeBridgeMcp(bridge, host);

  await touchFolder(cwd);
  console.log(
    `[daemon] Starting session for ${cwd}` +
      (resumeSessionId ? ` (resuming ${resumeSessionId.slice(0, 8)}…)` : " (new session)"),
  );
  bridge.sendAssistantEvent({ event: "cwd_changed", cwd, resumed: !!resumeSessionId });
  // Structured readiness signal to the Electron shell over the IPC
  // channel — the session loop is up. Lets main.mjs flip the tray to
  // "Ready" without sniffing our stdout for a log substring. No-op when
  // run via `npm run dev` (no IPC channel).
  if (process.send) {
    try {
      process.send({ type: "daemon_ready", cwd });
    } catch {
      /* channel gone */
    }
  }
  // Replay the new workspace's transcript so the panel reflects the
  // workspace you just switched to (not the previous one's chat).
  sendTranscriptReplay().catch(() => {});

  // Re-read the drafting setup append fresh each session start.
  const append = await buildSystemPromptAppend();

  // Did this turn see a proper `result` message before the stream ended?
  // The SDK ends the stream with a `result` on normal completion. On a
  // usage-limit / quota hit (and some transport failures) the stream just
  // ends with no result and no thrown error — leaving the taskpane pinned
  // to "Working…". We track this to recover.
  let sawResult = false;
  // Best-effort usage-limit detection from the SDK CLI's stderr. The exact
  // phrasing varies by SDK version and limit kind (per-minute / daily /
  // weekly); match broadly.
  let rateLimitHint = null;
  const RATE_LIMIT_RE =
    /(usage limit|rate limit|daily limit|weekly limit|quota|too many requests|429|limit reached|limit will reset|resets? at|upgrade to|out of (?:credits|quota))/i;

  // Fire-and-forget; index.mjs keeps running while the agent loop iterates.
  (async () => {
    try {
      for await (const msg of query({
        prompt: userMessageStream(),
        options: {
          cwd,
          systemPrompt: {
            type: "preset",
            preset: "claude_code",
            append,
          },
          mcpServers: { ...userMcpServers, office: officeMcp },
          disallowedTools: WORD_MCP_DISALLOWED,
          canUseTool: customPermissionHandler,
          includePartialMessages: true,
          abortController,
          // Surface the SDK CLI's stderr (MCP connect failures, internal
          // warnings, etc.) in our daemon log. Also sniff it for
          // usage-limit phrasing so we can show the user a clear message
          // instead of a silently stuck "Working…".
          stderr: (data) => {
            for (const line of String(data).split(/\r?\n/)) {
              if (!line.trim()) continue;
              console.error(`[sdk] ${line}`);
              if (!rateLimitHint && RATE_LIMIT_RE.test(line)) rateLimitHint = line.trim();
            }
          },
          ...(resumeSessionId ? { resume: resumeSessionId } : {}),
        },
      })) {
        if (currentSession !== session) break; // superseded by a later switch
        if (msg.type === "result") sawResult = true;
        handleAgentMessage(msg, session);
      }
      // Loop ended normally. If we never saw a `result`, the session is
      // still the live one, and the user didn't Stop, the stream died
      // unexpectedly — almost always a usage-limit hit. Tell the taskpane
      // (so it leaves "Working…") and auto-restart the loop so the next
      // message has a live consumer (same rationale as the Stop path).
      if (currentSession === session && !sawResult && !session.interrupted) {
        const friendly = rateLimitHint
          ? `Claude usage limit reached. ${rateLimitHint}`
          : "The agent stopped unexpectedly — this is usually a Claude usage limit. Wait for your limit to reset, or set ANTHROPIC_API_KEY to use an API key.";
        bridge.sendAssistantEvent({ event: "error", error: friendly });
        bridge.sendAssistantEvent({ event: "turn_complete", subtype: "stream_ended" });
        const { cwd: rcwd, sessionId: rsid, toolHost: rhost } = session;
        setImmediate(() => {
          startSessionForFolder(rcwd, rsid, { host: rhost ?? null }).catch((restartErr) =>
            console.error("[daemon] post-stream-end session restart failed:", restartErr.message),
          );
        });
      }
    } catch (err) {
      if (err.name === "AbortError" || /aborted/i.test(err.message ?? "")) {
        // Expected when switching sessions OR when the user clicked Stop.
        // In the stop case (session.interrupted = true) we still need to
        // flip the taskpane's agent status back to Ready (otherwise it
        // stays pinned to "Working…") AND auto-restart a fresh resuming
        // loop — without a live query() iterator awaiting
        // bridge.nextUserMessage(), the user's next message would enqueue
        // with nobody to consume it. Let the finally block null
        // currentSession first (via setImmediate); the fresh
        // startSessionForFolder then builds cleanly.
        if (currentSession === session && session.interrupted) {
          bridge.sendAssistantEvent({ event: "turn_complete", interrupted: true });
          const { cwd: rcwd, sessionId: rsid, toolHost: rhost } = session;
          setImmediate(() => {
            startSessionForFolder(rcwd, rsid, { host: rhost ?? null }).catch((restartErr) =>
              console.error("[daemon] post-stop session restart failed:", restartErr.message),
            );
          });
        }
      } else if (currentSession === session) {
        console.error("[daemon] Agent loop crashed:", err);
        // Detect auth failures and surface them as a distinct event so the
        // taskpane can show a recoverable banner ("sign in to Claude Code")
        // instead of just dumping the SDK's raw error. Matched generously:
        // SDK error messages have varied across versions.
        const msgText = String(err?.message ?? err);
        const isAuth =
          /\b(authentication|unauthorized|credential|api[- ]?key|sign[- ]?in|401)\b/i.test(
            msgText,
          ) || /OAUTH/i.test(msgText);
        if (isAuth) {
          bridge.sendAssistantEvent({ event: "auth_error", error: msgText });
        } else {
          bridge.sendAssistantEvent({ event: "error", error: msgText });
        }
      }
    } finally {
      session.settled = true;
      if (currentSession === session) currentSession = null;
    }
  })();

  return session;
}

async function switchFolder(rawCwd) {
  const cwd = resolve(rawCwd);
  // Validate the path is a directory.
  const s = await stat(cwd);
  if (!s.isDirectory()) throw new Error(`Not a directory: ${cwd}`);
  const saved = await getSessionForFolder(cwd);
  // Preserve the connected host's tool family across the workspace switch.
  const host = bridge.getContext().host ?? null;
  await startSessionForFolder(cwd, saved?.session_id ?? null, { host });
  return cwd;
}

// Re-launch the current session (same cwd, resuming via session_id) so that
// changes to CLAUDE.md, the drafting setup, or other config that's loaded at
// session-init take effect without losing conversation history.
async function restartCurrentSession({ reason = "config_changed" } = {}) {
  if (!currentSession) return;
  const cwd = currentSession.cwd;
  const sessionId = currentSession.sessionId;
  const host = currentSession.toolHost ?? null;
  console.log(`[daemon] Restarting session for ${cwd} (reason: ${reason})`);
  bridge.sendAssistantEvent({ event: "config_reloaded", reason });
  await startSessionForFolder(cwd, sessionId, { host });
}

function handleAgentMessage(msg, session = currentSession) {
  switch (msg.type) {
    case "system": {
      if (msg.subtype === "init") {
        console.log(`[agent] init session ${msg.session_id} (model: ${msg.model})`);
        bridge.sendAssistantEvent({
          event: "session_init",
          session_id: msg.session_id,
          model: msg.model,
        });
        // Record this session_id for the current cwd so the next switch back
        // resumes here.
        if (session && msg.session_id && msg.session_id !== session.sessionId) {
          session.sessionId = msg.session_id;
          saveSessionForFolder(session.cwd, msg.session_id).catch((err) =>
            console.warn("[daemon] Could not save session id:", err.message),
          );
        }
      }
      break;
    }
    case "stream_event": {
      // includePartialMessages stream events. Forward text_delta to the
      // taskpane as assistant_text. Other event types (content_block_start,
      // content_block_stop, message_start/stop) we currently ignore.
      const delta = msg.event?.delta;
      if (delta?.type === "text_delta" && typeof delta.text === "string" && delta.text.length > 0) {
        bridge.sendAssistantText(delta.text);
      }
      break;
    }
    case "assistant": {
      // The complete assistant message arrives after streaming. We skip
      // text blocks (already streamed as deltas) and only forward tool_use
      // announces — those are atomic and not streamed.
      const blocks = msg.message?.content ?? [];
      for (const block of blocks) {
        if (block.type === "tool_use") {
          bridge.sendAssistantEvent({
            event: "tool_use_announce",
            tool: block.name,
            input: block.input,
          });
        }
      }
      break;
    }
    case "result": {
      console.log(`[agent] turn complete (${msg.subtype})`);
      bridge.sendAssistantEvent({ event: "turn_complete", subtype: msg.subtype });
      break;
    }
    case "user": {
      // tool_result messages — we don't forward them; the bridge handles them.
      break;
    }
    default:
      // Other event types (api_retry, hook events, etc.) — ignore for POC.
      break;
  }
}

// ---------------------------------------------------------------------------
// Kick off.
// ---------------------------------------------------------------------------
startSessionForFolder(matterFolder).catch((err) => {
  console.error("[daemon] Failed to start initial session:", err);
  process.exit(1);
});

// Keep process alive even when nothing is happening.
process.on("SIGINT", () => {
  console.log("\n[daemon] Shutting down");
  process.exit(0);
});
