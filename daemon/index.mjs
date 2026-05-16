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
  getSessionId,
  saveSessionId,
  touchFolder,
  getRecentFolders,
  forgetFolder,
} from "./sessions.mjs";
import { readTranscript } from "./transcript.mjs";
import { diag } from "./diag.mjs";
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

// Word and Excel are fully independent surfaces: each host gets its own
// agent loop, session, message queue, transcript AND workspace. Nothing
// one host does ever touches the other (no shared "current session", no
// cross-host interrupt). Both maps are hoisted above createBridge so
// bridge handlers that fire during the top-level awaits below
// (preflightHttpMcpServers, etc.) don't hit a TDZ on these bindings.
//
//   sessions:        host -> live session { cwd, sessionId, host,
//                    abortController, settled, turnActive, interrupted }
//   workspaceByHost: host -> last-known cwd, so a pane that connects
//                    before its first message still resolves a workspace
//                    (and survives that host's loop ending).
const sessions = new Map();
const workspaceByHost = new Map();

function sessionFor(host) {
  return sessions.get(host) ?? null;
}

// This host's workspace: its live session's cwd, else the last cwd we
// recorded for it, else the launch default.
function cwdForHost(host) {
  return sessionFor(host)?.cwd ?? workspaceByHost.get(host) ?? matterFolder;
}

// Resolve the session id to replay for (host, cwd). Prefer this host's
// live session id (set once the SDK reports init); otherwise the id
// persisted for (host, cwd) — covers the window before the SDK has
// re-inited, and keeps Word's and Excel's transcripts separate.
async function resolveReplaySessionId(host, cwd) {
  const live = sessionFor(host);
  if (live?.sessionId) return live.sessionId;
  if (!host || !cwd) return null;
  try {
    return (await getSessionId(host, cwd)) ?? null;
  } catch {
    return null;
  }
}

// Reconstruct (host, cwd)'s prior conversation from its .jsonl and push
// it to THAT host's pane only. Sent on every taskpane hello and after a
// workspace switch (cwd_changed). Empty events => fresh chat, no divider.
async function sendTranscriptReplayTo(host, cwd) {
  try {
    const sessionId = await resolveReplaySessionId(host, cwd);
    const { events, truncated } = sessionId
      ? await readTranscript(sessionId, { maxEvents: 200 })
      : { events: [], truncated: false };
    bridge.sendToTaskpane(
      {
        type: "transcript_replay",
        session_id: sessionId ?? null,
        truncated,
        events,
      },
      host,
    );
  } catch (err) {
    console.warn("[daemon] transcript replay failed:", err?.message ?? err);
  }
}

// Start (or resume) this host's session on the next tick — past the
// current message handler / the agent loop's finally block (which clears
// sessions.get(host)), so startSessionForFolder builds cleanly. Every
// path that (re)starts a host's session — first message, post-stream-end,
// post-Stop — funnels through here; `reason` only flavors the failure log.
function scheduleSessionStart(cwd, sessionId, host, reason, { replay = true } = {}) {
  setImmediate(() => {
    startSessionForFolder(cwd, sessionId, { host: host ?? null, replay }).catch((err) =>
      console.error(`[daemon] ${reason} session start failed:`, err?.message ?? err),
    );
  });
}

// Called on every taskpane hello — for EITHER host, both panes possibly
// connected at once. Connecting a pane must NOT start the agent loop:
// connect-driven starts amplified the old connect/disconnect ping-pong
// and burn an unasked turn. We only re-render this pane's own
// (host, cwd) transcript. The loop starts lazily on the first user
// message (onUserMessage → ensureLoopForMessage).
async function onPaneConnect(host) {
  if (!host) return;
  const cwd = cwdForHost(host);
  diag(`hello → replay host=${host} cwd=${cwd} (no loop start on connect)`);
  await sendTranscriptReplayTo(host, cwd);
}

// Called when a user message arrives from `host`, BEFORE it's queued.
// Each host has its OWN loop — independent of the other host entirely.
// If this host's loop is already live, do nothing (its userMessageStream
// will consume the message). Otherwise start it, resuming this host's
// (host, cwd) conversation. Deferred via setImmediate so it lands after
// the message is queued and after any in-flight finally; the new loop
// then drains this host's queue. The OTHER host's loop is never touched.
async function ensureLoopForMessage(host) {
  if (!host) return;
  const cwd = cwdForHost(host);
  const live = sessionFor(host);
  if (live && !live.settled) {
    return; // this host's loop is live and will consume the message
  }
  let resumeId = null;
  try {
    resumeId = await getSessionId(host, cwd);
  } catch {
    /* fresh session if lookup fails */
  }
  diag(`message → ensure loop host=${host} cwd=${cwd} resume=${resumeId ?? "(new)"}`);
  // replay:false — the pane already shows the chat (incl. the message
  // that just triggered this). An empty transcript_replay here (a
  // brand-new session has no .jsonl yet) would wipe the user's prompt.
  scheduleSessionStart(cwd, resumeId, host, "user message", { replay: false });
}

// ---------------------------------------------------------------------------
// WebSocket bridge.
// ---------------------------------------------------------------------------
const bridge = createBridge({
  port: WS_PORT,
  token: BRIDGE_TOKEN,
  allowedOrigins: [HTTP_ORIGIN],
  onHello: (host) => onPaneConnect(host),
  onUserMessage: (host) => ensureLoopForMessage(host),
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
    set_cwd: async (msg, reply, host) => {
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
        const resolved = await switchFolder(cwd, host);
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
    get_cwd_state: async (msg, reply, host) => {
      const recent = await getRecentFolders();
      reply({
        type: "get_cwd_state_result",
        ok: true,
        // This host's own workspace, resolvable even before its first
        // message (lazy start ⇒ no session yet).
        current_cwd: cwdForHost(host),
        recent,
        request_id: msg.request_id,
      });
    },
    stop_agent: async (msg, reply, host) => {
      // User clicked Stop in this pane. Abort only THIS host's loop; the
      // query() iterator's catch path sees AbortError and (because we
      // flag the session interrupted) emits turn_complete interrupted so
      // the taskpane flips to Ready and auto-restarts a resuming loop.
      // The other host's loop is untouched.
      const s = sessionFor(host);
      if (s) {
        s.interrupted = true;
        s.abortController.abort();
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
    get_context: async (msg, reply, host) => {
      try {
        const cwd = cwdForHost(host);
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
    set_context: async (msg, reply, host) => {
      try {
        const cwd = cwdForHost(host);
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
        // Restart THIS host's loop so its agent re-reads CLAUDE.md and
        // picks up the updated context block on the next turn. The other
        // host is unaffected.
        restartSession(host, { reason: "context_changed" }).catch((err) =>
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

// Hoisted ABOVE the top-level awaits below on purpose. Module evaluation
// suspends at the first top-level `await` (loadUserMcpServers /
// preflightHttpMcpServers). The bridge WS server is already listening by
// then, so a taskpane `hello` can arrive mid-suspension and drive
// startSessionForFolder before the rest of the module body runs. Anything
// that path touches must be initialized first, or it hits a TDZ
// ReferenceError. Keep this (and any other start-path consts) up here.
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

// ---------------------------------------------------------------------------
// Async iterable that pulls user messages from the bridge and yields them
// to the Agent SDK in the SDKUserMessage shape.
//
// We also prepend a context header to each turn so the agent always knows the
// active doc and selection without having to call office_get_doc_info first.
// ---------------------------------------------------------------------------
async function* userMessageStream(host, session) {
  while (true) {
    let msg;
    try {
      msg = await bridge.nextUserMessage(host);
    } catch {
      // Bridge rejected the waiter — session was aborted. Exit cleanly so
      // the underlying query() iterator can shut down without a stray error.
      return;
    }
    // A turn is now in flight for this session. Cleared when we emit a
    // turn_complete (result / stream-ended / stop / error). Lets a
    // cross-host supersede know it must release the outgoing pane.
    if (session) session.turnActive = true;
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
// Session management. Each host runs its own independent query() loop with
// its own cwd; switching workspace or reloading config restarts only that
// host's loop. Histories live in ~/.claude/projects/<hash>/*.jsonl per the
// SDK's normal persistence. `sessions` / `workspaceByHost` are declared at
// the top of the module (above createBridge) so bridge handlers that fire
// during top-level awaits (e.g. preflightHttpMcpServers, which can block
// for 5s) don't hit a TDZ on those bindings.
// ---------------------------------------------------------------------------

async function startSessionForFolder(
  cwd,
  resumeSessionId = null,
  { host = null, replay = true } = {},
) {
  // Replace only THIS host's prior loop (same-host restart: workspace
  // switch, config reload, post-Stop/stream-end resume). The other
  // host's loop is never involved — no cross-host abort, no interrupt.
  // Clear only this host's queue so the other host's queued message is
  // untouched.
  const prior = sessionFor(host);
  if (prior) {
    prior.abortController.abort();
  }
  bridge.clearUserMessages(host);

  const abortController = new AbortController();
  const session = {
    cwd,
    sessionId: resumeSessionId,
    abortController,
    settled: false,
    turnActive: false,
    host,
  };
  sessions.set(host, session);
  workspaceByHost.set(host, cwd);

  // Register only this host's tool family. A session is created lazily
  // on the first user message (onUserMessage → ensureLoopForMessage), so
  // `host` is normally "word" or "excel"; it's only null in the degraded
  // case where a set_cwd arrived before any pane bound a host (both
  // families registered).
  const officeMcp = createOfficeBridgeMcp(bridge, host);

  await touchFolder(cwd);
  console.log(
    `[daemon] Starting session for ${cwd}` +
      (resumeSessionId ? ` (resuming ${resumeSessionId.slice(0, 8)}…)` : " (new session)"),
  );
  bridge.sendAssistantEvent({ event: "cwd_changed", cwd, resumed: !!resumeSessionId }, host);
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
  // Replay this host's transcript for the new workspace so its panel
  // reflects the workspace you just switched to (not the previous chat).
  // Skipped when this start was triggered by the user's own message
  // (replay:false): the pane already shows that message, and a fresh
  // session's empty replay would erase it.
  if (replay) sendTranscriptReplayTo(host, cwd).catch(() => {});

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
        prompt: userMessageStream(host, session),
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
        if (sessionFor(host) !== session) break; // this host's loop was restarted
        if (msg.type === "result") sawResult = true;
        handleAgentMessage(msg, session);
      }
      // Loop ended normally. If we never saw a `result`, the session is
      // still the live one, and the user didn't Stop, the stream died
      // unexpectedly — almost always a usage-limit hit. Tell the taskpane
      // (so it leaves "Working…") and auto-restart the loop so the next
      // message has a live consumer (same rationale as the Stop path).
      if (sessionFor(host) === session && !sawResult && !session.interrupted) {
        session.turnActive = false;
        const friendly = rateLimitHint
          ? `Claude usage limit reached. ${rateLimitHint}`
          : "The agent stopped unexpectedly — this is usually a Claude usage limit. Wait for your limit to reset, or set ANTHROPIC_API_KEY to use an API key.";
        bridge.sendAssistantEvent({ event: "error", error: friendly }, host);
        bridge.sendAssistantEvent({ event: "turn_complete", subtype: "stream_ended" }, host);
        const { cwd: rcwd, sessionId: rsid, host: rhost } = session;
        scheduleSessionStart(rcwd, rsid, rhost, "post-stream-end");
      }
    } catch (err) {
      if (err.name === "AbortError" || /aborted/i.test(err.message ?? "")) {
        // Expected on a same-host restart OR when the user clicked Stop.
        // In the stop case (session.interrupted = true) we still need to
        // flip the taskpane's agent status back to Ready (otherwise it
        // stays pinned to "Working…") AND auto-restart a fresh resuming
        // loop — without a live query() iterator awaiting
        // bridge.nextUserMessage(), the user's next message would enqueue
        // with nobody to consume it. Let the finally block clear this
        // host's session first (via setImmediate); the fresh
        // startSessionForFolder then builds cleanly.
        if (sessionFor(host) === session && session.interrupted) {
          session.turnActive = false;
          bridge.sendAssistantEvent({ event: "turn_complete", interrupted: true }, host);
          const { cwd: rcwd, sessionId: rsid, host: rhost } = session;
          scheduleSessionStart(rcwd, rsid, rhost, "post-stop");
        }
      } else if (sessionFor(host) === session) {
        session.turnActive = false;
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
          bridge.sendAssistantEvent({ event: "auth_error", error: msgText }, host);
        } else {
          bridge.sendAssistantEvent({ event: "error", error: msgText }, host);
        }
      }
    } finally {
      session.settled = true;
      if (sessionFor(host) === session) sessions.delete(host);
    }
  })();

  return session;
}

async function switchFolder(rawCwd, host = null) {
  const cwd = resolve(rawCwd);
  // Validate the path is a directory.
  const s = await stat(cwd);
  if (!s.isDirectory()) throw new Error(`Not a directory: ${cwd}`);
  // Switch ONLY the requesting pane's host to the target folder; resume
  // that (host, cwd)'s prior conversation if one is on record. The other
  // host stays in its own workspace, untouched.
  const resumeId = host ? await getSessionId(host, cwd) : null;
  await startSessionForFolder(cwd, resumeId, { host });
  return cwd;
}

// Re-launch one host's loop (same cwd, resuming via session_id) so that
// changes to CLAUDE.md, the drafting setup, or other config loaded at
// session-init take effect without losing conversation history. No-op if
// that host has no live loop.
async function restartSession(host, { reason = "config_changed" } = {}) {
  const s = sessionFor(host);
  if (!s) return;
  const { cwd, sessionId } = s;
  console.log(`[daemon] Restarting ${host} session for ${cwd} (reason: ${reason})`);
  bridge.sendAssistantEvent({ event: "config_reloaded", reason }, host);
  await startSessionForFolder(cwd, sessionId, { host });
}

function handleAgentMessage(msg, session) {
  switch (msg.type) {
    case "system": {
      if (msg.subtype === "init") {
        console.log(`[agent] init session ${msg.session_id} (model: ${msg.model})`);
        try {
          const tn = Array.isArray(msg.tools) ? msg.tools : Object.keys(msg.tools ?? {});
          diag(
            `init tools (${tn.length}):`,
            tn.filter((t) => /office|excel|mcp__/.test(String(t))).join(", ") ||
              "(no office/excel/mcp tools in init list!)",
          );
        } catch (e) {
          diag("init tools introspection failed:", e?.message);
        }
        bridge.sendAssistantEvent(
          {
            event: "session_init",
            session_id: msg.session_id,
            model: msg.model,
          },
          session?.host,
        );
        // Record this session_id for THIS (host, cwd) so the next time
        // this pane connects (or you switch back) it resumes here.
        if (session && msg.session_id && msg.session_id !== session.sessionId) {
          session.sessionId = msg.session_id;
          saveSessionId(session.host, session.cwd, msg.session_id).catch((err) =>
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
        bridge.sendAssistantText(delta.text, session?.host);
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
          diag("model called tool:", block.name);
          bridge.sendAssistantEvent(
            {
              event: "tool_use_announce",
              tool: block.name,
              input: block.input,
            },
            session?.host,
          );
        }
      }
      break;
    }
    case "result": {
      console.log(`[agent] turn complete (${msg.subtype})`);
      if (session) session.turnActive = false;
      bridge.sendAssistantEvent({ event: "turn_complete", subtype: msg.subtype }, session?.host);
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
//
// No eager session: an agent session needs a host, and the host is only
// known once a taskpane connects and says hello (→ ensureSessionFor
// ActivePane). Starting both-tool sessions pre-hello is exactly what
// caused the per-host tool/transcript churn. The bridge buffers any
// user_message until a session's loop consumes it, so nothing is lost.
//
// Tell the Electron shell we're up now (servers listening) so the tray
// flips to "Ready" without waiting for a pane — independent of, and
// idempotent with, the daemon_ready that startSessionForFolder also
// emits on the first real session.
// ---------------------------------------------------------------------------
console.log(`[daemon] Ready; waiting for a taskpane. Default workspace: ${matterFolder}`);
if (process.send) {
  try {
    process.send({ type: "daemon_ready" });
  } catch {
    /* no IPC channel (npm run dev) */
  }
}

// Keep process alive even when nothing is happening.
process.on("SIGINT", () => {
  console.log("\n[daemon] Shutting down");
  process.exit(0);
});
