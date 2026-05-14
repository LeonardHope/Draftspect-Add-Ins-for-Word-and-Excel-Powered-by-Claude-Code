/* global Office, Word */

// Daemon endpoints. The HTTP server that loaded this taskpane is on
// HTTP_PORT; the WebSocket bridge is on WS_PORT (one less by daemon convention).
const WS_URL = "ws://127.0.0.1:47823";
const TOKEN_URL = "/bridge-token";

// The bridge token — fetched at boot from the same-origin HTTP server. The
// bridge rejects any WS that doesn't present this token in its first hello.
let bridgeToken = null;

async function fetchBridgeToken() {
  try {
    const res = await fetch(TOKEN_URL, { credentials: "same-origin" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    bridgeToken = (await res.text()).trim();
  } catch (e) {
    console.warn("Failed to fetch bridge token:", e);
    bridgeToken = null;
  }
}

let ws = null;
let wsReady = false;
let lastSelection = null;
let activeDocUrl = null;

// ---------------------------------------------------------------------------
// Paragraph addressing
//
// We use Word's `paragraph.uniqueLocalId` (the OOXML w:paraId) — stable
// across insertions/deletions, which is what the agent needs when it does
// "read these paragraphs, then edit p7" across multiple tool calls. Available
// in WordApi 1.6+ (every recent Microsoft 365 Word).
//
// On older Word builds where uniqueLocalId is unsupported, we fall back to
// position-based IDs (`p${index}`). These are NOT stable across structural
// edits — when the fallback is in play, snapshotParagraphs() flags the mode
// and tools should re-read before each operation rather than caching IDs.
// ---------------------------------------------------------------------------
const fallbackId = (index) => `p${index}`;
const parseFallbackId = (id) => {
  const m = /^p(\d+)$/.exec(id);
  return m ? parseInt(m[1], 10) : null;
};

// Load paragraphs with text + style + uniqueLocalId (when supported). The
// returned object includes an `idMode`: "uniqueLocalId" (stable) or "index"
// (fallback). Callers use getId() / findIndexById() which respect the mode.
async function getParagraphsWithIds(context) {
  const paragraphs = context.document.body.paragraphs;
  try {
    paragraphs.load("items/text, items/style, items/uniqueLocalId");
    await context.sync();
    // It's possible for load() to succeed on a build that doesn't actually
    // populate uniqueLocalId — guard against that by checking the first item.
    if (paragraphs.items.length > 0 && !paragraphs.items[0].uniqueLocalId) {
      throw new Error("uniqueLocalId not populated");
    }
    return { paragraphs, idMode: "uniqueLocalId" };
  } catch {
    // Older Word: load without uniqueLocalId and fall back to index IDs.
    paragraphs.load("items/text, items/style");
    await context.sync();
    return { paragraphs, idMode: "index" };
  }
}

function getId(paragraph, index, idMode) {
  return idMode === "uniqueLocalId" ? paragraph.uniqueLocalId : fallbackId(index);
}

function findIndexById(paragraphs, id, idMode) {
  if (idMode === "uniqueLocalId") {
    return paragraphs.items.findIndex(p => p.uniqueLocalId === id);
  }
  const i = parseFallbackId(id);
  if (i === null || i < 0 || i >= paragraphs.items.length) return -1;
  return i;
}

// Run `body` with the document's changeTrackingMode set to TrackAll if
// `track_changes` is truthy, and unconditionally restore it in `finally` so
// that exceptions thrown mid-edit don't leave Word stuck in Track All mode.
// The restore-sync is wrapped in a try/catch so a failed restore doesn't
// shadow the original error.
async function withTrackChanges(context, track_changes, body) {
  let prevTracking = null;
  if (track_changes) {
    context.document.load("changeTrackingMode");
    await context.sync();
    prevTracking = context.document.changeTrackingMode;
    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
  }
  try {
    return await body();
  } finally {
    if (track_changes && prevTracking !== null) {
      try {
        context.document.changeTrackingMode = prevTracking;
        await context.sync();
      } catch { /* swallow — don't shadow original error */ }
    }
  }
}

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------
const $messages = document.getElementById("messages");
const $status = document.getElementById("status");
const $activeDoc = document.getElementById("active-doc");
const $input = document.getElementById("input");
const $send = document.getElementById("send");
const $composer = document.getElementById("composer");
const $chip = document.getElementById("selection-chip");
const $chipText = document.getElementById("selection-chip-text");
const $chipDetach = document.getElementById("selection-chip-detach");

let assistantTurnElem = null;
let attachSelection = true;

function setStatus(state, label) {
  $status.className = `status ${state}`;
  $status.textContent = label;
}

function appendUserMessage(text) {
  const el = document.createElement("div");
  el.className = "msg user";
  el.textContent = text;
  $messages.appendChild(el);
  $messages.scrollTop = $messages.scrollHeight;
  assistantTurnElem = null;
}

function appendAssistantDelta(delta) {
  if (!assistantTurnElem) {
    assistantTurnElem = document.createElement("div");
    assistantTurnElem.className = "msg assistant";
    $messages.appendChild(assistantTurnElem);
  }
  assistantTurnElem.textContent += delta;
  $messages.scrollTop = $messages.scrollHeight;
}

function appendEvent(text) {
  const el = document.createElement("div");
  el.className = "msg event";
  el.textContent = text;
  $messages.appendChild(el);
  $messages.scrollTop = $messages.scrollHeight;
  assistantTurnElem = null;
}

function appendToolUse(name, args) {
  const el = document.createElement("div");
  el.className = "msg tool";
  el.innerHTML = `<div class="tool-name"></div><div class="tool-args"></div>`;
  el.querySelector(".tool-name").textContent = `🔧 ${name}`;
  const argText = typeof args === "string" ? args : JSON.stringify(args, null, 2);
  el.querySelector(".tool-args").textContent = argText.length > 200 ? argText.slice(0, 197) + "..." : argText;
  $messages.appendChild(el);
  $messages.scrollTop = $messages.scrollHeight;
  assistantTurnElem = null;
}

function refreshSelectionChip() {
  if (attachSelection && lastSelection && lastSelection.text) {
    const preview = lastSelection.text.length > 60
      ? lastSelection.text.slice(0, 57) + "..."
      : lastSelection.text;
    $chipText.textContent = `Selection: "${preview}"`;
    $chip.hidden = false;
  } else {
    $chip.hidden = true;
  }
}

$chipDetach.addEventListener("click", () => {
  attachSelection = false;
  refreshSelectionChip();
});

// ---------------------------------------------------------------------------
// Settings — persisted to localStorage. New settings get added here and
// applied via applySettings().
// ---------------------------------------------------------------------------
const SETTINGS_KEY = "cc-word-addin-settings-v1";

function defaultSettings() {
  return {
    showDiagnostics: true,
    autoSwitchWorkspace: false,
    trackChangesMode: "always", // "always" | "modifications" | "never"
  };
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (!raw) return defaultSettings();
    return { ...defaultSettings(), ...JSON.parse(raw) };
  } catch {
    return defaultSettings();
  }
}

function saveSettings(s) {
  try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(s)); } catch { /* ignore */ }
}

let settings = loadSettings();

function applySettings() {
  $messages.dataset.showDiagnostics = String(settings.showDiagnostics);
  const $showDiag = document.getElementById("setting-show-diagnostics");
  if ($showDiag) $showDiag.checked = settings.showDiagnostics;
  const $autoSwitch = document.getElementById("setting-auto-switch-workspace");
  if ($autoSwitch) $autoSwitch.checked = settings.autoSwitchWorkspace;
  const $tcMode = document.getElementById("setting-track-changes-mode");
  if ($tcMode) $tcMode.value = settings.trackChangesMode || "always";
}

document.getElementById("setting-show-diagnostics").addEventListener("change", (e) => {
  settings.showDiagnostics = e.target.checked;
  saveSettings(settings);
  applySettings();
});

document.getElementById("setting-track-changes-mode").addEventListener("change", (e) => {
  settings.trackChangesMode = e.target.value;
  saveSettings(settings);
  applySettings();
  // Push to daemon so the agent sees the new mode on the next turn.
  if (wsReady) {
    wsSend({ type: "context_update", track_changes_mode: settings.trackChangesMode });
  }
});

document.getElementById("setting-auto-switch-workspace").addEventListener("change", (e) => {
  settings.autoSwitchWorkspace = e.target.checked;
  saveSettings(settings);
  applySettings();
  // Re-evaluate the suggest banner — turning on auto-switch may now silently
  // fire a marker-confidence switch; turning it off should leave the banner
  // visible if there's still a mismatch.
  refreshWorkspaceSuggestBanner();
});

applySettings();

// ---------------------------------------------------------------------------
// WebSocket
// ---------------------------------------------------------------------------
function wsConnect() {
  setStatus("idle", "Connecting…");
  ws = new WebSocket(WS_URL);

  ws.onopen = () => {
    wsReady = true;
    sendHello();
  };

  ws.onclose = () => {
    wsReady = false;
    setStatus("err", "Disconnected — retrying…");
    // Daemon may have been restarted, in which case it has rotated the
    // bridge token. Re-fetch from /bridge-token before each reconnect
    // attempt so the next hello carries the current token. fetchBridgeToken
    // tolerates the HTTP server being briefly unreachable too (sets the
    // token to null, the hello fails, we loop again).
    setTimeout(() => { fetchBridgeToken().then(wsConnect); }, 1500);
  };

  ws.onerror = (err) => {
    console.warn("WS error:", err);
  };

  ws.onmessage = (evt) => {
    let msg;
    try { msg = JSON.parse(evt.data); } catch { return; }
    handleServerMessage(msg);
  };
}

function wsSend(obj) {
  if (!wsReady) return;
  ws.send(JSON.stringify(obj));
}

function sendHello() {
  wsSend({
    type: "hello",
    token: bridgeToken,
    active_doc: activeDocUrl,
    selection: attachSelection ? lastSelection : null,
    track_changes_mode: settings.trackChangesMode || "always",
  });
}

async function handleServerMessage(msg) {
  switch (msg.type) {
    case "welcome":
      setStatus("ok", "Ready");
      refreshWorkspaceFromDaemon();
      break;

    case "assistant_text":
      appendAssistantDelta(msg.delta);
      break;

    case "assistant_event":
      if (msg.event === "tool_use_announce") {
        appendToolUse(msg.tool, msg.input);
      } else if (msg.event === "turn_complete") {
        if (wsReady) setStatus("ok", "Ready");
      } else if (msg.event === "error") {
        appendEvent(`Error: ${msg.error}`);
        if (wsReady) setStatus("ok", "Ready");
      } else if (msg.event === "session_init") {
        appendEvent(`Session ${msg.session_id?.slice(0, 8)}… (${msg.model})`);
      } else if (msg.event === "cwd_changed") {
        setWorkspaceDisplay(msg.cwd);
        appendEvent(`Switched to workspace: ${msg.cwd.split(/[\\/]/).filter(Boolean).pop()}${msg.resumed ? " (resumed prior session)" : ""}`);
        // Per-workspace context must be re-read for the new workspace.
        contextCache = null;
        if (document.body.dataset.activeTab === "setup") loadContext(true);
      } else if (msg.event === "config_reloaded") {
        const what = msg.reason === "context_changed" ? "context files" : "config";
        appendEvent(`Session reloaded — ${what} updated.`);
      }
      break;

    case "tool_call":
      runOfficeTool(msg);
      break;

    case "pong":
      break;

    default:
      // Request/response messages keyed by request_id end in "_result".
      // Resolve the matching pending request.
      if (typeof msg.type === "string" && msg.type.endsWith("_result") && msg.request_id) {
        const pending = pendingRequests.get(msg.request_id);
        if (pending) {
          pendingRequests.delete(msg.request_id);
          pending.resolve(msg);
        }
      }
      break;
  }
}

// ---- Request/response helper (for non-tool round-trips) -------------------
const pendingRequests = new Map();
const REQUEST_TIMEOUT_MS = 10_000;

function sendRequest(type, payload = {}) {
  return new Promise((resolve, reject) => {
    if (!wsReady) { reject(new Error("Not connected to daemon")); return; }
    const request_id = uuid();
    pendingRequests.set(request_id, { resolve, reject });
    wsSend({ type, request_id, ...payload });
    setTimeout(() => {
      if (pendingRequests.has(request_id)) {
        pendingRequests.delete(request_id);
        reject(new Error(`Request "${type}" timed out`));
      }
    }, REQUEST_TIMEOUT_MS);
  });
}

// ---------------------------------------------------------------------------
// Tool dispatcher
// ---------------------------------------------------------------------------
// Apply the user's track-changes mode setting on top of whatever the agent
// passed. The agent's schema-default value can't be trusted as the final
// arbiter: the user's mode setting is the policy.
//   - "always": force true regardless of what the agent passed.
//   - "never":  force false regardless.
//   - "modifications": honor the agent's flag — the system prompt tells the
//                      agent to pass false only for fresh-section drafts.
function effectiveTrackChanges(provided) {
  const mode = (typeof settings !== "undefined" && settings?.trackChangesMode) || "always";
  if (mode === "never") return false;
  if (mode === "always") return true;
  return provided !== false;
}

async function runOfficeTool(msg) {
  const { id, name, args } = msg;
  try {
    let result;
    // Apply the user's track-changes mode to every write-tool call before
    // dispatching, so user preference always wins over the agent's value.
    const writeArgs = () => ({
      ...args,
      track_changes: effectiveTrackChanges(args && args.track_changes),
    });
    switch (name) {
      case "office_get_selection":
        result = await toolGetSelection();
        break;
      case "office_read_paragraphs":
        result = await toolReadParagraphs(args);
        break;
      case "office_insert_paragraphs":
        result = await toolInsertParagraphs(writeArgs());
        break;
      case "office_replace_paragraphs":
        result = await toolReplaceParagraphs(writeArgs());
        break;
      case "office_replace_section":
        result = await toolReplaceSection(writeArgs());
        break;
      case "office_replace_text":
        result = await toolReplaceText(writeArgs());
        break;
      case "office_highlight":
        result = await toolHighlight(args);
        break;
      case "office_clear_highlights":
        result = await toolClearHighlights(args);
        break;
      case "office_add_comment":
        result = await toolAddComment(args);
        break;
      case "office_clear_comments":
        result = await toolClearComments(args);
        break;
      default:
        throw new Error(`Unknown tool: ${name}`);
    }
    wsSend({ type: "tool_result", id, ok: true, result });
  } catch (err) {
    console.error(`[tool ${name}] failed:`, err);
    wsSend({ type: "tool_result", id, ok: false, error: err?.message ?? String(err) });
  }
}

// ---------------------------------------------------------------------------
// Office.js helpers
// ---------------------------------------------------------------------------
function headingLevel(style) {
  // Returns the heading level if the style name *contains* "Heading <N>".
  // Patent firms commonly use custom styles like "M&G-Pat App-Heading 2" or
  // "Firm-Patent-Heading-1" — we want to recognize those as headings too,
  // not just Word's built-in "Heading 1" / "Heading 2".
  if (!style) return null;
  const m = /heading[\s_-]*(\d)/i.exec(style);
  return m ? parseInt(m[1], 10) : null;
}

// Build the "view" the agent sees: list of {id, style, text}.
// Pass idMode from getParagraphsWithIds() so the id field uses the right
// addressing scheme.
function snapshotParagraphs(paragraphs, idMode) {
  return paragraphs.items.map((p, i) => ({
    index: i,
    id: getId(p, i, idMode),
    style: p.style,
    text: p.text,
  }));
}

// Find the index of a heading paragraph by case-insensitive exact text match.
function findHeadingIndex(snapshot, headingText) {
  const target = headingText.trim().toLowerCase();
  return snapshot.findIndex(p =>
    headingLevel(p.style) !== null && p.text.trim().toLowerCase() === target
  );
}

// Infer the body-text style for a section. Walks forward from `fromIdx` for
// the first non-heading paragraph and returns its style. Falls back to
// walking backward. Used so that inserted paragraphs pick up the section's
// body style (e.g. "HBH Body Text") rather than the previous paragraph's
// style — which would be a heading and would cause new "body" paragraphs to
// render as headings.
function inferBodyStyle(snapshot, fromIdx) {
  for (let i = fromIdx; i < snapshot.length; i++) {
    if (snapshot[i] && headingLevel(snapshot[i].style) === null) {
      return snapshot[i].style || null;
    }
  }
  for (let i = fromIdx - 1; i >= 0; i--) {
    if (snapshot[i] && headingLevel(snapshot[i].style) === null) {
      return snapshot[i].style || null;
    }
  }
  return null;
}

// Find the end of a section starting at headingIdx (exclusive): next paragraph
// with a heading style of same-or-higher rank (i.e. lower or equal level).
function findSectionEnd(snapshot, headingIdx) {
  const startLevel = headingLevel(snapshot[headingIdx].style);
  for (let i = headingIdx + 1; i < snapshot.length; i++) {
    const lvl = headingLevel(snapshot[i].style);
    if (lvl !== null && lvl <= startLevel) return i;
  }
  return snapshot.length;
}

// ---------------------------------------------------------------------------
// Tool: office_get_selection
// ---------------------------------------------------------------------------
async function toolGetSelection() {
  return await Word.run(async (context) => {
    const sel = context.document.getSelection();
    sel.load("text");
    const selParas = sel.paragraphs;

    // Try to load uniqueLocalId directly on the selected paragraphs — that
    // gives us stable IDs without any text-matching against the full doc
    // (which is unsafe with repeated boilerplate in patent specs).
    let idMode = "uniqueLocalId";
    try {
      selParas.load("items/text, items/style, items/uniqueLocalId");
      await context.sync();
      if (selParas.items.length > 0 && !selParas.items[0].uniqueLocalId) {
        throw new Error("uniqueLocalId not populated");
      }
    } catch {
      selParas.load("items/text, items/style");
      await context.sync();
      idMode = "index";
    }

    const selSnapshot = selParas.items.map(p => ({
      // When uniqueLocalId is unavailable, return null rather than guessing
      // an index by text-matching — duplicate paragraphs in patent specs make
      // text-match unreliable. The agent should fall back to text references.
      id: idMode === "uniqueLocalId" ? p.uniqueLocalId : null,
      style: p.style,
      text: p.text,
    }));

    return {
      text: sel.text,
      is_empty: !sel.text || sel.text.length === 0,
      paragraphs: selSnapshot,
      addressing: idMode,
    };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_read_paragraphs
// ---------------------------------------------------------------------------
async function toolReadParagraphs({ ids, heading_section, range }) {
  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const snapshot = snapshotParagraphs(paragraphs, idMode);

    let picked;
    let truncate = false;
    if (ids && ids.length > 0) {
      const set = new Set(ids);
      picked = snapshot.filter(p => set.has(p.id));
    } else if (heading_section) {
      const startIdx = findHeadingIndex(snapshot, heading_section);
      if (startIdx === -1) {
        throw new Error(`Heading not found: "${heading_section}"`);
      }
      const endIdx = findSectionEnd(snapshot, startIdx);
      picked = snapshot.slice(startIdx, endIdx);
    } else if (range) {
      const [s, e] = range;
      picked = snapshot.slice(s, e);
    } else {
      // Default: outline-ish view. Patent docs use varied heading conventions
      // (built-in Heading N, custom style names, all-caps direct formatting),
      // so we don't pre-filter — return every paragraph with its style and a
      // truncated text preview. The agent picks out the headings.
      picked = snapshot;
      truncate = true;
    }

    return {
      paragraphs: picked.map(p => ({
        id: p.id,
        style: p.style,
        text: truncate && p.text.length > 120 ? p.text.slice(0, 117) + "..." : p.text,
      })),
      total_in_doc: snapshot.length,
      addressing: idMode,
    };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_insert_paragraphs
// ---------------------------------------------------------------------------
async function toolInsertParagraphs({ after, content, track_changes, style_per_para, provenance_comment }) {
  if (!content || content.length === 0) throw new Error("content must be non-empty");

  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const snapshot = snapshotParagraphs(paragraphs, idMode);

    // Resolve the anchor paragraph's index.
    let anchorIdx;
    if (after?.id) {
      anchorIdx = findIndexById(paragraphs, after.id, idMode);
      if (anchorIdx === -1) {
        throw new Error(`Anchor paragraph not found: ${after.id}`);
      }
    } else if (after?.heading) {
      anchorIdx = findHeadingIndex(snapshot, after.heading);
      if (anchorIdx === -1) throw new Error(`Heading not found: "${after.heading}"`);
    } else {
      throw new Error("after must specify either id or heading");
    }

    const anchor = paragraphs.items[anchorIdx];
    // If no explicit style is given, infer the section's body style so we
    // don't inherit a heading's formatting. Start one past the anchor (the
    // first paragraph the new content will sit alongside).
    const inferredStyle = (style_per_para && style_per_para.length === content.length)
      ? null
      : inferBodyStyle(snapshot, anchorIdx + 1);

    // Insert each new paragraph after the previous insertion (so order is preserved).
    let cursor = anchor;
    const inserted = [];
    for (let i = 0; i < content.length; i++) {
      const p = cursor.insertParagraph(content[i], Word.InsertLocation.after);
      const explicit = style_per_para && style_per_para[i];
      if (explicit) p.style = explicit;
      else if (inferredStyle) p.style = inferredStyle;
      inserted.push(p);
      cursor = p;
    }

    if (provenance_comment && inserted.length > 0) {
      inserted[0].getRange().insertComment(provenance_comment);
    }

    await context.sync();

    // Load IDs of inserted paragraphs. With uniqueLocalId, each new paragraph
    // already has one populated post-sync; in index fallback we synthesize.
    let newIds;
    if (idMode === "uniqueLocalId") {
      for (const p of inserted) p.load("uniqueLocalId");
      await context.sync();
      newIds = inserted.map(p => p.uniqueLocalId);
    } else {
      newIds = inserted.map((_, k) => fallbackId(anchorIdx + 1 + k));
    }

    return {
      inserted_count: inserted.length,
      new_para_ids: newIds,
      addressing: idMode,
    };
    });
  });
}

// ---------------------------------------------------------------------------
// Tool: office_replace_paragraphs
// ---------------------------------------------------------------------------
async function toolReplaceParagraphs({ ids, content, track_changes, style_per_para, provenance_comment }) {
  if (!ids || ids.length === 0) throw new Error("ids must be non-empty");
  if (!content || content.length === 0) throw new Error("content must be non-empty");
  if (ids.length !== content.length) {
    throw new Error(`ids.length (${ids.length}) must equal content.length (${content.length}). To grow or shrink a section, use office_insert_paragraphs or office_replace_section.`);
  }

  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);

    // Resolve each ID to a paragraph object.
    const targets = ids.map(id => {
      const idx = findIndexById(paragraphs, id, idMode);
      if (idx === -1) {
        throw new Error(`Paragraph not found: ${id}`);
      }
      return { id, idx, paragraph: paragraphs.items[idx] };
    });

    // Replace each paragraph's text. `insertText(..., "Replace")` swaps the
    // text content while keeping the paragraph element (so style is preserved
    // unless we override it below).
    for (let i = 0; i < targets.length; i++) {
      const t = targets[i];
      t.paragraph.insertText(content[i], Word.InsertLocation.replace);
      if (style_per_para && style_per_para[i]) {
        t.paragraph.style = style_per_para[i];
      }
    }

    if (provenance_comment && targets.length > 0) {
      targets[0].paragraph.getRange().insertComment(provenance_comment);
    }

    await context.sync();

    return {
      replaced_count: targets.length,
      paragraph_ids: targets.map(t => t.id),
    };
    });
  });
}

// ---------------------------------------------------------------------------
// Tool: office_replace_text — surgical sub-paragraph search/replace
// ---------------------------------------------------------------------------
async function toolReplaceText({ paragraph_ids, find, replace, match_case, whole_word, track_changes }) {
  if (!Array.isArray(paragraph_ids) || paragraph_ids.length === 0) {
    throw new Error("paragraph_ids must be a non-empty array");
  }
  if (typeof find !== "string" || find.length === 0) {
    throw new Error("find must be a non-empty string");
  }
  if (typeof replace !== "string") {
    throw new Error("replace must be a string (use \"\" to delete)");
  }

  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);

    // Resolve target paragraphs.
    const targets = [];
    for (const id of paragraph_ids) {
      const idx = findIndexById(paragraphs, id, idMode);
      if (idx === -1) {
        throw new Error(`Paragraph not found: ${id}`);
      }
      targets.push({ id, paragraph: paragraphs.items[idx] });
    }

    // Queue searches for each target.
    const probes = targets.map(t => {
      const results = t.paragraph.search(find, {
        matchCase: !!match_case,
        matchWholeWord: !!whole_word,
      });
      results.load("items");
      return { id: t.id, results };
    });
    await context.sync();

    // Apply replacements.
    let totalReplaced = 0;
    const perParagraph = [];
    for (const p of probes) {
      const matches = p.results.items.length;
      for (const r of p.results.items) {
        r.insertText(replace, Word.InsertLocation.replace);
      }
      perParagraph.push({ paragraph_id: p.id, replacements: matches });
      totalReplaced += matches;
    }
    await context.sync();

    return {
      total_replacements: totalReplaced,
      per_paragraph: perParagraph,
    };
    });
  });
}

// ---------------------------------------------------------------------------
// Tool: office_replace_section
// ---------------------------------------------------------------------------
async function toolReplaceSection({ heading, content, track_changes, style_per_para, provenance_comment }) {
  if (!content || content.length === 0) throw new Error("content must be non-empty");

  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
    let para1 = await getParagraphsWithIds(context);
    let paragraphs = para1.paragraphs;
    let idMode = para1.idMode;
    let snapshot = snapshotParagraphs(paragraphs, idMode);

    const startIdx = findHeadingIndex(snapshot, heading);
    if (startIdx === -1) throw new Error(`Heading not found: "${heading}"`);
    const endIdx = findSectionEnd(snapshot, startIdx);

    // Capture the existing section's body style BEFORE deleting, so inserted
    // paragraphs match. Without this, new paragraphs would inherit the
    // heading's style (centered/bold/large) and the section would render as
    // a wall of headings.
    const bodyStyle = inferBodyStyle(snapshot, startIdx + 1);

    // Delete paragraphs between the heading (exclusive) and endIdx (exclusive).
    const toDelete = paragraphs.items.slice(startIdx + 1, endIdx);
    const deletedCount = toDelete.length;
    for (const p of toDelete) p.delete();
    await context.sync();

    // Re-fetch. Heading index may have shifted by 0 (we only deleted *after* it)
    // but be safe and re-locate.
    const para2 = await getParagraphsWithIds(context);
    paragraphs = para2.paragraphs;
    idMode = para2.idMode;
    snapshot = snapshotParagraphs(paragraphs, idMode);
    const newHeadingIdx = findHeadingIndex(snapshot, heading);
    if (newHeadingIdx === -1) throw new Error("Lost the heading after delete; aborting");

    const headingPara = paragraphs.items[newHeadingIdx];
    let cursor = headingPara;
    const inserted = [];
    for (let i = 0; i < content.length; i++) {
      const p = cursor.insertParagraph(content[i], Word.InsertLocation.after);
      const explicit = style_per_para && style_per_para[i];
      if (explicit) p.style = explicit;
      else if (bodyStyle) p.style = bodyStyle;
      inserted.push(p);
      cursor = p;
    }

    if (provenance_comment && inserted.length > 0) {
      inserted[0].getRange().insertComment(provenance_comment);
    }

    await context.sync();

    // Load IDs of inserted paragraphs.
    let newIds;
    if (idMode === "uniqueLocalId") {
      for (const p of inserted) p.load("uniqueLocalId");
      await context.sync();
      newIds = inserted.map(p => p.uniqueLocalId);
    } else {
      newIds = inserted.map((_, k) => fallbackId(newHeadingIdx + 1 + k));
    }

    return {
      deleted_paragraphs: deletedCount,
      inserted_count: inserted.length,
      new_para_ids: newIds,
      addressing: idMode,
    };
    });
  });
}

// ---------------------------------------------------------------------------
// Tool: office_highlight
// ---------------------------------------------------------------------------
const SEVERITY_COLOR = {
  error:     "Red",
  warning:   "Yellow",
  info:      "Turquoise",
  uncertain: "Pink",
};

async function toolHighlight({ targets }) {
  if (!Array.isArray(targets) || targets.length === 0) {
    throw new Error("targets must be a non-empty array");
  }

  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);

    // Phase 1: queue all the searches. Office.js batches operations until
    // sync(), so doing this in one pass is faster than one-sync-per-target.
    const perTarget = [];
    const pendingSearches = [];   // { idx, target, searchResults }
    for (let i = 0; i < targets.length; i++) {
      const t = targets[i];
      const idx = findIndexById(paragraphs, t.paragraph_id, idMode);
      if (idx === -1) {
        perTarget[i] = { paragraph_id: t.paragraph_id, ok: false, error: "paragraph_not_found" };
        continue;
      }
      const paragraph = paragraphs.items[idx];
      if (t.query) {
        const searchResults = paragraph.search(t.query, { matchCase: false });
        searchResults.load("items");
        pendingSearches.push({ i, t, paragraph, searchResults });
      } else {
        // Whole-paragraph highlight; no search needed.
        pendingSearches.push({ i, t, paragraph, wholeParagraph: true });
      }
    }
    await context.sync();

    // Phase 2: apply highlights.
    let totalHighlights = 0;
    for (const p of pendingSearches) {
      const color = SEVERITY_COLOR[p.t.severity || "warning"] || "Yellow";
      try {
        if (p.wholeParagraph) {
          p.paragraph.getRange().font.highlightColor = color;
          perTarget[p.i] = { paragraph_id: p.t.paragraph_id, ok: true, matches: 1, severity: p.t.severity || "warning" };
          totalHighlights += 1;
        } else {
          const items = p.searchResults.items;
          for (const r of items) {
            r.font.highlightColor = color;
          }
          perTarget[p.i] = {
            paragraph_id: p.t.paragraph_id,
            ok: items.length > 0,
            matches: items.length,
            severity: p.t.severity || "warning",
            ...(items.length === 0 ? { error: "query_not_found", query: p.t.query } : {}),
          };
          totalHighlights += items.length;
        }
      } catch (e) {
        perTarget[p.i] = { paragraph_id: p.t.paragraph_id, ok: false, error: e.message };
      }
    }
    await context.sync();

    return {
      total_highlights: totalHighlights,
      per_target: perTarget,
    };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_clear_highlights
// ---------------------------------------------------------------------------
async function toolClearHighlights(args) {
  const { paragraph_ids, heading_section, all } = args || {};

  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const snapshot = snapshotParagraphs(paragraphs, idMode);

    let targets = [];
    if (all === true) {
      targets = paragraphs.items;
    } else if (Array.isArray(paragraph_ids) && paragraph_ids.length > 0) {
      for (const id of paragraph_ids) {
        const idx = findIndexById(paragraphs, id, idMode);
        if (idx !== -1) targets.push(paragraphs.items[idx]);
      }
    } else if (typeof heading_section === "string") {
      const startIdx = findHeadingIndex(snapshot, heading_section);
      if (startIdx === -1) throw new Error(`Heading not found: "${heading_section}"`);
      const endIdx = findSectionEnd(snapshot, startIdx);
      targets = paragraphs.items.slice(startIdx, endIdx);
    } else {
      throw new Error("Specify one of paragraph_ids, heading_section, or all: true");
    }

    // Office.js batches all property assignments and applies them at
    // context.sync(). A single bad paragraph (empty paragraph, paragraph
    // in a content control, paragraph with only tracked deletions, etc.)
    // fails the whole batch with a generic InvalidArgument. So:
    //   Fast path — try the whole batch in one sync.
    //   Slow path — if the batch fails, sync one paragraph at a time and
    //               record which paragraphs rejected the change.
    // null is the Office.js named value for clearing a highlight.
    try {
      for (const p of targets) {
        p.getRange().font.highlightColor = null;
      }
      await context.sync();
      return { cleared_paragraphs: targets.length, failed_paragraphs: 0 };
    } catch (batchErr) {
      let cleared = 0;
      const failed = [];
      for (let i = 0; i < targets.length; i++) {
        try {
          targets[i].getRange().font.highlightColor = null;
          await context.sync();
          cleared++;
        } catch (e) {
          failed.push({ index: i, error: e?.message ?? String(e) });
        }
      }
      return { cleared_paragraphs: cleared, failed_paragraphs: failed.length, failures: failed };
    }
  });
}

// ---------------------------------------------------------------------------
// Tool: office_clear_comments
// ---------------------------------------------------------------------------
async function toolClearComments(args) {
  const { paragraph_ids, heading_section, all } = args || {};

  return await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();

    // Fast path: delete every comment in the doc.
    if (all === true) {
      const total = comments.items.length;
      for (const c of comments.items) c.delete();
      await context.sync();
      return { cleared_comments: total };
    }

    // Scoped path: determine which paragraph each comment is anchored on.
    // For each comment, load its contentRange.paragraphs (the paragraph[s]
    // the comment is anchored in). Then match by text against our snapshot
    // to derive paragraph indices.
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const snapshot = snapshotParagraphs(paragraphs, idMode);

    let allowedIdxs;
    if (Array.isArray(paragraph_ids) && paragraph_ids.length > 0) {
      allowedIdxs = new Set();
      for (const id of paragraph_ids) {
        const idx = findIndexById(paragraphs, id, idMode);
        if (idx !== -1) allowedIdxs.add(idx);
      }
    } else if (typeof heading_section === "string") {
      const startIdx = findHeadingIndex(snapshot, heading_section);
      if (startIdx === -1) throw new Error(`Heading not found: "${heading_section}"`);
      const endIdx = findSectionEnd(snapshot, startIdx);
      allowedIdxs = new Set();
      for (let i = startIdx; i < endIdx; i++) allowedIdxs.add(i);
    } else {
      throw new Error("Specify one of paragraph_ids, heading_section, or all: true");
    }

    // Queue paragraph loads for each comment's anchor.
    const probes = comments.items.map(c => {
      const paras = c.contentRange.paragraphs;
      paras.load("items/text");
      return { comment: c, paras };
    });
    await context.sync();

    let deletedCount = 0;
    for (const { comment, paras } of probes) {
      if (paras.items.length === 0) continue;
      const firstText = paras.items[0].text;
      const idx = snapshot.findIndex(p => p.text === firstText);
      if (idx !== -1 && allowedIdxs.has(idx)) {
        comment.delete();
        deletedCount += 1;
      }
    }
    await context.sync();

    return { cleared_comments: deletedCount };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_add_comment
// ---------------------------------------------------------------------------
async function toolAddComment({ paragraph_id, query, text }) {
  if (!text || !text.trim()) throw new Error("text must be non-empty");

  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const idx = findIndexById(paragraphs, paragraph_id, idMode);
    if (idx === -1) {
      throw new Error(`Paragraph not found: ${paragraph_id}`);
    }
    const paragraph = paragraphs.items[idx];

    let target;
    if (query) {
      const searchResults = paragraph.search(query, { matchCase: false });
      searchResults.load("items");
      await context.sync();
      if (searchResults.items.length === 0) {
        throw new Error(`Query "${query}" not found in paragraph ${paragraph_id}`);
      }
      target = searchResults.items[0]; // first match
    } else {
      target = paragraph.getRange();
    }

    target.insertComment(text);
    await context.sync();

    return { paragraph_id, anchored_on: query || "whole_paragraph", comment_added: true };
  });
}

// ---------------------------------------------------------------------------
// Selection tracking — push context_update on changes (debounced).
// ---------------------------------------------------------------------------
let selectionDebounce = null;

async function captureSelection() {
  try {
    const r = await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load("text");
      await context.sync();
      return {
        text: sel.text,
        para_id: null,           // index-based IDs aren't worth computing on every cursor tick
        para_count: null,
      };
    });
    lastSelection = r;
    refreshSelectionChip();
    if (wsReady) {
      wsSend({
        type: "context_update",
        selection: attachSelection ? lastSelection : null,
      });
    }
  } catch (e) {
    // Selection may be transient; ignore.
  }
}

function onSelectionChanged() {
  if (selectionDebounce) clearTimeout(selectionDebounce);
  selectionDebounce = setTimeout(captureSelection, 100);
}

// ---------------------------------------------------------------------------
// Composer
// ---------------------------------------------------------------------------
$composer.addEventListener("submit", (e) => {
  e.preventDefault();
  const text = $input.value.trim();
  if (!text) return;
  if (!wsReady) {
    appendEvent("Not connected to daemon.");
    return;
  }
  appendUserMessage(text);
  wsSend({ type: "user_message", text });
  setStatus("working", "Working…");
  $input.value = "";
  // After sending, the chip resets to "attached" for the next turn.
  attachSelection = true;
  refreshSelectionChip();
});

$input.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !e.shiftKey) {
    e.preventDefault();
    $composer.dispatchEvent(new Event("submit"));
  }
});

// ---------------------------------------------------------------------------
// Theme — match Word's theme (light/dark) via Office.context.officeTheme.
// If Office doesn't expose a theme (older build), the CSS
// `prefers-color-scheme: dark` media query already handles the OS-level
// preference, so this is purely an override for when Word's theme differs
// from the OS.
// ---------------------------------------------------------------------------
function applyOfficeTheme() {
  try {
    const t = Office.context && Office.context.officeTheme;
    const bg = t && t.bodyBackgroundColor;
    if (!bg) return;
    const m = /^#?([0-9a-f]{6})$/i.exec(bg);
    if (!m) return;
    const hex = m[1];
    const r = parseInt(hex.slice(0, 2), 16);
    const g = parseInt(hex.slice(2, 4), 16);
    const b = parseInt(hex.slice(4, 6), 16);
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    document.documentElement.dataset.theme = luminance < 0.5 ? "dark" : "light";
  } catch {
    /* fall through; CSS media query handles OS-level preference. */
  }
}

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------
Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    setStatus("err", `Unsupported host: ${info.host}`);
    return;
  }

  applyOfficeTheme();

  try {
    activeDocUrl = Office.context.document.url || null;
  } catch { /* ignore */ }
  refreshMismatchIndicator();

  // Wire selection event.
  Word.run(async (context) => {
    context.document.onSelectionChanged.add(onSelectionChanged);
    await context.sync();
  }).catch(err => console.warn("Could not attach selection handler:", err));

  // Capture once on boot.
  captureSelection();

  // Initialize presets UI.
  initPresets();

  // Fetch the bridge token before opening the WS — the bridge will close any
  // connection that doesn't present it.
  fetchBridgeToken().then(() => wsConnect());
});

// ===========================================================================
// Tabs
// ===========================================================================
function setActiveTab(tabName) {
  document.body.dataset.activeTab = tabName;
  document.querySelectorAll(".tab-content").forEach(el => {
    el.hidden = el.dataset.tab !== tabName;
  });
  document.querySelectorAll(".tab").forEach(btn => {
    btn.setAttribute("aria-current", btn.dataset.tab === tabName ? "page" : "false");
  });
  if (tabName === "setup") {
    loadContext();
    loadWorkspaceSection();
  }
}

document.querySelectorAll(".tab").forEach(btn => {
  btn.addEventListener("click", () => setActiveTab(btn.dataset.tab));
});

// ===========================================================================
// Presets — saved prompts that the user can pin to quick-chips or browse
// in the Library tab.
// ===========================================================================
const PRESETS_KEY = "cc-word-addin-presets-v1";

function defaultPresets() {
  return [
    // QC
    // Summarize / outline
    { id: uuid(), title: "Summarize this document", category: "Summarize",
      prompt: "Read the whole document and give me a tight summary — main argument, key points, anything notable. Don't edit the document.",
      pinned: true, auto_send: true },
    { id: uuid(), title: "Outline this document", category: "Summarize",
      prompt: "Show me the heading outline of this document with paragraph counts per section. Don't edit anything.",
      pinned: false, auto_send: true },

    // Edit
    { id: uuid(), title: "Improve writing in selection", category: "Edit",
      prompt: "Improve the writing in my current selection — clearer, tighter, no redundancy, preserve meaning. Use track changes.",
      pinned: true, auto_send: true },
    { id: uuid(), title: "Fix typos and inconsistencies", category: "Edit",
      prompt: "Scan the whole document for typos, grammar errors, and inconsistencies (terminology, capitalization, punctuation). Use office_highlight with severity 'warning' for each issue and summarize them in chat.",
      pinned: false, auto_send: true },
    { id: uuid(), title: "Simplify the selection", category: "Edit",
      prompt: "Simplify the selected paragraph for clarity without losing meaning. Use track changes.",
      pinned: false, auto_send: true },

    // Review
    { id: uuid(), title: "Add comments on this section", category: "Review",
      prompt: "Review the section my selection is in. Add Word comments on each paragraph that has a problem (unclear phrasing, weak argument, missing detail). Don't edit the text itself.",
      pinned: false, auto_send: true },

    // Research (uses context files)
    { id: uuid(), title: "Answer using my context files", category: "Research",
      prompt: "Use the context files I've added to this workspace to answer: ",
      pinned: false, auto_send: false },

    ...defaultEditingPresets(),
  ];
}

// Editing presets are factored out so the migration in initPresets can
// append them to existing users' lists without duplicating the constants.
function defaultEditingPresets() {
  return [
    { id: uuid(), title: "Clear highlighting", category: "Editing",
      prompt: "Call office_clear_highlights with arguments {\"all\": true} to remove every highlight from the document.",
      pinned: true, auto_send: true },
  ];
}

function uuid() {
  if (crypto?.randomUUID) return crypto.randomUUID();
  return "p_" + Math.random().toString(36).slice(2, 10);
}

let presets = [];

function loadPresets() {
  try {
    const raw = localStorage.getItem(PRESETS_KEY);
    if (raw) return JSON.parse(raw);
  } catch { /* ignore */ }
  return null;
}

function savePresets() {
  try { localStorage.setItem(PRESETS_KEY, JSON.stringify(presets)); } catch { /* ignore */ }
}

function initPresets() {
  const existing = loadPresets();
  if (existing === null) {
    presets = defaultPresets();
    savePresets();
  } else {
    presets = existing;
    // Migration: append the Editing category for users whose presets were
    // seeded before it existed. Skips if they already have one (i.e. they
    // migrated previously, or they added their own Editing preset).
    if (!presets.some(p => p.category === "Editing")) {
      presets.push(...defaultEditingPresets());
      savePresets();
    }
  }
  renderLibrary();
  renderQuickChips();
}

// ---- Quick chips (pinned presets) -----------------------------------------
const $quickChips = document.getElementById("quick-chips");

function renderQuickChips() {
  const pinned = presets.filter(p => p.pinned);
  if (pinned.length === 0) {
    $quickChips.hidden = true;
    $quickChips.innerHTML = "";
    return;
  }
  $quickChips.hidden = false;
  $quickChips.innerHTML = "";
  for (const p of pinned) {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "quick-chip";
    chip.textContent = p.title;
    chip.title = p.prompt.length > 200 ? p.prompt.slice(0, 197) + "…" : p.prompt;
    chip.addEventListener("click", () => usePreset(p));
    $quickChips.appendChild(chip);
  }
}

// ---- Library tab list -----------------------------------------------------
const $libraryList = document.getElementById("library-list");

function renderLibrary() {
  $libraryList.innerHTML = "";
  if (presets.length === 0) {
    const empty = document.createElement("div");
    empty.className = "library-empty";
    empty.textContent = "No presets yet. Click \"+ New preset\" to add one.";
    $libraryList.appendChild(empty);
    return;
  }

  // Group by category. Uncategorized go under "Other".
  const groups = new Map();
  for (const p of presets) {
    const k = p.category && p.category.trim() ? p.category.trim() : "Other";
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(p);
  }

  for (const [category, items] of groups) {
    const heading = document.createElement("div");
    heading.className = "library-category";
    heading.textContent = category;
    $libraryList.appendChild(heading);
    for (const p of items) {
      $libraryList.appendChild(renderPresetRow(p));
    }
  }
}

function renderPresetRow(p) {
  const row = document.createElement("div");
  row.className = "preset-row";
  row.title = p.prompt.length > 300 ? p.prompt.slice(0, 297) + "…" : p.prompt;

  // Click row → use preset
  row.addEventListener("click", (e) => {
    if (e.target.closest(".preset-actions") || e.target.closest(".preset-pin")) return;
    usePreset(p);
  });

  const title = document.createElement("div");
  title.className = "preset-title";
  title.textContent = p.title;
  row.appendChild(title);

  const meta = document.createElement("div");
  meta.className = "preset-meta";
  if (p.auto_send) meta.textContent = "auto-send";
  row.appendChild(meta);

  const pin = document.createElement("button");
  pin.type = "button";
  pin.className = "preset-pin" + (p.pinned ? " pinned" : "");
  pin.textContent = p.pinned ? "📌" : "📍";
  pin.title = p.pinned ? "Pinned (click to unpin)" : "Pin to quick chips";
  pin.addEventListener("click", (e) => {
    e.stopPropagation();
    p.pinned = !p.pinned;
    savePresets();
    renderLibrary();
    renderQuickChips();
  });
  row.appendChild(pin);

  const actions = document.createElement("div");
  actions.className = "preset-actions";
  const editBtn = document.createElement("button");
  editBtn.className = "icon-btn";
  editBtn.type = "button";
  editBtn.title = "Edit";
  editBtn.textContent = "✎";
  editBtn.addEventListener("click", (e) => { e.stopPropagation(); openPresetModal(p); });
  actions.appendChild(editBtn);

  const delBtn = document.createElement("button");
  delBtn.className = "icon-btn";
  delBtn.type = "button";
  delBtn.title = "Delete";
  delBtn.textContent = "✕";
  delBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    // No confirm dialog (Office.js taskpanes block it on some builds); the
    // user can re-add a deleted preset if needed.
    presets = presets.filter(x => x.id !== p.id);
    savePresets();
    renderLibrary();
    renderQuickChips();
  });
  actions.appendChild(delBtn);

  row.appendChild(actions);
  return row;
}

// ---- Use preset (click handler) -------------------------------------------
function usePreset(p) {
  if (p.auto_send) {
    if (!wsReady) {
      appendEvent("Not connected to daemon — can't send preset.");
      return;
    }
    appendUserMessage(p.prompt);
    wsSend({ type: "user_message", text: p.prompt });
    setStatus("working", "Working…");
    attachSelection = true;
    refreshSelectionChip();
  } else {
    $input.value = p.prompt;
    $input.focus();
    // Cursor at end so user can extend the prompt
    $input.setSelectionRange($input.value.length, $input.value.length);
  }
  setActiveTab("chat");
}

// ---- Preset editor modal --------------------------------------------------
const $presetModal = document.getElementById("preset-modal");
const $presetModalTitle = document.getElementById("preset-modal-title");
const $presetTitle = document.getElementById("preset-title");
const $presetPrompt = document.getElementById("preset-prompt");
const $presetCategory = document.getElementById("preset-category");
const $presetAutoSend = document.getElementById("preset-auto-send");
const $presetPinned = document.getElementById("preset-pinned");
const $presetSave = document.getElementById("preset-save");
const $presetCancel = document.getElementById("preset-cancel");
const $presetModalClose = document.getElementById("preset-modal-close");

let editingPresetId = null;

function openPresetModal(p) {
  if (p) {
    editingPresetId = p.id;
    $presetModalTitle.textContent = "Edit preset";
    $presetTitle.value = p.title;
    $presetPrompt.value = p.prompt;
    $presetCategory.value = p.category || "";
    $presetAutoSend.checked = !!p.auto_send;
    $presetPinned.checked = !!p.pinned;
  } else {
    editingPresetId = null;
    $presetModalTitle.textContent = "New preset";
    $presetTitle.value = "";
    $presetPrompt.value = "";
    $presetCategory.value = "";
    $presetAutoSend.checked = false;
    $presetPinned.checked = false;
  }
  $presetModal.hidden = false;
  setTimeout(() => $presetTitle.focus(), 0);
}

function closePresetModal() {
  $presetModal.hidden = true;
  editingPresetId = null;
}

$presetCancel.addEventListener("click", closePresetModal);
$presetModalClose.addEventListener("click", closePresetModal);
$presetModal.addEventListener("click", (e) => {
  if (e.target === $presetModal) closePresetModal();
});

$presetSave.addEventListener("click", () => {
  const title = $presetTitle.value.trim();
  const prompt = $presetPrompt.value;
  if (!title) { $presetTitle.focus(); return; }
  if (!prompt.trim()) { $presetPrompt.focus(); return; }

  const data = {
    title,
    prompt,
    category: $presetCategory.value.trim(),
    auto_send: $presetAutoSend.checked,
    pinned: $presetPinned.checked,
  };

  if (editingPresetId) {
    const idx = presets.findIndex(p => p.id === editingPresetId);
    if (idx !== -1) presets[idx] = { ...presets[idx], ...data };
  } else {
    presets.push({ id: uuid(), ...data });
  }

  savePresets();
  renderLibrary();
  renderQuickChips();
  closePresetModal();
});

document.getElementById("add-preset").addEventListener("click", () => openPresetModal(null));

// ===========================================================================
// Setup tab — Context files (folders or individual files saved to the
// workspace's CLAUDE.md, loaded by Claude Code each session).
// ===========================================================================
let contextCache = null;
let contextLoadingPromise = null;
const $contextList = document.getElementById("context-list");

async function removeContextEntryAt(idx) {
  if (!contextCache) return;
  contextCache = contextCache.filter((_, i) => i !== idx);
  await saveContext();
}

async function loadContext(force = false) {
  if (contextCache && !force) { renderContext(); return; }
  if (contextLoadingPromise) return contextLoadingPromise;
  contextLoadingPromise = (async () => {
    try {
      $contextList.innerHTML = '<div class="references-loading">Loading…</div>';
      const r = await sendRequest("get_context");
      contextCache = Array.isArray(r.entries) ? r.entries : [];
      renderContext();
    } catch (e) {
      showListMessage($contextList, "references-empty", `Could not load: ${e.message}`);
    } finally {
      contextLoadingPromise = null;
    }
  })();
  return contextLoadingPromise;
}

function renderContext() {
  $contextList.innerHTML = "";
  if (!contextCache || contextCache.length === 0) {
    const empty = document.createElement("div");
    empty.className = "references-empty";
    empty.textContent = "No context files yet — add a folder or file to give Claude background.";
    $contextList.appendChild(empty);
    return;
  }
  contextCache.forEach((e, i) => {
    const row = document.createElement("div");
    row.className = "reference-row";
    row.title = e.path;

    const info = document.createElement("div");
    info.className = "reference-info";
    const pathEl = document.createElement("div");
    pathEl.className = "reference-path";
    const tag = document.createElement("span");
    tag.className = "kind-tag";
    tag.textContent = e.kind || "?";
    pathEl.appendChild(tag);
    pathEl.appendChild(document.createTextNode(e.path));
    info.appendChild(pathEl);
    if (e.description) {
      const desc = document.createElement("div");
      desc.className = "reference-description";
      desc.textContent = e.description;
      info.appendChild(desc);
    }
    row.appendChild(info);

    const remove = document.createElement("button");
    remove.className = "reference-remove";
    remove.type = "button";
    remove.title = "Remove";
    remove.textContent = "✕";
    remove.addEventListener("click", () => removeContextEntryAt(i));
    row.appendChild(remove);

    $contextList.appendChild(row);
  });
}

async function saveContext() {
  if (!contextCache) return false;
  try {
    const r = await sendRequest("set_context", { entries: contextCache });
    if (r.errors && r.errors.length > 0) {
      const lines = r.errors.map(e => `${e.path} — ${e.error}`).join("\n");
      alert(`Some entries could not be saved:\n\n${lines}`);
    }
    contextCache = Array.isArray(r.saved) ? r.saved : contextCache;
    renderContext();
    return true;
  } catch (e) {
    alert(`Could not save: ${e.message}`);
    return false;
  }
}

// ---- Add-folder modal (shared between guidelines + samples) ----------------
const $addFolderModal = document.getElementById("add-folder-modal");
const $addFolderModalTitle = document.getElementById("add-folder-modal-title");
const $addFolderPath = document.getElementById("add-folder-path");
const $addFolderDescription = document.getElementById("add-folder-description");
const $addFolderError = document.getElementById("add-folder-error");
const $addFolderSave = document.getElementById("add-folder-save");
const $addFolderCancel = document.getElementById("add-folder-cancel");
const $addFolderModalClose = document.getElementById("add-folder-modal-close");
const $addFolderBrowse = document.getElementById("add-folder-browse");

function openAddFolderModal() {
  $addFolderModalTitle.textContent = "Add folder or file";
  $addFolderPath.value = "";
  $addFolderDescription.value = "";
  $addFolderError.hidden = true;
  $addFolderModal.hidden = false;
  setTimeout(() => $addFolderPath.focus(), 0);
}
function closeAddFolderModal() {
  $addFolderModal.hidden = true;
}

document.querySelectorAll(".add-folder-trigger").forEach(btn => {
  btn.addEventListener("click", () => openAddFolderModal());
});
$addFolderCancel.addEventListener("click", closeAddFolderModal);
$addFolderModalClose.addEventListener("click", closeAddFolderModal);
$addFolderModal.addEventListener("click", (e) => {
  if (e.target === $addFolderModal) closeAddFolderModal();
});

$addFolderSave.addEventListener("click", async () => {
  const path = $addFolderPath.value.trim();
  const description = $addFolderDescription.value.trim();
  $addFolderError.hidden = true;
  if (!path) {
    $addFolderError.textContent = "Path is required.";
    $addFolderError.hidden = false;
    return;
  }
  if (!contextCache) await loadContext();
  if (!contextCache) contextCache = [];
  contextCache = [...contextCache, { path, description }];
  const ok = await saveContext();
  if (ok) closeAddFolderModal();
});

// ---- Native folder/file picker ---------------------------------------------
// Forwards the pick request through the daemon to the Electron main process,
// which shows a real macOS NSOpenPanel. This is the same panel every native
// Mac app uses, so it can navigate to Google Drive, iCloud, "Shared with me",
// recent items, sidebar shortcuts — none of which a synthetic in-page browser
// can reach. Resolves to `{ path, kind }` or `null` if the user cancelled.
async function pickPathNative({ start_path = null, include_files = false, title = null } = {}) {
  const r = await sendRequest("pick_path", {
    default_path: start_path,
    include_files,
    title,
  });
  if (!r.ok) throw new Error(r.error || "Picker failed");
  if (r.canceled) return null;
  return { path: r.path, kind: r.kind };
}

$addFolderBrowse.addEventListener("click", async () => {
  const startPath = $addFolderPath.value.trim() || currentWorkspaceCwd || null;
  try {
    const picked = await pickPathNative({ start_path: startPath, include_files: true });
    if (!picked) return;
    $addFolderPath.value = picked.path;
    $addFolderDescription.focus();
  } catch (e) {
    console.error("[picker]", e);
  }
});

$addFolderPath.addEventListener("keydown", (e) => {
  if (e.key === "Enter") { e.preventDefault(); $addFolderSave.click(); }
});
$addFolderDescription.addEventListener("keydown", (e) => {
  if (e.key === "Enter") { e.preventDefault(); $addFolderSave.click(); }
});

// ===========================================================================
// Workspace selection — managed in the Setup tab's Workspace section.
// The topbar chip is read-only status (current workspace name + mismatch warning);
// clicking it navigates to the Setup tab where the actual switching UI lives.
// ===========================================================================
let currentWorkspaceCwd = null;

// Path containment check that handles the /Workspace-10 vs /Workspace-1 trap —
// `startsWith` alone returns true for both. We normalize the parent by
// stripping any trailing separator and then require child to be the parent
// exactly OR start with `<parent>/`. Case-sensitive (matches Node's path
// semantics on macOS even though HFS+ is case-insensitive — close enough for
// our use).
function isInOrUnder(child, parent) {
  if (!child || !parent) return false;
  const p = parent.replace(/\/+$/, "");
  return child === p || child.startsWith(p + "/");
}

// Replace a list container with a single-line empty-state message. Uses
// textContent so an attacker-controlled error string can't inject markup.
function showListMessage(container, klass, message) {
  container.innerHTML = "";
  const div = document.createElement("div");
  div.className = klass;
  div.textContent = message;
  container.appendChild(div);
}

const $workspaceChip = document.getElementById("workspace-chip");
const $workspaceFolder = document.getElementById("workspace-folder");
const $workspacesList = document.getElementById("workspaces-list");
const $addWorkspace = document.getElementById("add-workspace");
const $workspaceError = document.getElementById("workspace-error");
const $workspaceWarning = document.getElementById("workspace-warning");

// Suggest-banner: shown in Setup → Workspace folder when the active doc's
// folder isn't inside the current workspace and the daemon's suggestWorkspaceRoot
// returns a candidate. Lets the user one-click confirm, pick a different
// folder, or dismiss.
const $workspaceSuggest = document.getElementById("workspace-suggest");
const $workspaceSuggestPath = document.getElementById("workspace-suggest-path");
const $workspaceSuggestAccept = document.getElementById("workspace-suggest-accept");
const $workspaceSuggestPick = document.getElementById("workspace-suggest-pick");
const $workspaceSuggestDismiss = document.getElementById("workspace-suggest-dismiss");
let pendingWorkspaceSuggestion = null;     // { cwd, confidence } from the daemon
let dismissedWorkspaceSuggestionForDoc = null; // activeDocUrl the user dismissed for

function setWorkspaceDisplay(cwd) {
  currentWorkspaceCwd = cwd;
  const name = cwd ? cwd.split(/[\\/]/).filter(Boolean).pop() : "(no workspace)";
  $workspaceFolder.textContent = name;
  $workspaceFolder.title = cwd || "";
  refreshMismatchIndicator();
}

// Decide whether the workspace chip should show a warning. True when an active
// doc exists and its filesystem location is NOT inside the current workspace
// folder — the agent's filesystem tools won't see the doc's siblings.
function refreshMismatchIndicator() {
  let mismatch = false;
  if (activeDocUrl && currentWorkspaceCwd) {
    let docPath = decodeURIComponent(activeDocUrl.replace(/^file:\/\//, ""));
    // Strip the filename
    const docDir = docPath.split(/[\\/]/).slice(0, -1).join("/");
    // Mismatch if docDir isn't underneath the workspace cwd
    mismatch = docDir && !isInOrUnder(docDir, currentWorkspaceCwd);
  }
  const baseTitle = "The agent reads source files (CLAUDE.md, notes, references) from the workspace folder — click to switch workspaces.";
  if (mismatch) {
    $workspaceChip.classList.add("mismatch");
    $workspaceWarning.hidden = false;
    $workspaceChip.title = baseTitle + " (⚠ The current workspace doesn't match the doc you're editing.)";
  } else {
    $workspaceChip.classList.remove("mismatch");
    $workspaceWarning.hidden = true;
    $workspaceChip.title = baseTitle;
  }
}

async function refreshWorkspaceFromDaemon() {
  try {
    const r = await sendRequest("get_cwd_state");
    if (r.current_cwd) setWorkspaceDisplay(r.current_cwd);
    await refreshWorkspaceSuggestBanner();
  } catch { /* ignore on initial boot */ }
}

async function loadWorkspaceSection() {
  $workspaceError.hidden = true;
  await refreshWorkspaceSuggestBanner();
  try {
    const r = await sendRequest("get_cwd_state");
    renderWorkspacesList(r.recent || [], r.current_cwd);
  } catch (e) {
    showListMessage($workspacesList, "references-empty", `Could not load: ${e.message}`);
  }
}

// Banner controller: shown when the active Word doc lives outside the
// current workspace and the daemon has a candidate workspace root to suggest.
// When confidence=marker AND autoSwitchWorkspace is on, we skip the banner
// and fire the switch directly (the user previously committed to that
// folder by dropping a marker; no reason to ask again).
async function refreshWorkspaceSuggestBanner() {
  pendingWorkspaceSuggestion = null;
  $workspaceSuggest.hidden = true;

  if (!activeDocUrl) return;
  if (dismissedWorkspaceSuggestionForDoc === activeDocUrl) return;

  const docDir = decodeURIComponent(activeDocUrl.replace(/^file:\/\//, ""))
    .split(/[\\/]/).slice(0, -1).join("/");
  if (currentWorkspaceCwd && isInOrUnder(docDir, currentWorkspaceCwd)) return;

  let suggestion = null;
  try {
    const r = await sendRequest("suggest_workspace", { doc_path: activeDocUrl });
    suggestion = r.suggestion;
  } catch (e) {
    console.warn("[suggest_workspace]", e);
    return;
  }
  if (!suggestion) return;

  if (suggestion.confidence === "marker" && settings.autoSwitchWorkspace) {
    doSwitch(null, { autodetectFromDoc: true });
    return;
  }

  pendingWorkspaceSuggestion = suggestion;
  $workspaceSuggestPath.textContent = suggestion.cwd;
  $workspaceSuggestPath.title = suggestion.cwd;
  $workspaceSuggest.hidden = false;
}

function renderWorkspacesList(recent, currentCwd) {
  // Ensure the currently-active workspace is in the list even if not in recent
  // (e.g., daemon was launched at this folder but no one's "switched" to it
  // yet, so it isn't in the recents file).
  let list = Array.isArray(recent) ? [...recent] : [];
  if (currentCwd && !list.some(f => f.cwd === currentCwd)) {
    list.unshift({
      cwd: currentCwd,
      display_name: currentCwd.split(/[\\/]/).filter(Boolean).pop(),
    });
  }

  $workspacesList.innerHTML = "";
  if (list.length === 0) {
    $workspacesList.innerHTML = '<div class="references-empty">No workspaces yet.</div>';
    return;
  }

  for (const f of list) {
    const isActive = f.cwd === currentCwd;
    const row = document.createElement("div");
    row.className = "reference-row" + (isActive ? " active" : "");
    row.title = f.cwd;
    if (!isActive) row.style.cursor = "pointer";

    const info = document.createElement("div");
    info.className = "reference-info";
    const pathEl = document.createElement("div");
    pathEl.className = "reference-path";
    if (isActive) {
      const tag = document.createElement("span");
      tag.className = "active-tag";
      tag.textContent = "Active";
      pathEl.appendChild(tag);
    }
    pathEl.appendChild(document.createTextNode(
      f.display_name || f.cwd.split(/[\\/]/).filter(Boolean).pop()
    ));
    const subPath = document.createElement("div");
    subPath.className = "reference-description";
    subPath.style.fontFamily = "ui-monospace, SFMono-Regular, Menlo, Consolas, monospace";
    subPath.textContent = f.cwd;
    info.appendChild(pathEl);
    info.appendChild(subPath);
    row.appendChild(info);

    if (!isActive) {
      const remove = document.createElement("button");
      remove.className = "reference-remove";
      remove.type = "button";
      remove.title = "Remove from recent";
      remove.textContent = "✕";
      remove.addEventListener("click", async (e) => {
        e.stopPropagation();
        try {
          await sendRequest("forget_folder", { cwd: f.cwd });
          loadWorkspaceSection();
        } catch (err) {
          console.warn("forget_folder failed:", err);
        }
      });
      row.appendChild(remove);

      row.addEventListener("click", () => doSwitch(f.cwd));
    }

    $workspacesList.appendChild(row);
  }
}

async function doSwitch(cwd, { autodetectFromDoc = false } = {}) {
  $workspaceError.hidden = true;
  try {
    const payload = autodetectFromDoc
      ? { autodetect_from_doc: activeDocUrl }
      : { cwd };
    const r = await sendRequest("set_cwd", payload);
    if (!r.ok) throw new Error(r.error || "switch failed");
    // The daemon emits cwd_changed which updates the chip via assistant_event.
    // Clear chat — visually distinguishing the new session from the old.
    $messages.innerHTML = "";
    assistantTurnElem = null;
    // Refresh the workspace section so recent-folders ordering updates.
    loadWorkspaceSection();
  } catch (e) {
    $workspaceError.textContent = e.message;
    $workspaceError.hidden = false;
  }
}

// Clicking the topbar chip jumps to the Setup tab where the workspace UI lives.
$workspaceChip.addEventListener("click", () => setActiveTab("setup"));

// Suggest-banner buttons.
$workspaceSuggestAccept.addEventListener("click", () => {
  if (!pendingWorkspaceSuggestion) return;
  const cwd = pendingWorkspaceSuggestion.cwd;
  pendingWorkspaceSuggestion = null;
  $workspaceSuggest.hidden = true;
  doSwitch(cwd);
});
$workspaceSuggestPick.addEventListener("click", async () => {
  const startPath = pendingWorkspaceSuggestion?.cwd || null;
  pendingWorkspaceSuggestion = null;
  $workspaceSuggest.hidden = true;
  try {
    const picked = await pickPathNative({ start_path: startPath, title: "Choose a workspace folder" });
    if (picked) doSwitch(picked.path);
  } catch (e) {
    console.error("[picker]", e);
  }
});
$workspaceSuggestDismiss.addEventListener("click", () => {
  dismissedWorkspaceSuggestionForDoc = activeDocUrl;
  pendingWorkspaceSuggestion = null;
  $workspaceSuggest.hidden = true;
});

// "+ Add workspace" — opens the native folder picker; on pick, switch to that folder.
$addWorkspace.addEventListener("click", async () => {
  try {
    const picked = await pickPathNative({ title: "Choose a workspace folder" });
    if (picked) doSwitch(picked.path);
  } catch (e) {
    console.error("[picker]", e);
  }
});
