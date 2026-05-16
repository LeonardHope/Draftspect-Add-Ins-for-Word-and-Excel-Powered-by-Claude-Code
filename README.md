# Claude Code Add-Ins for Word and Excel

[![CI](https://github.com/LeonardHope/Claude-Code-Add-Ins-For-Word-and-Excel/actions/workflows/ci.yml/badge.svg)](https://github.com/LeonardHope/Claude-Code-Add-Ins-For-Word-and-Excel/actions/workflows/ci.yml)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

Open Word or Excel, click a button, and chat with **your own** Claude Code — the same agent, with the same skills, MCP servers, and custom instructions you already use in the terminal — but now it can also read your document, edit your cells, and leave comments on your paragraphs.

The headline feature: **point Claude at any folders or files on your machine** — your notes, prior drafts, reference exports, a vendor's specs, last quarter's deck — and they become context for whatever you're writing in Word or Excel. The agent reads them on demand, the same way it would in the terminal.

Everything runs on your computer; only API calls go to Anthropic.

---

## Status

Single-user tool you run on your own machine — intentionally scoped, not a hardened multi-tenant service (see [Security & privacy](#security--privacy)). macOS is the primary platform; Windows is supported and tested under Parallels.

- **Word** — the more mature surface; most of the polish is here.
- **Excel** — newer; the tools work but the UX trails Word.

---

## What you can ask it to do

Open your document, click the add-in, and try things like:

- _"Summarize this contract in plain English."_
- _"Improve the writing in the selection. Use track changes."_
- _"Find every cell in column C that contains an email and copy it to column E."_
- _"Review the Background section. Highlight anything weak in yellow and leave a comment explaining why."_

The really useful asks are the ones that bring **context files** into play — folders or files you've added in the Setup tab become background material the agent can read on demand:

- _"Compare this draft against my notes folder and tell me what's missing."_
- _"Rewrite the Methodology section using the terminology from `~/Research/glossary.md`."_
- _"Cross-check every figure number in this paper against the data in `~/Project/results/`."_
- _"Pull the financial assumptions from last quarter's deck into a new sheet here."_

Because it's your local Claude Code, anything you've configured — custom agents, MCP servers, your `CLAUDE.md` files, hooks — still applies.

---

## Prerequisites

Get these in place **before** you launch the app:

| Requirement                     | How to check / get it                                                                                                                                                                                                                                |
| ------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **macOS** _or_ Windows 10/11    | (anything else won't load the manifests)                                                                                                                                                                                                             |
| **Node.js 18+**                 | `node --version` — install via [nodejs.org](https://nodejs.org/) or [Volta](https://volta.sh/) / [nvm](https://github.com/nvm-sh/nvm)                                                                                                                |
| **Microsoft Word and/or Excel** | Microsoft 365, Office 2019, or Office 2021. Word build needs **WordApi 1.4+** (any reasonably recent build). Excel needs **ExcelApi 1.4+**. The manifests reject older versions automatically.                                                       |
| **Claude Code, signed in**      | `claude` in a terminal. Sign in once (Pro/Max subscription works), and the daemon picks up your OAuth from the system keychain. **Or** export `ANTHROPIC_API_KEY=sk-ant-...` in your shell before launch; that takes precedence over keychain OAuth. |
| **Git**                         | for cloning the repo                                                                                                                                                                                                                                 |

If you don't have Claude Code yet, [install it](https://docs.claude.com/en/docs/claude-code/overview) first. The add-in is a thin layer on top of it — without it nothing connects.

---

## Install

```bash
git clone https://github.com/LeonardHope/Claude-Code-Add-Ins-For-Word-and-Excel.git
cd Claude-Code-Add-Ins-For-Word-and-Excel
npm install
npm start
```

That's it. A tray icon (menu bar on macOS, system tray on Windows) appears. The first launch offers to **install the add-in into Word and Excel for you** — click _Install_ and you're done.

Open Word or Excel, go to **Insert → Office Add-ins → Shared Folder**, and pick _Claude Code for Word_ or _Claude Code for Excel_.

> **One install per host.** The manifests use fixed add-in `<Id>` GUIDs, so a single Word (or Excel) installation can only sideload one copy of this add-in at a time. To run two clones side by side, change the `<Id>` GUID in one clone's `manifests/word.xml` / `manifests/excel.xml` to a fresh UUID.

> **Daemon-only debugging.** If something's wrong and you want to see daemon output in the terminal, run `npm run dev` instead of `npm start`. Skips the Electron shell.

---

## How it works

```
   ┌─ Your computer ────────────────────────────────────────┐
   │                                                        │
   │   Tray app  ─► Daemon ◄──WebSocket──► Task pane        │
   │   (Electron)   (Node)                 (in Word/Excel)  │
   │                  │                                     │
   │                  │ forwards MCP servers from           │
   │                  ▼ ~/.claude.json                      │
   │             Your other MCP servers                     │
   │             (Visio, Gmail, custom…)                    │
   └──────────────────┬─────────────────────────────────────┘
                      │
                      │ Claude Agent SDK (local OAuth or API key)
                      ▼
              ┌───────────────┐
              │ Anthropic API │
              └───────────────┘
```

Three pieces, all on your machine:

1. **Tray app** (Electron) — spawns and watches the daemon, installs/uninstalls the add-in, surfaces native file pickers.
2. **Daemon** (Node) — wraps the Claude Agent SDK; speaks a tiny WebSocket protocol with the task pane; forwards every MCP server in your `~/.claude.json` into the agent session.
3. **Task pane** — the HTML panel Word and Excel show on the right. Uses Office.js to read/edit the document; chats with the daemon over `ws://127.0.0.1:47823`.

The task pane and daemon both live on `localhost`. Nothing on your network can talk to either of them — the WebSocket bridge requires a per-launch token, and the HTTP server only accepts the task pane's own origin.

---

## Using the add-in

1. **Launch the tray app first.** Word/Excel won't connect until the daemon is up.
2. Open a document or workbook. Open the add-in panel from **Insert → My Add-ins**.
3. In the **Setup** tab, pick a **workspace folder**.
   - The agent's `cwd`. Determines which `CLAUDE.md`, context files, and saved conversation history apply.
   - Auto-detected via marker (`CLAUDE.md` or `.claude`) or proposed via a heuristic banner. On explicit pick, an empty `CLAUDE.md` is dropped so next time is silent.
4. **Add context files.** This is the part that makes the add-in actually useful. In the Setup tab → Context files, click _Add folder or file_ and point at anything you'd want Claude to know about while it's working in this document — notes, prior drafts, source exports, a vendor's spec, a glossary. Each entry can have a one-line description so Claude knows when to consult it. Entries are saved into the workspace's `CLAUDE.md`, so the agent reads them on demand via standard `Read` / `Glob` / `Grep`.
5. Chat in the **Chat** tab. Pinned presets (chips above the input) are one-click prompts.

### Default presets

| Preset                            | What it does                                    |
| --------------------------------- | ----------------------------------------------- |
| **Summarize this document**       | Read top-to-bottom, return a tight summary      |
| **Outline this document**         | Heading outline + paragraph counts              |
| **Improve writing in selection**  | Tighten the current selection, track changes on |
| **Fix typos and inconsistencies** | Sweep + highlight + chat summary                |
| **Simplify the selection**        | Plain-language rewrite with track changes       |
| **Add comments on this section**  | Review-mode pass, no text edits                 |
| **Answer using my context files** | RAG over the folders you've added               |
| **Clear highlighting**            | Wipe every highlight in one click               |

### Settings (Setup tab → Preferences)

- **Track changes mode.** Always / modifications only / never. Always-on is the safe default — every edit is reviewable.
- **Auto-switch workspace.** When you switch documents, swap the workspace silently if a marker is found. Off by default.
- **Show diagnostic info in chat.** Session-init banners, tool-call bubbles, turn boundaries. Useful while debugging; noisy day-to-day.

---

## What the agent can touch in your documents

### Word

Ten tools, all going through Office.js (no file-system mutation of the live `.docx`):

| Tool                        | Use it for                                              |
| --------------------------- | ------------------------------------------------------- |
| `office_get_selection`      | The implicit subject of "this", "here", "the selection" |
| `office_read_paragraphs`    | Read by ID, heading section, or range                   |
| `office_insert_paragraphs`  | Add new content after a paragraph or heading            |
| `office_replace_paragraphs` | Whole-paragraph rewrites, 1-to-1                        |
| `office_replace_section`    | Find a heading, replace its section                     |
| `office_replace_text`       | Surgical sub-paragraph search/replace                   |
| `office_highlight`          | Color-coded by severity: error/warning/info/uncertain   |
| `office_clear_highlights`   | By paragraph, section, or all                           |
| `office_add_comment`        | Anchored on a paragraph or specific text                |
| `office_clear_comments`     | By paragraph, section, or all                           |

Every write tool respects your track-changes setting.

### Excel

Seven tools, A1 notation, 2D values arrays:

| Tool                                      | Use it for                                            |
| ----------------------------------------- | ----------------------------------------------------- |
| `excel_get_selected_range`                | The implicit subject of "these cells"                 |
| `excel_list_sheets`                       | Worksheet inventory + used ranges                     |
| `excel_read_range`                        | Read a range or a whole sheet's used range            |
| `excel_write_range`                       | Write a 2D values array (shape-checked)               |
| `excel_find_value`                        | Substring / whole-cell match across one or all sheets |
| `excel_insert_rows` / `excel_delete_rows` | 1-based row indices                                   |

### What the agent will refuse

- **Filesystem writes to the live document.** A permission guard refuses `Write`/`Edit`/`MultiEdit` against `.docx`/`.docm`/`.xlsx`/`.xlsm` paths (and any `Bash` command whose string mentions them). Office holds your document open with unsaved changes; a filesystem write would corrupt it. The agent must use the in-host tools instead.

---

## Troubleshooting

<details>
<summary><strong>Task pane shows "Disconnected — retrying…"</strong></summary>

The daemon isn't running, or crashed. Open the tray menu → **Open logs** (`~/.claude/office-addins/daemon.log`). If the log ends mid-startup, the daemon will auto-restart up to 3 times before giving up.

</details>

<details>
<summary><strong>"Sign-in required" banner appears</strong></summary>

The Agent SDK couldn't authenticate. Either:

- Sign in to Claude Code: `claude` in a terminal, follow the prompt.
- Or set `ANTHROPIC_API_KEY` in your shell and **relaunch** the tray app (the daemon reads env at boot).

After signing in, quit the tray app and reopen it.

</details>

<details>
<summary><strong>An MCP server I configured isn't visible to the agent</strong></summary>

At daemon start the log shows `Loaded N MCP server(s) from ~/.claude.json: …` and an HTTP preflight line per server. The SDK silently drops any server whose initial handshake fails — restart the daemon (tray menu → **Restart daemon**) once the server is back up.

`~/.claude.json` is read **once at daemon boot**. If you add or change an MCP server while the daemon is running, it won't be picked up until you restart — use the tray menu's **Restart daemon** item (no need to quit and relaunch the whole app).

</details>

<details>
<summary><strong>Auto-detect can't find my workspace</strong></summary>

If your folder has no `CLAUDE.md` or `.claude` marker and the heuristic doesn't fire, click **Pick…** in the suggest banner (or **+ Add workspace** in Setup) and choose the folder yourself. A `CLAUDE.md` is auto-created so the next open is silent.

</details>

<details>
<summary><strong>Auto-install of the add-in didn't work</strong></summary>

Quickest manual fallback (macOS and Windows): in Word/Excel, **Insert → My Add-ins → Manage My Add-ins → Upload My Add-in**, point at the appropriate file in `manifests/`.

If that menu item is missing on your build:

- **macOS:** copy `manifests/word.xml` to `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` (create the folder if missing). Same for Excel.
- **Windows:** put both manifests in a folder, then File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs → add the path → check **Show in Menu** → restart Office. Microsoft's full guide: [Sideload an Office Add-in on Windows](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).
</details>

<details>
<summary><strong>Context-file changes don't seem to take effect</strong></summary>

Saving any context entry restarts the agent session (preserving conversation history via `resume`) so the new `CLAUDE.md` loads. If you don't see the change, check that you're on the correct workspace — the topbar chip shows the active one.

</details>

---

## Authentication & distribution

The daemon authenticates via the Claude Agent SDK, which reads your **Claude Code OAuth credential** from the system keychain by default. Signed in to a Claude Pro/Max account? The add-in just works.

To use an API key instead, export it before launching:

```bash
export ANTHROPIC_API_KEY=sk-ant-...
npm start
```

The SDK prefers `ANTHROPIC_API_KEY` over keychain OAuth when present.

> **Distribution note.** This repo is open source — each user clones and runs it on their own machine with their own auth. That use is sanctioned. Shipping a hosted/packaged product that uses Anthropic subscription OAuth on behalf of users requires partner approval; see the [Agent SDK overview](https://docs.claude.com/en/api/agent-sdk/overview). If this ever becomes a product, expect to switch to BYO-API-key.

---

## Security & privacy

This is a single-user tool you run on your own machine. The threat model is "your own machine, your own documents" — not a hardened multi-tenant service. Specifics worth knowing:

- **The `.docx`/`.xlsx` filesystem-write denial is an accident guard, not a security boundary.** It refuses obvious `Write`/`Edit`/`Bash` attempts to overwrite an open Office file (which would corrupt it), but it is regex-based — an agent determined to bypass it (string-built paths, `base64`-decoded payloads, etc.) can. It exists to stop _accidental_ overwrites, not to contain a hostile agent. Edit Office files through the `office_*` / `excel_*` tools, not the filesystem.
- **The local bridge is loopback-only and token-gated, but trusts every process running as you.** The WebSocket bridge binds `127.0.0.1` and requires a per-daemon 24-byte token; `/bridge-token` serves that token over same-origin HTTP. Browser cross-origin reads are blocked by CORS, but any _local process_ running as your user (another app, a script) can read the token file and the endpoint. There's no protection against a malicious local process — out of scope for a personal tool.
- **`daemon.log` records your message text verbatim.** Every chat message you send is written in plaintext to `~/.claude/office-addins/daemon.log` (diagnostics, reachable from the tray menu). If you put confidential documents or prompts through the tool, treat that log as sensitive — it's not redacted. Delete it to scrub history.

---

## Developing

### File layout

```
.
├── app/
│   ├── main.mjs           Electron tray shell; spawns + watches daemon
│   ├── sideload.mjs       Add-in install/uninstall (mac wef/, win registry)
│   └── tray-icon.png
├── daemon/
│   ├── index.mjs          Agent SDK loop, permission handler, system prompt
│   ├── bridge.mjs         WebSocket server + tool-call protocol
│   ├── office-tools.mjs   Word + Excel tool defs (zod schemas)
│   ├── workspace.mjs      Folder resolver + suggester + marker creator
│   ├── context.mjs        Per-workspace context-file block in CLAUDE.md
│   ├── sessions.mjs       Recent-workspace persistence
│   └── system-prompt.md   Office-add-in-specific system prompt
├── taskpane/
│   ├── shared/
│   │   ├── taskpane.js    Entry point: WS, dispatcher, UI, composer, boot
│   │   ├── tools-word.js  office_* tool implementations + Word helpers
│   │   ├── tools-excel.js excel_* tool implementations + Excel helpers
│   │   ├── paths.js       Pure path helpers (URL/Windows/POSIX normalization)
│   │   └── styles.css
│   ├── word/index.html
│   ├── excel/index.html
│   └── icon-{32,80}.png   Add-in icons (regen via scripts/build-icons.py)
├── manifests/
│   ├── word.xml
│   └── excel.xml
├── scripts/
│   └── build-icons.py     Regenerate placeholder icons
├── tests/
│   ├── workspace.test.mjs
│   ├── context.test.mjs
│   └── sessions.test.mjs
├── .github/workflows/ci.yml
├── package.json
└── README.md
```

### Tests

Pure-logic modules have unit tests (no Office.js, no SDK). Run them with:

```bash
npm test
```

`node:test` runs all of `tests/*.test.mjs`. CI runs the same suite on Node 20.x and 22.x against every PR.

### Convention

- Feature branches + PRs; nothing direct to `main`.
- CI must be green before merge.
- Don't bypass the filesystem-write guard on Office files. It's there to protect the user's unsaved work.

---

## License

MIT. See [LICENSE](LICENSE).
