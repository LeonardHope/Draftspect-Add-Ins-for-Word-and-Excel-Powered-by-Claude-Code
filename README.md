# Office Claude — Powered by Claude

> _Working title; rename before public release._

Microsoft Office add-ins (Word + Excel) that bring your **local Claude Code** — every skill, custom agent, MCP server, hook, and `CLAUDE.md` you already have — directly into Word and Excel. The agent reads and edits the active document through Office.js, reads any folders or files you've added as context using the same `Read`/`Glob`/`Grep`/`Bash` it uses in your terminal, and can reach every other MCP-connected app you've configured.

Each user runs the daemon on their own machine with their own Claude Code subscription (or API key). Nothing leaves your computer except API calls to Anthropic.

## Status

Early POC. Word host is the more mature surface (forked from a patent-drafting add-in that exercised it on real work). Excel host is new — tools work but the polish lags Word.

Forked + generalized from a private patent-drafting-specific build.

---

## What it does

### Word
- **10 editing tools** the agent uses to act on the active document: get selection, read paragraphs, insert/replace paragraphs, surgical sub-paragraph replace, replace section, highlight, clear highlights, add comment, clear comments.
- **Track-changes-aware** — user-configurable per workspace (always / modifications-only / never), respected by every write tool.
- **Comments + highlights** with severity colors (`error`/`warning`/`info`/`uncertain` → red/yellow/turquoise/pink).
- **`.docx` write safety** — a `canUseTool` guard refuses any filesystem write against `.docx`/`.docm` files. Word holds them open; filesystem writes would corrupt.

### Excel
- **7 tools**: get selected range, list sheets, read range, write range, find value, insert rows, delete rows.
- **2D values addressing** in A1 notation, sheet-qualified when needed.
- **Shape validation** — `excel_write_range` refuses mismatched dimensions before touching the workbook.
- **`.xlsx`/`.xlsm` write safety** — same filesystem-write refusal as Word.

### Shared
- **Workspace folder** — pick any folder; that becomes the agent's cwd. Auto-detected via marker file (`CLAUDE.md` or `.claude`), or proposed via a smart heuristic banner when no marker exists. On explicit pick, an empty `CLAUDE.md` is dropped so next time is silent.
- **Context files** — list of folders or specific files saved to the workspace's `CLAUDE.md`. The agent reads them on demand.
- **Native macOS / Windows folder pickers** — NSOpenPanel on Mac, Common Item Dialog on Windows. Google Drive, iCloud, OneDrive, Shared-with-me all reachable.
- **Per-workspace conversation history** via the Agent SDK's `resume` — switching workspaces re-loads the prior conversation for that folder.
- **Cross-app MCP** — every MCP server configured in `~/.claude.json` (Visio, Gmail, your custom ones) is forwarded into the agent session.

---

## Prerequisites

- macOS (tested) or Windows 10/11 (best-effort; tested via Parallels VM)
- Node.js 18+ (for the daemon)
- Microsoft Word and/or Excel (Microsoft 365, 2019, or 2021)
- **Claude Code installed and signed in** — the daemon inherits your OAuth credential from the system keychain. Alternatively, set `ANTHROPIC_API_KEY` in your environment to use an API key.

### Install

```bash
git clone <repo-url> "Office Claude"
cd "Office Claude"
npm install
```

### Run

```bash
npm start
```

A tray icon appears (menu bar on macOS, system tray on Windows). Click it for the menu: status, logs, restart daemon, open Word / Excel, quit.

For daemon-only debugging without the Electron shell:

```bash
npm run dev
```

---

## Install the add-in in Word + Excel

**The Electron app installs itself.** On first launch, you'll see:

> *Install Office Claude in Word + Excel?* — Click **Install**.

After that the add-in is registered with both apps. Open Word or Excel, then **Insert → Office Add-ins → Shared Folder**, and pick *Office Claude (Word)* or *Office Claude (Excel)*. Reinstall or uninstall any time from the tray menu.

Behind the scenes:
- **macOS** — drops `manifests/{word,excel}.xml` into the per-host `wef/` folder inside Office's container sandbox. Word and Excel scan that folder at launch.
- **Windows** — copies the manifests to `%APPDATA%\Office Claude\manifests\` and registers that folder as a Trusted Catalog in the Office Trust Center via `HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\`. No admin rights needed.

### If the auto-install didn't work

Quickest manual fallback (works the same on macOS and Windows): in Word or Excel, **Insert → My Add-ins → Manage My Add-ins → Upload My Add-in** and point at the appropriate file in `manifests/`. Repeat once per app.

If that menu item isn't visible on your build of Office, fall back to the legacy procedure:

- **macOS:** copy `manifests/word.xml` to `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` (create the folder if missing). Same for Excel.
- **Windows:** put both manifests in a folder (e.g. `C:\OfficeAddins\`), then File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs → add the path → check **Show in Menu** → restart Office. Microsoft's full guide: [Sideload an Office Add-in on Windows](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins).

---

## Usage

1. Launch Office Claude (the tray app) before opening Word/Excel.
2. Open a document/workbook. Open the add-in panel from Insert → My Add-ins.
3. In the Setup tab, pick a workspace folder (the folder Claude treats as cwd). You'll get a confirm banner suggesting one based on your active doc's location.
4. (Optional) Add **context files** — folders or specific files Claude should consider as background.
5. Chat in the Chat tab. Use pinned presets (chips above the input) for one-click common operations.

### Default presets
- **Summarize this document** / **Outline this document**
- **Improve writing in selection** (Word) — preserves meaning, uses track changes
- **Fix typos and inconsistencies** — highlights and summarizes
- **Add comments on this section** — review without editing
- **Answer using my context files** — RAG over your added folders
- **Clear highlighting** — wipe every highlight in one click

---

## Troubleshooting

**Taskpane shows "Disconnected — retrying…"**
The daemon isn't running, or crashed. Open the tray menu → Open logs (`~/.claude/word-addin/daemon.log`).

**The picker comes up behind Word/Excel (macOS)**
Should be fixed by the focus-steal call, but if it persists, click Office Claude in the dock-less app switcher (⌘Tab) once before clicking the picker button.

**Auto-detect can't find my workspace**
If your folder has no `CLAUDE.md` or `.claude` marker and your folder structure doesn't match the suggestion heuristic, click "Pick…" in the suggest banner (or the **+ Add workspace** button) and choose the folder manually. A `CLAUDE.md` will be auto-created so the next open is silent.

**An MCP server I configured isn't visible to the agent**
At daemon start you should see `Loaded N MCP server(s) from ~/.claude.json: ...` followed by a preflight line confirming each HTTP server is reachable. If a server was down at daemon start, the SDK silently drops it for the lifetime of that session — restart the daemon.

**Context-file changes don't seem to take effect**
Saving any context entry restarts the agent session via `resume` so the new `CLAUDE.md` content loads. If you don't see the change, check that you're on the correct workspace (the topbar chip).

---

## File layout

```
.
├── app/
│   ├── main.mjs          Electron menu-bar shell; manages daemon child + IPC
│   ├── tray-icon.png     macOS template icon
│   └── tray-icon-win.ico (optional) Windows icon
├── daemon/
│   ├── index.mjs         Agent SDK loop, WS bridge, MCP forwarding
│   ├── bridge.mjs        WS server + tool-call protocol
│   ├── office-tools.mjs  Word + Excel tool definitions (zod schemas)
│   ├── workspace.mjs     Folder resolver + suggester + marker creator
│   ├── context.mjs       Per-workspace context-file list (in CLAUDE.md)
│   ├── sessions.mjs      Recent-workspace persistence
│   └── system-prompt.md  Office-add-in-specific system prompt
├── taskpane/
│   ├── shared/
│   │   ├── taskpane.js   Host-aware taskpane (Word + Excel branches)
│   │   └── styles.css
│   ├── word/index.html   Word manifest target
│   └── excel/index.html  Excel manifest target
├── manifests/
│   ├── word.xml
│   └── excel.xml
├── package.json
└── README.md
```

---

## Authentication

The daemon authenticates via the Claude Agent SDK, which by default reads your Claude Code OAuth credential from the system keychain. If you have Claude Code signed in to a Pro/Max account, the add-in just works.

If you'd rather use an API key (or you're not signed in to Claude Code), set `ANTHROPIC_API_KEY` in your environment before launching:

```bash
export ANTHROPIC_API_KEY=sk-ant-...
npm start
```

The SDK prefers `ANTHROPIC_API_KEY` over keychain OAuth when present.

> **Note on distribution.** This repo is open source — each user clones and runs it on their own machine with their own auth. That use is sanctioned. If you want to ship a hosted / packaged commercial product that uses Anthropic subscription OAuth on behalf of users, that's a separate conversation with Anthropic; the [Agent SDK overview](https://code.claude.com/docs/en/agent-sdk/overview) requires API keys (or partner approval) for distributed products.

---

## License

MIT. See [LICENSE](LICENSE).
