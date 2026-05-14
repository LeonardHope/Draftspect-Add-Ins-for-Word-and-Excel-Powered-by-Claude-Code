# Claude Code Add-Ins for Word and Excel — project guidance for Claude Code

Microsoft Office add-ins (Word + Excel) that wrap a local Claude Code daemon and expose it inside Office via Office.js task panes. Cross-platform (macOS + Windows). Forked from a private patent-drafting-specific build ("Patspect") and stripped to a general audience.

This is **not a product** — it's a public-eventually GitHub repo of integration code. Each user clones it and runs it on their own machine with their own Claude Code. The naming convention reflects that:
- npm package: `claude-code-office-addins`
- Per-host add-in DisplayName: `Claude Code for Word`, `Claude Code for Excel`
- Tray app: `Claude Code for Office`
- Repo title: `Claude Code Add-Ins for Word and Excel`

These are descriptive labels (the integration code IS for Claude Code), not a product brand. The Anthropic branding rule against using "Claude Code" in a *product name* doesn't apply because we're not shipping a product.

## Architecture (3 paragraphs)

The **daemon** (`daemon/`) is the brain. It wraps the Claude Agent SDK, hosts a WebSocket bridge for the taskpane, and registers an in-process MCP server (`office-tools.mjs`) that exposes Word + Excel tools. The daemon inherits the user's MCP servers from `~/.claude.json` and forwards them into the SDK session (the SDK doesn't read that file by default — see `feedback_sdk_does_not_load_claude_json`).

The **Electron shell** (`app/main.mjs`) is the tray app. It spawns the daemon as a child process over an IPC channel (used for native folder pickers and add-in install/uninstall), restarts it on crash, and surfaces status in a menu-bar / system-tray icon. The shell also handles **auto-sideload** (`app/sideload.mjs`) — on first run it copies manifests into Word's and Excel's wef/ folder on macOS, or registers a Trusted Catalog via the Windows registry. No manual XML drops or Trust Center configuration required.

The **taskpane** (`taskpane/`) is what shows up inside Word/Excel. One shared `taskpane.js` handles both hosts; the active host is detected via `Office.context.host` and the tool dispatcher routes `office_*` calls to Word handlers and `excel_*` calls to Excel handlers. The user picks a **workspace folder** (the agent's cwd, where its `CLAUDE.md` lives) and optionally adds **context files** — paths the agent reads on demand via standard `Read`/`Glob`/`Grep`.

## Key constraints

- **Auth.** Each user clones the repo and runs it on their own machine with their own Claude Code OAuth (or `ANTHROPIC_API_KEY`). That use is sanctioned. Distributing a packaged/hosted product that uses subscription OAuth on behalf of other users is **not** allowed without Anthropic partner approval — see `feedback_subscription_auth_only` and `reference_anthropic_april_2026_policy`. If this ever ships as a real product, switch to BYO API key.
- **Branding.** This is a public repo of integration code, not a product. Naming is descriptive ("Claude Code for Word", etc.) — that's allowed because we're not packaging a third-party product on top of Claude Code. If this ever DID become a packaged product, the "Claude Code" naming would need to be replaced. See `feedback_product_branding_compliance`.
- **Filesystem-write safety.** A `canUseTool` guard refuses `Write`/`Edit`/`MultiEdit` against `.docx`/`.docm`/`.xlsx`/`.xlsm` paths — the active doc is held by Office with unsaved changes, and a filesystem write would corrupt it. The agent must use `office_*` / `excel_*` tools instead.
- **Feature branches, not main.** Per the user's standing preference, every feature goes on a branch and lands via PR — no direct commits to main. (Once this repo is on GitHub.)
- **MCP forwarding.** The daemon must forward `~/.claude.json`'s `mcpServers` into the SDK session. The SDK doesn't read that file. If a server is unreachable at daemon startup, the SDK silently drops it for the session's lifetime — restart to retry. See `feedback_sdk_silently_drops_failed_mcp`.

## Useful commands

```bash
npm start              # Launch the Electron tray app (daemon + UI)
npm run dev            # Daemon only, in the terminal (for debugging)
node --check daemon/index.mjs taskpane/shared/taskpane.js app/main.mjs
```

## File layout

- `daemon/` — Node daemon, WS bridge, MCP plumbing, Office tool defs, workspace detection
- `daemon/system-prompt.md` — Office-add-in-specific append to the Claude Code preset prompt
- `app/main.mjs` — Electron menu-bar shell + daemon lifecycle
- `app/sideload.mjs` — macOS/Windows add-in install/uninstall
- `taskpane/shared/taskpane.js` — host-aware taskpane (Word + Excel branches)
- `taskpane/shared/styles.css` — taskpane styles
- `taskpane/word/index.html`, `taskpane/excel/index.html` — per-host entry points
- `manifests/word.xml`, `manifests/excel.xml` — Office Add-in manifests

## Where to read more

- `README.md` — user-facing install/usage/troubleshooting docs.
- Memories under `~/.claude/projects/-Users-leonard-Projects-Claude-Code-Add-Ins-for-Word-and-Excel/memory/` — non-obvious constraints, Office.js gotchas, SDK behavior notes carried over from the Patspect work.

## Related private repo

Patspect lives at `/Users/leonard/Projects/Claude Code Plugin for Word`. It's the patent-drafting-specific original. Do not assume Patspect changes flow back here automatically — this is a fork, not a shared package. If a fix applies to both, it needs to be applied in both places.
