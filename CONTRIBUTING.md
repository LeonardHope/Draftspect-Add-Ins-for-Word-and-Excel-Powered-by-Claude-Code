# Contributing

This is a personal-use, run-it-yourself tool, not a distributed product. It
runs against **subscription auth** (your local Claude Code's keychain
credential), which Anthropic's Agent SDK terms restrict to personal use
absent a partner agreement — so it isn't packaged or published for general
installation. Contributions are welcome in that spirit: improvements you'd
run yourself.

It's a generalized fork of a private patent-drafting Word add-in
("Patspect"). Fixes flow across by hand, tracked in a local port file — see
"Shared with the upstream fork" below.

## Setup

- Node `>=18` (see `engines` in `package.json`). No build step — the daemon
  and taskpane run straight from source.
- `npm install`
- [Claude Code](https://docs.claude.com/en/docs/claude-code/overview)
  installed and signed in to a Pro/Max account (the daemon reads the same
  OAuth credential), or `ANTHROPIC_API_KEY` exported.

Run it:

- `npm start` — Electron tray shell + managed daemon (production-style).
- `npm run dev` — daemon directly in the terminal (for debugging).

Sideload the manifests into Word/Excel to exercise the add-in end to end —
the tray app does this for you on first run; manual fallback is in the
[README](README.md#install). Note the one-install-per-host GUID caveat
there.

## Workflow

- Branch per change off `main`; merge via PR. No direct-to-`main` commits.
- CI (`node --check` + `npm test` on Node 20/22 + `npm audit`) must be
  green before merge.
- Keep the diff scoped to the change. No drive-by refactors, no
  speculative abstractions, no backwards-compat shims for code that hasn't
  shipped to anyone.
- Default to surgical edits in the codebase, the same way the agent's
  system prompt steers it toward surgical document edits.

## Before opening a PR

- `npm test` passes, and `node --check` is clean on changed `.mjs`/`.js`
  (CI runs both, but check locally first).
- Manually exercise the affected path in Word **and** Excel where
  relevant — a clean parse and green tests verify code correctness, not
  feature correctness. If you couldn't test it in a host, say so
  explicitly in the PR.
- If the change touches code shared with the upstream Patspect fork, note
  it so the port tracker stays current.

## Architecture

See the [Developing](README.md#developing) section of the README for the
file layout (tray shell / daemon / bridge / host-aware taskpane) and the
test setup.

## Shared with the upstream fork

This repo was generalized from Patspect; the two are not a shared package.
When a Patspect change touches code that also lives here (taskpane editing
tools, daemon WS handlers, system-prompt rules, Office.js gotchas, MCP
plumbing), it's ported here by hand and recorded in a local-only
`PORT_TO_OFFICE_ADDINS.md` (gitignored). Patent-specific upstream changes
do not port.
