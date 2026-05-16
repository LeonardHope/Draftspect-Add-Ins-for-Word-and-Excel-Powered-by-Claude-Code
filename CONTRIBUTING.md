# Contributing

This is a single-user tool you run on your own machine, not a distributed
product. It runs against **subscription auth** (your local Claude Code's
keychain credential), which Anthropic's Agent SDK terms restrict to
personal use absent a partner agreement — so it isn't packaged or
published for general installation. Contributions are welcome in that
spirit: improvements you'd run yourself.

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
- CI (`node --check` + `npm test` on Node 20/22, `npm audit`, and
  `prettier --check`) must be green before merge. Run `npm run format`
  before committing.
- Keep the diff scoped to the change. No drive-by refactors, no
  speculative abstractions, no backwards-compat shims for code that hasn't
  shipped to anyone.
- Default to surgical edits in the codebase, the same way the agent's
  system prompt steers it toward surgical document edits.

## Before opening a PR

- `npm test` passes, `node --check` is clean on changed `.mjs`/`.js`, and
  `npm run format:check` passes (CI runs all three, but check locally
  first).
- Manually exercise the affected path in Word **and** Excel where
  relevant — a clean parse and green tests verify code correctness, not
  feature correctness. If you couldn't test it in a host, say so
  explicitly in the PR.

## Architecture

See the [Developing](README.md#developing) section of the README for the
file layout (tray shell / daemon / bridge / host-aware taskpane) and the
test setup.
