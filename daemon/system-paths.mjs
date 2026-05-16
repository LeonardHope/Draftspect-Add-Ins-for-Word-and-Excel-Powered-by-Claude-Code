// Single source of truth for the OS-managed `$HOME` children that must
// never be treated as a workspace folder.
//
// Previously duplicated (and drifted) across daemon/workspace.mjs,
// daemon/sessions.mjs, and app/main.mjs — one copy was unconditional
// while the other two were correctly macOS-gated. The canonical behavior
// is the macOS-gated one: on macOS, picking ~/Library (etc.) as a
// workspace is a real foot-gun (an Office doc opened from there would
// point the agent's cwd at a sandbox container); on Windows there is no
// equivalent confusable set.
//
// Computed per call (homedir() is cheap) rather than memoized at import:
// it stays correct if $HOME changes and is trivial to exercise under a
// fake HOME in tests, with no module-cache surprises.

import { homedir } from "node:os";
import { join } from "node:path";

function systemHomeChildren() {
  if (process.platform !== "darwin") return [];
  const h = homedir();
  return [
    join(h, "Library"),
    join(h, "Movies"),
    join(h, "Music"),
    join(h, "Pictures"),
    join(h, "Public"),
  ];
}

// True when `p` is one of those OS-managed children — i.e. it must NOT be
// offered or persisted as a workspace.
export function isSystemHomeChild(p) {
  return systemHomeChildren().includes(p);
}
