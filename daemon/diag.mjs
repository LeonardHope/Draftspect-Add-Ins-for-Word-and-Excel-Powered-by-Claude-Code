// Opt-in diagnostic logger. Quiet by default; set CC_OFFICE_DEBUG=1 to
// surface per-session tool registration, the SDK's init tool list, and
// every tool the model calls — the breadcrumbs needed to debug live
// Office issues (which CI can't exercise).
//
//   CC_OFFICE_DEBUG=1 npm run dev
//
// Lines are prefixed "[diag]" so they're easy to grep in daemon.log.

export const DIAG = process.env.CC_OFFICE_DEBUG === "1";

export function diag(...args) {
  if (DIAG) console.log("[diag]", ...args);
}
