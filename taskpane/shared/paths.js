// Path-shape helpers for the browser side of the add-in. Office hands us
// paths in a few shapes (POSIX, Windows native, file:// URLs, the
// "/C:/..." remnant Office.context.document.url sometimes returns); we
// normalize everything to forward-slash form for comparison.
//
// All exports are pure — no DOM, no Office.js — so this module is also
// unit-testable from Node.

function _isWindowsPath(p) {
  return typeof p === "string" && /^[a-zA-Z]:[\\/]/.test(p);
}

// Convert any of macOS path, Windows path, file:// URL, or "/C:/..." URL-stripped
// remnant into a uniform forward-slash representation. Strips trailing slashes.
function _normalizePathForCompare(p) {
  if (!p) return "";
  let s = String(p);
  // file:// URLs: drop scheme + decode. "file:///C:/x/y" → "/C:/x/y"; on
  // POSIX "file:///Users/x" → "/Users/x". The Windows-shaped result has a
  // spurious leading slash before the drive letter; fix that below.
  if (/^file:\/\//i.test(s)) {
    s = s.replace(/^file:\/\//i, "");
    try {
      s = decodeURIComponent(s);
    } catch {
      /* leave as-is */
    }
  }
  // "/C:/x" → "C:/x" (Office.context.document.url comes in this shape).
  s = s.replace(/^\/([a-zA-Z]):/, "$1:");
  // Backslash → forward slash for consistent comparison.
  s = s.replace(/\\/g, "/");
  // Trim trailing separator(s) so containment math is uniform.
  s = s.replace(/\/+$/, "");
  return s;
}

// Path containment check that handles the /Workspace-10 vs /Workspace-1 trap
// (`startsWith` alone returns true for both) and is correct on Windows
// (drive letters, backslashes, case-insensitive filesystem). Both inputs are
// normalized to forward-slash form; comparison is case-insensitive when
// either side starts with a drive letter.
export function isInOrUnder(child, parent) {
  if (!child || !parent) return false;
  const c = _normalizePathForCompare(child);
  const p = _normalizePathForCompare(parent);
  if (!c || !p) return false;
  const caseInsensitive = _isWindowsPath(c) || _isWindowsPath(p);
  const cc = caseInsensitive ? c.toLowerCase() : c;
  const pp = caseInsensitive ? p.toLowerCase() : p;
  return cc === pp || cc.startsWith(pp + "/");
}

// Doc URL (Office.context.document.url) → directory path, normalized for
// comparison with the current workspace cwd. Returns "" if the URL doesn't
// refer to a filesystem location (e.g. SharePoint / OneDrive cloud).
export function docDirFromActiveUrl(activeDocUrl) {
  if (!activeDocUrl) return "";
  // SharePoint / OneDrive cloud URLs are not filesystem paths; skip them so
  // we don't mis-detect "mismatch" for cloud-hosted docs.
  if (/^https?:\/\//i.test(activeDocUrl)) return "";
  const docPath = _normalizePathForCompare(activeDocUrl);
  // Strip the filename. _normalizePathForCompare gave us forward slashes.
  const idx = docPath.lastIndexOf("/");
  return idx > 0 ? docPath.slice(0, idx) : "";
}
