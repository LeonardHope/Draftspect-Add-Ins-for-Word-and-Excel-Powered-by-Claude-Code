// Cloud-mirror map: remember a manual "Change workspace" pick made while a
// cloud document (SharePoint / OneDrive / any HTTPS URL) is active, so the
// next time a doc URL under that same cloud prefix opens, the workspace
// auto-follows to the local sync folder.
//
// Storage: ~/.draftspect/cloud-mirrors.json — a small JSON map of
// {urlPrefix → localFolder}. Both sides are full strings; matching is a
// longest-prefix scan, so finer-grained later picks naturally win over
// broader earlier ones.
//
// Learning trigger: a `set_cwd` with explicitPick=true while the pane's
// activeDoc is a non-file URL. The URL prefix is inferred from the leaf
// segment of the picked local folder (the SharePoint and OneDrive sync
// folders mirror folder names verbatim), with a fallback to "the doc's
// parent folder" when the leaf can't be located in the URL.

import { readFile, writeFile, mkdir, stat } from "node:fs/promises";
import { homedir } from "node:os";
import { join, dirname, basename } from "node:path";

const STORE_DIR = join(homedir(), ".draftspect");
const STORE_PATH = join(STORE_DIR, "cloud-mirrors.json");

// Decode percent-encoding but otherwise keep the URL intact (incl. spaces
// and other punctuation that SharePoint hands out verbatim).
function decodeUrl(u) {
  try {
    return decodeURIComponent(String(u));
  } catch {
    return String(u);
  }
}

// Find the last occurrence of "/seg/" in the URL path, where `seg` is the
// leaf of the user's locally-picked folder. Returns the URL up through the
// segment + trailing slash, or null if not found.
function urlPrefixFromLeaf(docUrl, leaf) {
  if (!leaf) return null;
  const u = decodeUrl(docUrl);
  const needle = "/" + leaf + "/";
  const idx = u.lastIndexOf(needle);
  if (idx < 0) return null;
  return u.slice(0, idx + needle.length);
}

// Fallback: the URL up to the last "/" before the filename. Maps "any doc
// in the same cloud folder" to the picked local folder — narrower than the
// leaf-match path, but always works.
function urlPrefixFromDocParent(docUrl) {
  const u = decodeUrl(docUrl);
  const slash = u.lastIndexOf("/");
  if (slash < 0) return null;
  return u.slice(0, slash + 1);
}

export function isCloudUrl(s) {
  return typeof s === "string" && /^https?:\/\//i.test(s);
}

export async function loadMirrors() {
  try {
    const raw = await readFile(STORE_PATH, "utf8");
    const obj = JSON.parse(raw);
    return obj && typeof obj === "object" ? obj : {};
  } catch {
    return {};
  }
}

async function saveMirrors(map) {
  await mkdir(STORE_DIR, { recursive: true });
  await writeFile(STORE_PATH, JSON.stringify(map, null, 2) + "\n", "utf8");
}

// Record a (docUrl → localFolder) hint. Uses leaf-match if possible, falls
// back to the doc's parent URL otherwise. Idempotent — overwrites prior
// entries at the same URL prefix.
export async function rememberCloudMirror(docUrl, localFolder) {
  if (!isCloudUrl(docUrl) || !localFolder) return null;
  const leaf = basename(localFolder);
  const prefix = urlPrefixFromLeaf(docUrl, leaf) ?? urlPrefixFromDocParent(docUrl);
  if (!prefix) return null;
  const map = await loadMirrors();
  map[prefix] = localFolder;
  await saveMirrors(map);
  return { urlPrefix: prefix, localFolder };
}

// Look up the workspace folder for a given doc URL. Picks the longest
// prefix that's a prefix of the URL, then builds the doc's *folder* by
// appending whatever subpath sits between that prefix and the document
// filename. Returns null if nothing matches or if the computed folder
// doesn't actually exist on disk (don't pretend a stale mapping is live).
export async function lookupCloudMirror(docUrl) {
  if (!isCloudUrl(docUrl)) return null;
  const u = decodeUrl(docUrl);
  const map = await loadMirrors();
  let bestPrefix = null;
  let bestLocal = null;
  for (const [prefix, local] of Object.entries(map)) {
    if (u.startsWith(prefix) && (bestPrefix === null || prefix.length > bestPrefix.length)) {
      bestPrefix = prefix;
      bestLocal = local;
    }
  }
  if (!bestPrefix) return null;
  // Path from the matched prefix to the doc, minus the filename. Empty
  // when the doc sits directly inside the mapped folder.
  const tail = u.slice(bestPrefix.length);
  const tailFolder = dirname(tail);
  const folder = !tailFolder || tailFolder === "." ? bestLocal : join(bestLocal, tailFolder);
  try {
    const s = await stat(folder);
    if (!s.isDirectory()) return null;
  } catch {
    return null;
  }
  return folder;
}
