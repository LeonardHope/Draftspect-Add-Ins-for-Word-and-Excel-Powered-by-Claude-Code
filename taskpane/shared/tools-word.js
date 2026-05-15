/* global Word */
//
// Word tool implementations + Word-specific helpers.
//
// Each `tool*` function is invoked from the dispatcher in taskpane.js when a
// matching `office_*` tool_call arrives from the daemon. They all run inside
// Word.run(...) so Office.js manages the request context lifecycle. The
// helpers above (paragraph addressing, withTrackChanges, snapshot, heading
// detection) are private to this module — taskpane.js doesn't need them.

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

async function getParagraphsWithIds(context) {
  const paragraphs = context.document.body.paragraphs;
  try {
    paragraphs.load("items/text, items/style, items/uniqueLocalId");
    await context.sync();
    if (paragraphs.items.length > 0 && !paragraphs.items[0].uniqueLocalId) {
      throw new Error("uniqueLocalId not populated");
    }
    return { paragraphs, idMode: "uniqueLocalId" };
  } catch {
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
// Heading / section helpers
// ---------------------------------------------------------------------------
function headingLevel(style) {
  // Returns the heading level if the style name *contains* "Heading <N>".
  // Many organizations use prefixed style names like "Acme-Heading 2" or
  // "Firm-Heading-1" — we want to recognize those as headings too, not just
  // Word's built-in "Heading 1" / "Heading 2".
  if (!style) return null;
  const m = /heading[\s_-]*(\d)/i.exec(style);
  return m ? parseInt(m[1], 10) : null;
}

// Build the "view" the agent sees: list of {id, style, text}.
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

// Infer the body-text style for the section that contains the insert
// position `fromIdx`. Algorithm:
//
//   1. Find the section we're in by walking BACKWARD from fromIdx-1 to the
//      nearest heading paragraph. That heading is the top of the section.
//   2. From the heading, walk FORWARD looking for the first body paragraph
//      in this section. Stop at the NEXT heading — that means the section
//      has no body content above the insert (rare — insert is right after
//      the heading, into an empty section).
//   3. Return that paragraph's style. That's the canonical body style for
//      the section.
//
// The old implementation walked forward from `fromIdx` first, which leaked
// the *next* section's body style when the insert was at the end of a
// section (the inferrer skipped past the next heading instead of stopping
// at it). Templates that vary the body style between sections made that
// visibly wrong (no first-line indent, wrong alignment on inserted text).
//
// If there's no heading above the insert (pre-heading content / very top
// of the doc), fall back to the first non-heading paragraph in the doc.
//
// Returns the INDEX of the canonical body paragraph (or -1). inferBodyStyle
// is a thin wrapper that returns that paragraph's style name; callers that
// also clone the paragraph's *direct* formatting (indent, alignment,
// spacing — which many templates apply manually on top of a plain style
// like "Normal", so the style name alone doesn't carry them) use this to
// grab the live Word.Paragraph.
function findBodyReferenceIndex(snapshot, fromIdx) {
  // Step 1: find the section's heading by walking backward.
  let sectionStart = -1;
  for (let i = fromIdx - 1; i >= 0; i--) {
    if (snapshot[i] && headingLevel(snapshot[i].style) !== null) {
      sectionStart = i;
      break;
    }
  }

  // Step 2: from the heading, the first body paragraph in THIS section.
  // Stop at the next heading.
  if (sectionStart !== -1) {
    for (let i = sectionStart + 1; i < snapshot.length; i++) {
      if (!snapshot[i]) continue;
      if (headingLevel(snapshot[i].style) !== null) break;
      return i;
    }
  }

  // Fallback: first non-heading paragraph in the doc (pre-heading content).
  for (let i = 0; i < snapshot.length; i++) {
    if (snapshot[i] && headingLevel(snapshot[i].style) === null) return i;
  }
  return -1;
}

function inferBodyStyle(snapshot, fromIdx) {
  const idx = findBodyReferenceIndex(snapshot, fromIdx);
  return idx === -1 ? null : (snapshot[idx].style || null);
}

// Paragraph-format properties cloned from a reference body paragraph onto
// inserted ones. Many templates set these directly (not via the named
// style), so a style-name copy alone leaves inserts left-aligned with no
// first-line indent and no inter-paragraph spacing.
const CLONED_PARA_FORMAT_PROPS = [
  "alignment", "firstLineIndent", "leftIndent", "rightIndent",
  "lineSpacing", "spaceBefore", "spaceAfter",
  "lineUnitBefore", "lineUnitAfter",
];

// Queue a load of the format props off a reference paragraph. Call before
// the next context.sync().
function loadParagraphFormat(refPara) {
  refPara.load(CLONED_PARA_FORMAT_PROPS.join(", "));
}

// Snapshot a synced reference paragraph's format values into a plain
// object. Decouples the captured formatting from the live proxy —
// necessary for office_replace_section, where the reference paragraph is
// deleted before the new paragraphs are inserted.
function snapshotParagraphFormat(refPara) {
  const out = {};
  for (const prop of CLONED_PARA_FORMAT_PROPS) {
    try {
      const v = refPara[prop];
      if (v !== undefined && v !== null) out[prop] = v;
    } catch { /* prop unsupported on this build */ }
  }
  return out;
}

// Apply a snapshotted format object onto a target paragraph.
function applyParagraphFormat(targetPara, fmt) {
  if (!fmt) return;
  for (const prop of CLONED_PARA_FORMAT_PROPS) {
    const v = fmt[prop];
    if (v !== undefined && v !== null) {
      try { targetPara[prop] = v; } catch { /* prop unsupported on this build */ }
    }
  }
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
export async function toolGetSelection() {
  return await Word.run(async (context) => {
    const sel = context.document.getSelection();
    sel.load("text");
    const selParas = sel.paragraphs;

    // Try to load uniqueLocalId directly on the selected paragraphs — that
    // gives us stable IDs without any text-matching against the full doc
    // (which is unsafe whenever the doc has repeated/boilerplate paragraphs).
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
      // an index by text-matching — duplicate paragraphs in any doc make
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
export async function toolReadParagraphs({ ids, heading_section, range, preview }) {
  return await Word.run(async (context) => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const snapshot = snapshotParagraphs(paragraphs, idMode);

    let picked;
    // Truncation is on by default in the no-args ("dump the whole doc")
    // mode — that's an orientation pass, not a content read. Targeted forms
    // (`ids`, `heading_section`, `range`) always return full text. The agent
    // can also explicitly force `preview: false` in the no-args mode for a
    // full dump (heavy, but escape-hatch valid).
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
      // Default: outline-ish view. Word documents in the wild use varied
      // heading conventions (built-in Heading N, custom style names, all-caps
      // direct formatting), so we don't pre-filter — return every paragraph
      // with its style and a truncated text preview. The agent picks out the
      // headings.
      picked = snapshot;
      truncate = preview !== false;
    }

    // Preview length: long enough to surface most in-paragraph references
    // (numbers, defined terms, citations) while still bounding the response
    // size on large docs. Was 117 originally — too aggressive; the agent
    // treated previews as full reads and missed mid-paragraph content.
    const PREVIEW_LEN = 500;

    return {
      paragraphs: picked.map(p => {
        const full = p.text;
        const isTruncated = truncate && full.length > PREVIEW_LEN;
        return {
          id: p.id,
          style: p.style,
          text: isTruncated ? full.slice(0, PREVIEW_LEN - 1) + "…" : full,
          // Structured truncation flag so the agent doesn't have to detect
          // "…" suffixes. When `truncated` is true, the paragraph has more
          // content; re-read it via `ids: [<id>]` (always full text) to get
          // it. `full_length` lets the agent decide if a re-read is worth
          // the round-trip.
          ...(isTruncated ? { truncated: true, full_length: full.length } : {}),
        };
      }),
      total_in_doc: snapshot.length,
      addressing: idMode,
      // When truncation is active across the response, surface it at the
      // top level too. Cheap signal for the agent's next-action decision.
      ...(truncate ? { preview_mode: true, preview_limit: PREVIEW_LEN } : {}),
    };
  });
}

// ---------------------------------------------------------------------------
// Tool: office_insert_paragraphs
// ---------------------------------------------------------------------------
export async function toolInsertParagraphs({ after, content, track_changes, style_per_para, provenance_comment }) {
  if (!content || content.length === 0) throw new Error("content must be non-empty");

  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);
    const snapshot = snapshotParagraphs(paragraphs, idMode);

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
    const useInferred = !(style_per_para && style_per_para.length === content.length);
    const refIdx = useInferred ? findBodyReferenceIndex(snapshot, anchorIdx + 1) : -1;
    const inferredStyle = (refIdx === -1) ? null : (snapshot[refIdx].style || null);

    // Load the reference body paragraph's *direct* formatting (indent,
    // alignment, spacing). Templates often apply these manually, so
    // copying the style name alone leaves inserts mis-formatted.
    let refFmt = null;
    if (refIdx !== -1) {
      const refPara = paragraphs.items[refIdx];
      loadParagraphFormat(refPara);
      await context.sync();
      refFmt = snapshotParagraphFormat(refPara);
    }

    let cursor = anchor;
    const inserted = [];
    for (let i = 0; i < content.length; i++) {
      const p = cursor.insertParagraph(content[i], Word.InsertLocation.after);
      const explicit = style_per_para && style_per_para[i];
      if (explicit) p.style = explicit;
      else if (inferredStyle) p.style = inferredStyle;
      if (refFmt && !explicit) applyParagraphFormat(p, refFmt);
      inserted.push(p);
      cursor = p;
    }

    if (provenance_comment && inserted.length > 0) {
      inserted[0].getRange().insertComment(provenance_comment);
    }

    await context.sync();

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
export async function toolReplaceParagraphs({ ids, content, track_changes, style_per_para, provenance_comment }) {
  if (!ids || ids.length === 0) throw new Error("ids must be non-empty");
  if (!content || content.length === 0) throw new Error("content must be non-empty");
  if (ids.length !== content.length) {
    throw new Error(`ids.length (${ids.length}) must equal content.length (${content.length}). To grow or shrink a section, use office_insert_paragraphs or office_replace_section.`);
  }

  return await Word.run(async (context) => {
    return await withTrackChanges(context, track_changes, async () => {
    const { paragraphs, idMode } = await getParagraphsWithIds(context);

    const targets = ids.map(id => {
      const idx = findIndexById(paragraphs, id, idMode);
      if (idx === -1) {
        throw new Error(`Paragraph not found: ${id}`);
      }
      return { id, idx, paragraph: paragraphs.items[idx] };
    });

    // `insertText(..., "Replace")` swaps the text content while keeping the
    // paragraph element (so style is preserved unless we override below).
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
export async function toolReplaceText({ paragraph_ids, find, replace, match_case, whole_word, track_changes }) {
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

    const targets = [];
    for (const id of paragraph_ids) {
      const idx = findIndexById(paragraphs, id, idMode);
      if (idx === -1) {
        throw new Error(`Paragraph not found: ${id}`);
      }
      targets.push({ id, paragraph: paragraphs.items[idx] });
    }

    const probes = targets.map(t => {
      const results = t.paragraph.search(find, {
        matchCase: !!match_case,
        matchWholeWord: !!whole_word,
      });
      results.load("items");
      return { id: t.id, results };
    });
    await context.sync();

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
export async function toolReplaceSection({ heading, content, track_changes, style_per_para, provenance_comment }) {
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

    // Capture the existing section's body style AND direct formatting
    // BEFORE deleting, so inserted paragraphs match. Without this, new
    // paragraphs inherit the heading's style (centered/bold/large) and
    // lose any manually-applied indent/justify/spacing.
    const refIdx = (style_per_para && style_per_para.length === content.length)
      ? -1
      : findBodyReferenceIndex(snapshot, startIdx + 1);
    const bodyStyle = (refIdx === -1) ? null : (snapshot[refIdx].style || null);
    let refFmt = null;
    if (refIdx !== -1) {
      const refPara = paragraphs.items[refIdx];
      loadParagraphFormat(refPara);
      await context.sync();
      refFmt = snapshotParagraphFormat(refPara);
    }

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
      if (refFmt && !explicit) applyParagraphFormat(p, refFmt);
      inserted.push(p);
      cursor = p;
    }

    if (provenance_comment && inserted.length > 0) {
      inserted[0].getRange().insertComment(provenance_comment);
    }

    await context.sync();

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

export async function toolHighlight({ targets }) {
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
export async function toolClearHighlights(args) {
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
export async function toolClearComments(args) {
  const { paragraph_ids, heading_section, all } = args || {};

  return await Word.run(async (context) => {
    const comments = context.document.body.getComments();
    comments.load("items");
    await context.sync();

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
export async function toolAddComment({ paragraph_id, query, text }) {
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
      target = searchResults.items[0];
    } else {
      target = paragraph.getRange();
    }

    target.insertComment(text);
    await context.sync();

    return { paragraph_id, anchored_on: query || "whole_paragraph", comment_added: true };
  });
}
