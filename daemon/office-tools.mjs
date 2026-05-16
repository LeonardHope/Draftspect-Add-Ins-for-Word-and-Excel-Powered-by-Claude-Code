import { createSdkMcpServer, tool } from "@anthropic-ai/claude-agent-sdk";
import { z } from "zod";
import { diag } from "./diag.mjs";

// Wrap a bridge tool result for MCP. Handlers return {content: [...]}.
function asMcpResult(result, { isError = false } = {}) {
  const text = typeof result === "string" ? result : JSON.stringify(result, null, 2);
  return { content: [{ type: "text", text }], ...(isError ? { isError: true } : {}) };
}

function asMcpError(err) {
  return asMcpResult(`Error: ${err?.message ?? String(err)}`, { isError: true });
}

/**
 * Build the office-bridge MCP server. Tools forward calls over the WS bridge
 * to the taskpane, which executes them in Office.js and returns results.
 *
 * `host` scopes which tool family is registered:
 *   - "word"  → only the office_* tools
 *   - "excel" → only the excel_* tools
 *   - null    → both (used before a taskpane has said hello — at that
 *               point nothing can run anyway; the session is re-narrowed
 *               to the real host on first hello)
 *
 * Registering only the active host's family (instead of both, then
 * rejecting wrong-host calls at dispatch) means the agent never sees a
 * tool it can't use — the wrong-host class of mistake is structurally
 * impossible, not just caught.
 *
 * @param {{ callTaskpaneTool: (name: string, args: object) => Promise<any> }} bridge
 * @param {"word"|"excel"|null} host
 */
export function createOfficeBridgeMcp(bridge, host = null) {
  // Route to this session's host pane explicitly (the tool-name prefix is
  // the bridge's fallback). host is null only in the degraded pre-bind
  // case; the prefix still routes correctly there.
  const call = (name, args) => bridge.callTaskpaneTool(name, args, host);

  const office_get_selection = tool(
    "office_get_selection",
    "Get the user's current selection in the Word document. Returns selected text, the paragraph IDs the selection covers, and intra-paragraph character ranges. Call this whenever the user refers to 'this', 'here', 'the selection', or asks to edit existing content without specifying location.",
    {},
    async () => {
      try {
        const r = await call("office_get_selection", {});
        return asMcpResult(r);
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const office_read_paragraphs = tool(
    "office_read_paragraphs",
    "Read paragraphs from the active Word document. Specify exactly one of `ids` (paragraph IDs), `heading_section` (a heading text — returns everything under that heading until the next same-or-higher-level heading), or `range` (an [start_index, end_index] paragraph range). With NO arguments, returns every paragraph in the doc as a **preview** truncated to ~500 characters each — use this for orientation only. When a preview paragraph has `truncated: true` (and a `full_length` field), the actual paragraph is longer; re-read it via `ids: [<id>]` (full text, no truncation) before quoting or analyzing in-paragraph content. To force a full no-args dump of the whole doc, pass `preview: false` (heavy — only use for short docs).",
    {
      ids: z
        .array(z.string())
        .optional()
        .describe("Paragraph IDs to read. Always returns full text (no truncation)."),
      heading_section: z
        .string()
        .optional()
        .describe("Read all paragraphs under this heading. Full text."),
      range: z
        .tuple([z.number().int(), z.number().int()])
        .optional()
        .describe("[start, end) paragraph indices. Full text."),
      preview: z
        .boolean()
        .optional()
        .describe(
          "When calling with no other args: defaults to true (truncated preview). Set to false for a full-text dump of every paragraph.",
        ),
    },
    async (args) => {
      try {
        const r = await call("office_read_paragraphs", args);
        return asMcpResult(r);
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const office_insert_paragraphs = tool(
    "office_insert_paragraphs",
    "Insert new paragraphs into the active Word document. Specify `after` as either { id: <paragraph_id> } to insert after a specific paragraph, or { heading: <heading_text> } to insert at the top of that section (immediately after the heading). Set `track_changes: true` to insert as a tracked-change for review. Leave `style_per_para` unset to inherit style from surrounding context.",
    {
      after: z
        .union([z.object({ id: z.string() }), z.object({ heading: z.string() })])
        .describe("Anchor for the insertion."),
      content: z.array(z.string()).describe("One string per new paragraph."),
      track_changes: z.boolean().optional().default(false),
      style_per_para: z
        .array(z.string())
        .optional()
        .describe("Optional: a Word style name per paragraph. Length must match content."),
      provenance_comment: z
        .string()
        .optional()
        .describe(
          "Optional: add a Word comment on the first inserted paragraph (e.g. source citation).",
        ),
    },
    async (args) => {
      try {
        const r = await call("office_insert_paragraphs", args);
        return asMcpResult(r);
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const office_replace_paragraphs = tool(
    "office_replace_paragraphs",
    "Replace the entire text of specific paragraphs, 1:1 by ID. **Use this ONLY when the whole paragraph is genuinely being rewritten from scratch** — the user said 'rewrite p47' or 'redraft this paragraph.' For anything narrower (changing a word, fixing a phrase, renaming a term, adjusting one sentence inside an otherwise-intact paragraph), use `office_replace_text` instead — never re-emit a whole paragraph to change part of it. The paragraph's existing style is preserved unless you pass style_per_para. Pass track_changes: true so the user can review the rewrite.",
    {
      ids: z.array(z.string()).describe("Paragraph IDs to replace, in order."),
      content: z
        .array(z.string())
        .describe("Replacement text — one string per ID. Length must equal ids.length."),
      track_changes: z.boolean().optional().default(false),
      style_per_para: z
        .array(z.string())
        .optional()
        .describe("Optional: a Word style name per paragraph. Length must match content."),
      provenance_comment: z
        .string()
        .optional()
        .describe(
          "Optional: add a Word comment on the first replaced paragraph (e.g. source citation).",
        ),
    },
    async (args) => {
      try {
        const r = await call("office_replace_paragraphs", args);
        return asMcpResult(r);
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const office_highlight = tool(
    "office_highlight",
    "Highlight text in the active Word document, color-coded by severity. Use for review and QC workflows where you want the user to *see* findings directly in the doc rather than hunt through a chat list. Pass an array of targets — each highlights either the whole paragraph (no `query`) or specific text within it (a `query` substring). Severities map to highlight colors: error→red, warning→yellow (default), info→turquoise, uncertain→pink. Pair with `office_add_comment` to attach an explanation to each highlight so the user knows what's wrong and how to fix it.",
    {
      targets: z
        .array(
          z.object({
            paragraph_id: z
              .string()
              .describe("Paragraph ID (e.g., 'p47') containing the text to highlight."),
            query: z
              .string()
              .optional()
              .describe(
                "Optional text to find and highlight within the paragraph. If absent, the whole paragraph is highlighted.",
              ),
            severity: z
              .enum(["error", "warning", "info", "uncertain"])
              .optional()
              .default("warning")
              .describe(
                "error=red (definite problems), warning=yellow (should-fix), info=turquoise (style notes), uncertain=pink (agent uncertainty / needs human review).",
              ),
          }),
        )
        .describe("Highlights to apply, processed in one batch."),
    },
    async (args) => {
      try {
        const r = await call("office_highlight", args);
        return asMcpResult(r);
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const office_clear_highlights = tool(
    "office_clear_highlights",
    "Remove highlight formatting from text in the active Word document. Specify exactly one of: `paragraph_ids` (clear highlights from these specific paragraphs), `heading_section` (clear an entire section by heading), or `all: true` (clear every highlight in the document). Call this when the user has addressed prior findings.",
    {
      paragraph_ids: z
        .array(z.string())
        .optional()
        .describe("Clear highlights from these specific paragraphs."),
      heading_section: z
        .string()
        .optional()
        .describe("Clear highlights from this entire section (matched by heading text)."),
      all: z.boolean().optional().describe("If true, clear every highlight in the document."),
    },
    async (args) => {
      try {
        const r = await call("office_clear_highlights", args);
        return asMcpResult(r);
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const office_clear_comments = tool(
    "office_clear_comments",
    "Remove comments from the active Word document. Specify exactly one of: `paragraph_ids` (clear comments anchored on these paragraphs), `heading_section` (clear comments in this section), or `all: true` (clear every comment in the doc). Useful when starting a fresh review pass or after addressing prior findings.",
    {
      paragraph_ids: z
        .array(z.string())
        .optional()
        .describe("Clear comments anchored on these specific paragraphs."),
      heading_section: z
        .string()
        .optional()
        .describe("Clear comments in this entire section (matched by heading text)."),
      all: z.boolean().optional().describe("If true, clear every comment in the document."),
    },
    async (args) => {
      try {
        const r = await call("office_clear_comments", args);
        return asMcpResult(r);
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const office_add_comment = tool(
    "office_add_comment",
    "Add a Word comment anchored to a paragraph, optionally on a specific text match within it. Use to attach explanations to highlights during QC, cite sources when drafting from disclosure documents, or flag uncertainty. Comments appear in Word's review pane; the user can resolve or reply.",
    {
      paragraph_id: z.string().describe("Paragraph ID (e.g., 'p47') to anchor the comment on."),
      query: z
        .string()
        .optional()
        .describe(
          "Optional text within the paragraph to anchor on. If absent, the comment anchors on the whole paragraph.",
        ),
      text: z.string().describe("Comment body."),
    },
    async (args) => {
      try {
        const r = await call("office_add_comment", args);
        return asMcpResult(r);
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const office_replace_text = tool(
    "office_replace_text",
    'Surgically replace specific text *within* one or more paragraphs without rewriting the rest of the paragraph. **THIS IS THE DEFAULT WRITE TOOL.** Use it for any sub-paragraph edit: changing a single word, fixing a phrase, correcting a sentence, renaming a term across many paragraphs, etc. A single user request can resolve to many `office_replace_text` calls — that is the preferred shape, not one `office_replace_paragraphs` or `office_replace_section` re-emitting everything. Preserves all surrounding text and the paragraph\'s style verbatim. Pass `replace: ""` to delete the matched text. Pair with `track_changes: true` so the user can review.',
    {
      paragraph_ids: z
        .array(z.string())
        .describe(
          "Paragraph IDs to operate on. Required — scopes the search to specific paragraphs.",
        ),
      find: z
        .string()
        .describe("Exact text to find. The search is bounded to within each listed paragraph."),
      replace: z
        .string()
        .describe('Replacement text. Use "" (empty string) to delete the matched text.'),
      match_case: z.boolean().optional().default(false),
      whole_word: z.boolean().optional().default(false),
      track_changes: z
        .boolean()
        .optional()
        .default(true)
        .describe(
          "Default true — surgical edits should virtually always be tracked so the user can review.",
        ),
    },
    async (args) => {
      try {
        const r = await call("office_replace_text", args);
        return asMcpResult(r);
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const office_replace_section = tool(
    "office_replace_section",
    "Replace an entire section of the active Word document, identified by its heading text. Finds the heading and replaces everything from immediately after the heading until the next same-or-higher-level heading. **Use ONLY for explicit whole-section rewrites** where the user has clearly asked to throw the existing content away and start over: 'redraft the entire Background,' 'start the Summary over from scratch.' For any request that would leave most paragraphs of the section unchanged, do NOT use this tool — use `office_replace_text` for the actual changes (or `office_replace_paragraphs` for the genuinely-rewritten paragraphs). Re-emitting unchanged paragraphs just to keep them blows up the diff and risks destroying user edits. Set `track_changes: true` when revising existing content; leave false (default) when drafting from scratch.",
    {
      heading: z.string().describe("The heading text identifying the section."),
      content: z.array(z.string()).describe("One string per new paragraph."),
      track_changes: z.boolean().optional().default(false),
      style_per_para: z.array(z.string()).optional(),
      provenance_comment: z.string().optional(),
    },
    async (args) => {
      try {
        const r = await call("office_replace_section", args);
        return asMcpResult(r);
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  // ---- Excel tools ---------------------------------------------------------
  // Same WS-bridge round-trip pattern as the Word tools above. The Excel
  // taskpane registers handlers for these names and executes them inside
  // Excel.run(...) via Office.js.

  const excel_get_selected_range = tool(
    "excel_get_selected_range",
    "Return the user's current selection in the active Excel workbook. Includes the range address (e.g. \"Sheet1!B2:D5\"), the values as a 2D array, and the worksheet name. Call whenever the user refers to 'this', 'these cells', 'the selection', or asks to edit existing content without specifying location.",
    {},
    async () => {
      try {
        return asMcpResult(await call("excel_get_selected_range", {}));
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const excel_list_sheets = tool(
    "excel_list_sheets",
    "List every worksheet in the active workbook with its name, position, and used-range address. Call this to orient yourself before reading or writing data.",
    {},
    async () => {
      try {
        return asMcpResult(await call("excel_list_sheets", {}));
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const excel_read_range = tool(
    "excel_read_range",
    'Read values from an Excel range. Pass `address` in A1 notation, optionally sheet-qualified (e.g. "A1:C10" or "Sheet2!A1:C10"). To read an entire sheet\'s used range, pass `sheet` and omit `address`. Returns a 2D values array plus the resolved address.',
    {
      address: z.string().optional().describe("A1-notation range, optionally sheet-qualified."),
      sheet: z
        .string()
        .optional()
        .describe(
          "Worksheet name; defaults to the active sheet. Used when address is omitted (reads the whole used range) or to disambiguate.",
        ),
    },
    async (args) => {
      try {
        return asMcpResult(await call("excel_read_range", args));
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const excel_write_range = tool(
    "excel_write_range",
    "Write a 2D array of values into an Excel range. The shape of `values` must match the address dimensions (rows × cols). Pass `address` in A1 notation, optionally sheet-qualified.",
    {
      address: z.string().describe("Target range in A1 notation."),
      values: z
        .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
        .describe("2D array of values. Outer = rows, inner = columns."),
      sheet: z.string().optional().describe("Worksheet name; defaults to the active sheet."),
    },
    async (args) => {
      try {
        return asMcpResult(await call("excel_write_range", args));
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const excel_find_value = tool(
    "excel_find_value",
    "Find cells whose value matches a string (case-insensitive substring match by default) across one or all sheets. Returns the address and current value of each match.",
    {
      query: z.string().describe("Substring to search for."),
      sheet: z.string().optional().describe("Worksheet to search; omit to search every sheet."),
      match_case: z.boolean().optional().default(false),
      whole_cell: z
        .boolean()
        .optional()
        .default(false)
        .describe("If true, match only cells whose entire value equals the query."),
    },
    async (args) => {
      try {
        return asMcpResult(await call("excel_find_value", args));
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const excel_insert_rows = tool(
    "excel_insert_rows",
    "Insert blank rows in a worksheet. Pass `at` as a 1-based row index — `count` rows will be inserted at that position, shifting existing rows down.",
    {
      sheet: z.string().optional().describe("Worksheet name; defaults to the active sheet."),
      at: z.number().int().positive().describe("1-based row index at which to insert."),
      count: z.number().int().positive().default(1),
    },
    async (args) => {
      try {
        return asMcpResult(await call("excel_insert_rows", args));
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const excel_delete_rows = tool(
    "excel_delete_rows",
    "Delete rows from a worksheet, starting at a 1-based row index for `count` rows.",
    {
      sheet: z.string().optional().describe("Worksheet name; defaults to the active sheet."),
      at: z.number().int().positive().describe("1-based row index to start deleting from."),
      count: z.number().int().positive().default(1),
    },
    async (args) => {
      try {
        return asMcpResult(await call("excel_delete_rows", args));
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const wordTools = [
    office_get_selection,
    office_read_paragraphs,
    office_insert_paragraphs,
    office_replace_paragraphs,
    office_replace_text,
    office_replace_section,
    office_highlight,
    office_clear_highlights,
    office_add_comment,
    office_clear_comments,
  ];
  const excelTools = [
    excel_get_selected_range,
    excel_list_sheets,
    excel_read_range,
    excel_write_range,
    excel_find_value,
    excel_insert_rows,
    excel_delete_rows,
  ];
  const tools =
    host === "word" ? wordTools : host === "excel" ? excelTools : [...wordTools, ...excelTools];

  diag(
    `createOfficeBridgeMcp host=${host ?? "both"} → ${tools.length} tools:`,
    tools.map((t) => t?.name).join(", "),
  );

  return createSdkMcpServer({
    name: "office",
    version: "0.1.0",
    // alwaysLoad: ensure these tools are in the initial prompt, not deferred
    // behind tool-search. They are the only sanctioned editing path for the
    // active Office document.
    alwaysLoad: true,
    tools,
  });
}
