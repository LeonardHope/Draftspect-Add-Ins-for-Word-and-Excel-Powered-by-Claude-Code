# Excel add-in demo script

A 2-minute walkthrough that proves each `excel_*` tool actually drives the
live workbook. Run the prompts in order in the task pane.

## Setup

1. Open a blank workbook in Excel (Book1 is fine).
2. Open the Claude Code task pane.
3. Workspace: point it at `examples/excel-demo/` (this folder). It has a
   `CLAUDE.md`, so it's detected without a prompt.
4. Setup tab → Context files → add `q3-sales.csv`.

## Prompts

| #   | Prompt                                                                                     | Proves                      | Tool exercised                                 |
| --- | ------------------------------------------------------------------------------------------ | --------------------------- | ---------------------------------------------- |
| 1   | "Put the Q3 sales data from the context file into the sheet, with a header row."           | Writing a block of cells    | `excel_write_range`                            |
| 2   | "Add a 'Total' row at the bottom that sums Units and Revenue with formulas."               | Formulas land as live cells | `excel_write_range`                            |
| 3   | "Read back D2:D9 and tell me the revenue total."                                           | Round-trip read             | `excel_read_range`                             |
| 4   | "Which row has the highest Revenue? Select that rep's name."                               | Lookup + selection          | `excel_find_value`, `excel_get_selected_range` |
| 5   | "Insert a blank row above South so I can add a new rep, then list the sheet's used range." | Structural edit             | `excel_insert_rows`, `excel_list_sheets`       |

## What "working" looks like

- After #1 the grid fills A1:D9 instantly while the pane shows
  _"Writing cells…"_ then returns to **Ready**.
- After #2 clicking the total cell shows a `=SUM(...)` formula in the
  formula bar, not a baked number.
- #3's answer matches the column you can see (sum of Revenue = 1,418,115).
- #4 moves the actual Excel selection to the West/R. Nguyen revenue leader.
- Doing all of this in Excel does **not** disturb a conversation you have
  open in the Word add-in at the same time — each host is independent.
