# Excel demo workspace

This folder is a ready-made workspace for showing the Excel add-in working
end to end. Open **any** workbook in Excel, open the Claude Code task pane,
and point its workspace at **this folder** (it contains this `CLAUDE.md`, so
auto-detect picks it up without prompting).

## What the assistant should know

You are helping build and inspect a small **Q3 regional sales** worksheet.
The source data lives in `q3-sales.csv` in this workspace — add it as a
context file in the task pane (Setup tab → Context files) so you can read it
on demand.

When asked to put data into the sheet:

- Write a header row first, then the data rows beneath it.
- Put the table at `A1` of the active sheet unless told otherwise.
- Use a real Excel formula (e.g. `=SUM(D2:D9)`) for any total — never a
  pre-computed number — so it's obvious the cells are live.
- After writing, read the range back and state the totals so the user can
  see the round-trip succeeded.

Keep replies short; this is a demonstration, not a report.
