# Word add-in demo script

A 2-minute walkthrough that proves each `office_*` tool actually drives the
live document. Run the prompts in order in the task pane.

## Setup

1. Open **`starter.docx`** (in this folder) in Word. It already has the
   title, the announcement paragraph with deliberate house-style
   violations, and real **Heading 2** sections (_Background_,
   _Next steps_) for the section tools to work on.
2. Open the Claude Code task pane.
3. Workspace: it's this folder (`examples/word-demo/`) — detected from the
   doc's location, no prompt.
4. Setup tab → Context files → add `style-guide.md`.
5. Leave Track Changes **on** (the default).

For reference, `starter.docx` contains:

```
Acme relay Q3 rollout
We are thrilled to announce that acme relay will ship to 5 pilot teams on 6/3/26!
Background            (Heading 2)
The relay project began in March and was scoped by the platform team.
Next steps            (Heading 2)
Teams will get onboarding docs and we will collect feedback after two weeks.
```

## Prompts

| #   | Prompt                                                                                         | Proves                     | Tool exercised                                      |
| --- | ---------------------------------------------------------------------------------------------- | -------------------------- | --------------------------------------------------- |
| 1   | "Read the draft and list each paragraph with its id."                                          | Reading document structure | `office_read_paragraphs`                            |
| 2   | "Fix every product-name and house-style violation per the style guide."                        | Targeted in-place edits    | `office_replace_text` / `office_replace_paragraphs` |
| 3   | "Add a line 'Owner: Platform team' right after the title."                                     | Inserting at a location    | `office_insert_paragraphs`                          |
| 4   | "Rewrite the 'Next steps' section as a numbered list of three concrete actions."               | Section-level rewrite      | `office_replace_section`                            |
| 5   | "Highlight the launch date and add a comment asking the author to confirm it."                 | Annotation                 | `office_highlight` + `office_add_comment`           |
| 6   | Select a sentence in Word, then: "Summarize what I've selected and suggest a tighter version." | Selection awareness        | `office_get_selection`                              |
| 7   | "Clear the highlights and comments you added."                                                 | Cleanup                    | `office_clear_highlights` + `office_clear_comments` |

## What "working" looks like

- Every edit lands as a **tracked change** (colored markup / revision bar)
  you can Accept or Reject in Word — nothing is silently overwritten.
- After #2 the doc reads "**Acme Relay**", "**five** pilot teams",
  "**3 June 2026**", and the exclamation mark is gone.
- #4 replaces only the _Next steps_ section; _Background_ is untouched.
- #5 leaves a visible highlight plus a real Word comment in the margin;
  #7 removes exactly those.
- Doing all of this in Word does **not** disturb a conversation you have
  open in the Excel add-in at the same time — each host is independent.
