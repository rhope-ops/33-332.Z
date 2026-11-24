# Rule conflict check: Paste format (2026-11-24)

Changed line (AGENTS.md "Project rules", added; exact text approved by the PAM in the question dialog of 2026-11-23):

- Paste format (2026-11-23; cannot be changed): every script a person pastes on a gov device (Ion and each test snippet) is a .txt file opened by a raw link; never a .ps1 file and never only a code block in a page.

Step 1 (author's search): no conflicting line. Related: the Platform rule ("one generated text pasted into PowerShell ISE")
and the DECISIONS rows of 2026-11-21 (Open the script opens a .txt) and 2026-11-23 (one bundle file, paste/Ion-script.txt).

Fresh reviewer: the `reviewer` subagent, same model tier as the author, given only the changed line, AGENTS.md and
docs/DECISIONS.md, with the skill's brief.

Result: no MAJOR finding. MINOR findings (recorded; the rule text is not changed again):

1. Date: the decision row was dated 2026-11-24, the rule 2026-11-23. Fixed in the row (the approval was on 2026-11-23).
2. Overlap with the Platform rule (both name how a script reaches the gov device). Accepted: Platform says one text
   pasted into ISE; Paste format says that text is a .txt by a raw link.
3. "opened by a raw link" can be read as a download (the gov-device row says no download). Meaning: opened in the
   browser and copied, not saved. To say in each step page.
4. Older test pages of 2026-11-21 have only a code block. They predate the rule; a new run of them moves the snippet
   to a .txt first.
5. "never a .ps1 file" binds only what a person pastes; the repo's own .ps1 source and test files stay.
6. Outside this change (for a later audit): the role-word row says "the PAM" while AGENTS.md says "the Maintainer";
   a `.check` code-generated line may still name paste/*.ps1; the "no parallel downloads" row and the later
   "2-4 parallel downloads" row do not link.

Scenario: not run. `.claude/rule-tests.md` is a kit file (a kit update replaces it), so this project rule has no
scenario there.

rule-conflict-check: a9d4b4efc454
