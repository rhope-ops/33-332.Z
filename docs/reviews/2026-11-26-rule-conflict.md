# Rule conflict check: Results e-mail (2026-11-26)

Changed line (AGENTS.md "Project rules", added; exact text approved by the Maintainer in the question dialog of 2026-11-26):

- Results e-mail (2026-11-26): after all device results of a PAM e-mail are on origin/main, move that message to the Gmail trash in the same task (the Maintainer's own request for this action, 2026-11-26, under "The Maintainer's approval"; the message holds a gov sender's name and unit).

Step 1 (author's search): no line on e-mail or trash in AGENTS.md. Related: the hard rule "The Maintainer's approval"
(a data write), the project rule "Data boundary", and the DECISIONS row of 2026-11-25 that trashed one results e-mail.

Fresh reviewer: the `reviewer` subagent, same model tier as the author, given only the changed line, AGENTS.md and
docs/DECISIONS.md, with the skill's brief. 2 runs.

Run 1, on the first approved text ("... in the same task, with no question (the Maintainer's standing approval: ...);
trash only the messages whose results are on origin/main."):

- MAJOR: "with no question" can read as a project rule that relaxes the hard rule "The Maintainer's approval". Fix:
  point at the hard rule's own-request clause. The Maintainer approved the new text (above), 2026-11-26.
- MINOR, taken into the new text: the repeated limit (removed); "all" results of a message.
- MINOR, kept: the row did not hold the exact text (fixed in the row); a risk to repeat the sender's name in chat
  ("Data boundary" holds it).

Run 2, on the new text: no MAJOR. MINOR findings (recorded; the rule text is not changed again):

1. A standing approval can read as removing an "ask" case. Fixed in the DECISIONS row: the request of 2026-11-26 is
   the standing approval under the own-request clause.
2. Reading the message brings the sender's name into the session. "Data boundary" bans it in chat, the repo and a
   test; no change.
3. Role words: "the PAM" and "the Maintainer" in one line. In this repo one person holds both roles; no change.
4. The DECISIONS row did not name the 2026-11-25 row it replaces. Fixed in its "Replaces" cell.

Scenario: not run. `.claude/rule-tests.md` is a kit file, so this project rule has no scenario there (lint: untested).

rule-conflict-check: c3b610834ee5
