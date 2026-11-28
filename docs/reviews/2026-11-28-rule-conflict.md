# Rule conflict check: Role title standardization (2026-11-28)

Changed lines (`AGENTS.md` "Project rules", role titles standardized per `DEC-ROLE` in Fleet Master Plan v3 / Batch F5, approved by the Administrator 2026-11-28 and recorded in `docs/DECISIONS.md`):

- `No reconnaissance (2026-11-20; cannot be changed): a tool touches only the one site or team the person gives; nothing lists other sites, teams or users; any doubt: ask the Administrator first.`
- `Results e-mail (2026-11-26): after all device results of an Administrator e-mail are on origin/main, move that message to the Gmail trash in the same task (the Administrator's own request for this action, 2026-11-26, under "The Administrator's approval"; the message holds a gov sender's name and unit).`

Step 1 (author's search): Kit `16.0.0` standardized kit rules to `the Administrator`. `docs/DECISIONS.md` was updated with the 2026-11-28 decision row superseding the legacy role words (`the PAM` / `the Maintainer`) in `Owner role word`, `Project rules match 90-161.Z`, and `Results e-mail`.

Fresh reviewer: `reviewer` subagent (`41793ced5bec02de5260017a96c62f8a`), same model tier (`inherit`), given only the changed lines (`74dc3b7e3148`), `AGENTS.md`, `docs/DECISIONS.md`, and `.claude/skills/rule-change/SKILL.md` brief.

Result: **no MAJOR findings**.

MINOR findings (recorded; rule text unchanged):
1. Overlap (hard rule vs. project rule): `AGENTS.md` `"The Administrator's approval"` and `No reconnaissance` (`any doubt: ask the Administrator first`) both escalate scope doubts to the Administrator; kept as domain-specific reinforcement.
2. Editorial reference note: two historical task-scoped rows (`2026-11-25`) still quote historical `the PAM` phrasing; archived context only, no live conflict.

Checked and OK: Changed rule lines vs. all `Hard rules`, `Defaults`, `Commands`, remaining `Project rules`, and `docs/DECISIONS.md`.

rule-conflict-check: 74dc3b7e3148
