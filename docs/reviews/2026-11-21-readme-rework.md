# README rework, 2026-11-21

Source: the PAM's comparison page, second tab "The site as a README" (https://claude.ai/artifact/BTg6BaUceeGTEH4vMiCvs8#site).
Decision: docs/DECISIONS.md, row of 2026-11-21. Branch: claude/nice-wozniak-7ak6yz (not merged).

## What the branch holds
- README.md rebuilt from that tab: banner, title, intro, Key Facts (Universal, Access, Read-only, Accountable, Boundary,
  Verifiable, Accessible), WCO line as a quote, Copy and Open links, How it works, What it looks for, Limitations,
  If it finds PII (GitHub Caution box), Safety (shorter), Files (without index.html and build_site.py).
- Removed by the PAM: the spot-check note, every internal decision reference, What it scans (now Limitations), Run it,
  Settings, Standards and policy, Develop.
- scripts/build_site.py: `unquote()` reads a quote line or a GitHub alert line as its plain text (self-test case added).
- index.html rebuilt: Key Facts labels Access and Boundary, new Accountable line.

## To finish (in order)
1. Wait for the Ion session's network-drives merge: "Accountable: see who added or changed PII, and when." and the
   intro's "local, network, and cloud directories" are true only after it.
2. Merge main into the branch. Keep Limitations true: add a row for anything the tool still does not scan.
3. Fresh review of the diff (lane C): scripts/build_site.py `unquote()` and the README.
4. Show the PAM the README as GitHub renders it, every changed line, before the merge.

## Open items
- build_site.py MINORs from the last review (unusual inputs that pass the check): `:100:`, `:-1:`, "01." inside a paragraph.
- The Files table may miss newer src modules (for example src/PdfText.psm1); check against src/.
- The site's WCO line shows as small gray text; the PAM asked for a gray bar there too (not built).
- PR 16 to 18 history holds internal references; the PAM's choice between a clean-up and a history rewrite is open.
