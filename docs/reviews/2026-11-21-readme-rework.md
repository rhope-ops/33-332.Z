# README rework, 2026-11-21

Source: the PAM's comparison page, second tab "The site as a README" (https://claude.ai/artifact/BTg6BaUceeGTEH4vMiCvs8#site).
Decision: docs/DECISIONS.md, row of 2026-11-21. Merged at the PAM's request, 2026-11-21 (no wait for network-drives).

## What the branch holds
- README.md rebuilt from that tab: banner, title, intro, Key Facts (Universal, Access, Read-only, Accountable, Boundary,
  Verifiable, Accessible), WCO line as a quote, Copy and Open links, How it works, What it looks for, Limitations,
  If it finds PII (GitHub Caution box), Safety (shorter), Files (without index.html and build_site.py).
- Removed by the PAM: the spot-check note, every internal decision reference, What it scans (now Limitations), Run it,
  Settings, Standards and policy, Develop.
- scripts/build_site.py: `unquote()` reads a quote line or a GitHub alert line as its plain text (self-test case added).
- index.html rebuilt: Key Facts labels Access and Boundary, new Accountable line.

## Open items
- build_site.py MINORs from the last review (unusual inputs that pass the check): `:100:`, `:-1:`, "01." inside a paragraph.
- The Files table may miss newer src modules (for example src/PdfText.psm1); check against src/.
- The site's WCO line shows as small gray text; the PAM asked for a gray bar there too (not built).
- PR 16 to 18 history holds internal references; the PAM's choice between a clean-up and a history rewrite is open.
