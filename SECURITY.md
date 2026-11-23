# Security

## Report a vulnerability

Use GitHub private vulnerability reporting for this repository: on the repository's Security tab, choose "Report a vulnerability".

- Do not open a public issue for a security problem.
- Do not put government data, real site addresses, names or other personal data in a report (AGENTS.md "Project rules", data boundary). Use invented data.

## Scope

The PowerShell in `src/` and the paste bundle built from it (`paste/Ion-script.txt`).

## What the tool does not do

- No compiled code: no .exe, .dll or .jar of its own and nothing compiled at run time; it loads only parts of .NET Framework that come with Windows, by name (`tests/Bundle.Tests.ps1` checks the bundle).
- No execution-policy change, no unblocked files, no admin rights, no stored credentials (AGENTS.md "Platform").
- Read-only on Microsoft 365: it changes, moves or deletes nothing; it writes only its own results workbook on the computer.

## Supported version

Only the newest bundle on main.
