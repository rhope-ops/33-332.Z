# Code standard: what ion adds

Decided by the Administrator, 2026-11-20 (docs/DECISIONS.md): DoD/M365 code, so the strictest settings the kit offers, the same as rhope-ops/rade and rhope-ops/kanon. A person must be able to read any script and know what it does before they run it on a gov device.

- Core rules and the check: [docs/code-standard.md](code-standard.md) (a kit file). Extras: [the RADE extras](https://github.com/rhope-ops/rade/blob/main/docs/code-standard-rade.md) (X1 run target in the header, X2 Windows PowerShell 5.1, X4 decisions name who and when, X5 named values, X7 a discarded error says why).
- Settings: the `code-` lines of [.check](../.check): `code-c1-safety: on` (Network:, Writes:, Credentials: each start a header line with a value), `code-c2: full` (full help for each function longer than 10 lines), `code-max-line-chars: 200`, `code-strict-gate: on`, `code-generated: paste/*.ps1` (paste output: `paste/Ion-script.txt`, built from `src/`).
- Gate: the strict gate (`code-strict-gate: on`) in [scripts/code_standard.py](../scripts/code_standard.py) and IMMUTABLE NIPR safety in [scripts/nipr_guard.py](../scripts/nipr_guard.py) (both run by `scripts/check.sh`).
- PowerShell runs on Windows PowerShell 5.1 and PnP.PowerShell 1.12.0 only: `tests/Lint.ps1` checks the syntax; `.github/workflows/ps51.yml` runs the tests on Windows PowerShell 5.1.
