# Ion polish audit (2026-11-21)

Source: a read-only audit by an agent (Sonnet), 2026-11-21, of src/ on branch claude/festive-wozniak-yo5ss3. Recommendations only, until the PAM chose (question dialog, 2026-11-21): build B1-B6, A1-A5, C1, and plan C2 and C4-C6. Built: B1-B6 (PR 36), A1-A5 and C1 (PR 37). C4 and C5 move to the app and C6 is planned first (the PAM, 2026-11-21). C3 needs the PAM's decision (it conflicts with "no choices").

## A. Code cleanup
1. Time as mm:ss is formatted in 3 places (ChannelScan Get-ScanTiming, DriveScan Get-DriveHeader, Find-ChannelPII header); ScanMeter has Format-Duration. Export it and use it. S.
2. DriveScan Invoke-DriveScan assigns `$path`, which is the parameter `$Path` (names are not case-sensitive). Rename to `$workbook`; split the two-statement `finally` line. S.
3. Invoke-OneFileScan and Invoke-DriveFile repeat the file-name hit loop, the skip reasons and the NOT SCANNED row; Invoke-FileScan and Invoke-DriveScan repeat the meter, queue, banner, rule and total setup. Make shared helpers. M.
4. Stale text: the Find-ChannelPII header "Not scanned" list (OCR and old Office readers exist now) and its description (no drive folders); the Get-QueueWorkers comment. S.
5. `$listName` is set twice in Find-ChannelPII; Read-TeamSiteAddress has no help block; Get-ChannelFiles and Get-RootFiles read every item once per channel (items x channels): index once. M.
6. "Channel" names for things that are also folders (fields, Write-ChannelLine). Low value. S.

## B. Screen and UX
1. A cancelled or failed scan shows INCOMPLETE only as a count. Add a red first Results line: where it stopped and how many files were read. S.
2. The last line "nothing stays signed in" is wrong after a drive scan (no sign-in). S.
3. No opening hint: what Ion scans, how to stop early (results so far are saved), the size limit in use. S.
4. The site prompt is one long line with 3 formats; one error message doubles its phrase ("not a Teams link or a site address: Not a team site address"). S.
5. The "most serious hits" table can wrap in a narrow ISE pane: cut the file column to the window width. S/M.
6. Plain words: "Time reading (all workers)", "Files read at once", High/Medium/Low without a key; put the "Saved" path before the breach paragraph. S.

## C. Features (inside the project rules)
1. A confirm step after listing: N files, M folders, X to be skipped, about T minutes; "Scan now? [Y/n]". S.
2. Options per run: the size limit, "only files changed since <date>", hide Low confidence. M.
3. Scan only chosen folders: conflicts with the PAM's "no choices" decision (DECISIONS.md, 2026-11-20); needs the PAM. M.
4. Resume after a cancel: the INCOMPLETE row names the position; a later run can skip folders done in an earlier workbook. L.
5. Compare with the last run: new, resolved and unchanged hits on the Summary tab. L.
6. Review-tracking columns (reviewed by, action, reported to the Privacy Manager, date) and hits per owner. M.
