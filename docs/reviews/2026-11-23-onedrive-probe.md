# OneDrive probe (2026-11-23)

The PAM approved PR 3 of the plan of 2026-11-23 (docs/reviews/2026-11-23-ion-esignature-onedrive-plan.md on the branch
of PR 57): Ion reads a synced OneDrive or SharePoint folder. Ion tells a OneDrive item from a real link with
PowerShell's own `LinkType` and `Target` (a real link has one; a OneDrive item has none), and reads such an item only
inside a folder that the OneDrive app syncs. This probe checks both on the gov device, on Windows PowerShell 5.1,
before PR 3 merges.

## What it does

- Read-only. It lists the one folder that you give and its direct subfolders (no deeper), at most 2,000 items. It opens
  no file and writes nothing.
- It prints counts only, one line for each kind: file or folder, reparse point or not, online only or not, and the
  `LinkType` and `Target` that PowerShell gives, and whether the folder is one that OneDrive syncs. It prints no name,
  no path and no number from a file.
- Expected result: `in OneDrive=True`, and the OneDrive items show `reparse=True LinkType=(none) Target=(none)`. If
  not, Ion skips them as links, and PR 3 needs another test.

## How to run it

1. On the gov device, open Windows PowerShell ISE. You see the window **Windows PowerShell ISE**. Go on.
2. In the window **Windows PowerShell ISE**, do these actions:
   1. Click **File**.
   2. Click **New**.
   You see an empty script pane **Untitled1.ps1**. Go on.
3. On this page, do these actions:
   1. Select all the text in the box below.
   2. Press **Ctrl+C**.
   You see no change: the probe text is on the clipboard. Go on.
4. In the script pane **Untitled1.ps1**, press **Ctrl+V**. You see the probe text in the pane. Go on.
5. In the window **Windows PowerShell ISE**, press **F5**. You see the question **OneDrive or synced folder**. Go on.
6. On the question **OneDrive or synced folder**, do these actions:
   1. Type the path of your OneDrive folder or of a synced library, for example `C:\Users\<you>\OneDrive - <organization>`.
   2. Press **Enter**.
   You see the line **in OneDrive** and lines that start with a number. Wait until the line **Done.** shows.
7. In the ISE output pane, do these actions:
   1. Select the lines from **in OneDrive** to **Done.**.
   2. Press **Ctrl+C**.
   You see the lines selected. Go on.
8. In the chat with the agent, do these actions:
   1. Press **Ctrl+V**.
   2. Send the message.
   You see the lines in the chat. The task is complete.

```powershell
# Ion OneDrive probe: read-only, one folder and its direct subfolders, at most 2,000 items. Prints counts only.
$folder = (Read-Host 'OneDrive or synced folder').Trim().Trim('"').TrimEnd('\')
$top = New-Object System.IO.DirectoryInfo($folder)
if (-not $top.Exists) { Write-Host 'Result: this folder cannot be opened. Check the path, then press F5 again.'; return }
$roots = @(foreach ($name in @('OneDrive', 'OneDriveCommercial', 'OneDriveConsumer')) { [Environment]::GetEnvironmentVariable($name) })
try { foreach ($tenant in @(Get-ChildItem -Path 'HKCU:\Software\Microsoft\OneDrive\Accounts\*\Tenants' -ErrorAction Stop)) { $roots += @($tenant.GetValueNames()) } }
catch { $roots += @() }
$inside = @($roots | Where-Object { $_ -and ($folder + '\').StartsWith($_.TrimEnd('\') + '\', [System.StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
'in OneDrive={0} (sync folders found: {1})' -f $inside, @($roots | Where-Object { $_ }).Count
$items = New-Object System.Collections.Generic.List[object]
foreach ($entry in $top.GetFileSystemInfos()) {
    if ($items.Count -ge 2000) { break }
    $items.Add($entry)
    if ($entry -is [System.IO.DirectoryInfo]) {
        try { foreach ($inner in $entry.GetFileSystemInfos()) { if ($items.Count -ge 2000) { break }; $items.Add($inner) } } catch { $items.Add('blocked') }
    }
}
$lines = foreach ($item in $items) {
    if ($item -is [string]) { 'a folder that cannot be opened'; continue }
    $kind = 'file'; if ($item -is [System.IO.DirectoryInfo]) { $kind = 'folder' }
    $reparse = [bool]($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint)
    $online = (([int]$item.Attributes -band 0x441000) -ne 0)
    $type = '(none)'; $target = '(none)'
    try { $read = [string]([psobject]$item).LinkType; if ($read) { $type = $read } } catch { $type = '(error)' }
    try { if ((@(([psobject]$item).Target) | Where-Object { $_ }).Count -gt 0) { $target = '(set)' } } catch { $target = '(error)' }
    '{0} reparse={1} online={2} LinkType={3} Target={4}' -f $kind, $reparse, $online, $type, $target
}
$lines | Group-Object | Sort-Object Count -Descending | ForEach-Object { '{0,6}  {1}' -f $_.Count, $_.Name }
Write-Host 'Done. The lines above show counts only. Copy them and paste them in the chat.'
```
