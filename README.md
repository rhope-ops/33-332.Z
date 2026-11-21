# 33-332.Z · Privacy Act Management

<!-- This top part mirrors the site page (index.html), for networks that block github.io (the PAM, 2026-11-21).
     Change both together. -->

## Ion · DoD365 PII Reporter

Keep your eyes on the files in one Teams team or SharePoint site you can access. Ion checks them for Personally Identifiable Information (PII) and gives you a comprehensive report: an Excel workbook that summarizes and categorizes its findings. It only reads, and every value it finds is masked.

### Key Facts

| | |
|---|---|
| **Universal:** | Ion scans **everything** it can access. |
| **Your access:** | uses your CAC sign-on to scan, and saves nothing between sessions. |
| **Read-only:** | it makes no changes to your Teams or SharePoint content. |
| **Data boundary:** | nothing leaves DoD365; the export is saved to your computer. |
| **Verifiable:** | inspect the commented code yourself within a .txt before execution. |
| **Accessible:** | paste and run via ISE; no download, no admin rights, no policy change. |

Have concerns? Speak with your WCO.

### Get the script

| | |
|---|---|
| **[Copy the script](https://github.com/rhope-ops/33-332.Z/blob/main/paste/33332Z-pii-spot-check.ps1)** | Opens the script on GitHub. Press **Copy raw file** (the copy icon at the top right of the file), then in PowerShell ISE: **File > New**, **Ctrl+V**, **F5**. |
| **[Open the script as .txt](https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/paste/33332Z-pii-spot-check.ps1)** | Opens the script as plain text in your browser, to read it before you run it. **Ctrl+S** saves it as a .txt file. |

The steps to run it are at the top of the script. If one link is blocked on your network, try the other.

### How it works

Sign in → Read files in memory → Find PII → **Mask it** → Report

### What it looks for

| Type | Found when | Shown as |
|---|---|---|
| SSN | 3-2-4 digits the SSA could issue | `***-**-1234` |
| DoD ID | 10 digits after "DoD ID", "EDIPI" or "CAC" | `******7890` |
| Date of birth | A date after "DOB" or "born" | `[date masked]` |
| Payment card | 13 to 19 digits with a card network's prefix that pass the card check | `************1111` |
| Passport | Letters and digits after "passport" | `*****567` |
| Phone | A US 10-digit number | `***-***-0142` |
| Personal e-mail | Any address not ending in .gov or .mil (Gmail, Yahoo and the like rank higher) | `***@domain` |
| Home address | Number, street name and street type | `[address masked]` |
| Name | After "Name:", "Spouse:", "NOK:" and similar | Initials |

Not read: encrypted files, files over 25 MB, and images when Windows has no OCR language (each is listed in the workbook, to check by hand), and channel posts and chats (not checked at all).

### If it finds PII

> PII where it should not be is a potential or actual breach; report it to your servicing Privacy Manager/Monitor without delay.
> <sub>AFI 33-332 para 3.1.2, as changed by AFI33-332_DAFGM2026-01, page 1, second paragraph.</sub>

---

# Details

A read-only PowerShell script that a **member** runs to find files that may hold PII in one Microsoft Teams team or SharePoint site. It uses only your own access: owner rights are not needed, and a private or shared channel is scanned only if you are in it. It scans every visible document library of the one site you give (in a team, the Documents library with one folder per standard channel), every folder and the files at each library's top, with two progress bars (the whole scan, and the file in work) and live counts. At the end it shows the hits per PII type on screen and saves one Excel workbook: a Summary tab first, then one tab per PII type, worst first. **Every found value is masked** (for example `***-**-1234`), so the results are not a PII spill themselves.

> This is a spot check, not a compliance control. The standard tool for PII in Teams is **Microsoft Purview** (DLP policies and Content Search), run by your wing's privacy or cybersecurity office.

## What it scans

| Content | Scanned? | How |
|---|---|---|
| Files in standard channels (Shared tab, formerly Files), or in every visible document library of a SharePoint site (every folder and its top; Site Assets and Preservation Hold Library included; not Style Library, Form Templates or subsites; a library you cannot open is one BLOCKED row) | Yes | Your own CAC sign-in to that one site |
| .docx .xlsx .pptx (and macro versions), .txt .csv .md .log .json .xml .html | Yes | Text read in memory; nothing saved to disk |
| .odt .ods .odp (OpenDocument), .vsdx (Visio), .rtf | Yes | Text read in memory |
| PDF: page text, form fields, comments | Yes | A scanned PDF is read with Windows OCR (first 30 pages); an encrypted PDF is listed as **NOT SCANNED** |
| Images: .jpg .png .tif .bmp .gif | Yes, with Windows OCR | Uses the OCR built into Windows 10/11; without an OCR language each image is **NOT SCANNED** |
| Old .doc .xls .ppt and Outlook .msg | In part | Text runs read; numbers in an old .xls are not, so it gets a note to check by hand |
| .zip and .eml (e-mail) | Yes | Each file inside is read with its own reader and named `file.zip/inner/file`; a zip in a zip in a zip is not opened |
| Files over 25 MB | No | Listed as **NOT SCANNED**: check them by hand |
| Channel posts and replies | No | One sign-in only (the PAM, 2026-11-20). Reading posts needs a Microsoft Graph permission (ChannelMessage.Read.All) that a tenant admin must approve; ask your privacy or cybersecurity office for a Microsoft Purview content search |
| A folder on a mapped or network drive (`X:\<folder>` or `\\<server>\<share>\<folder>`) | Yes | The folder and all its subfolders, read-only with your Windows logon; nothing above it; a folder you cannot open is one BLOCKED row; a link to another folder is not followed |
| Private and shared channels | Run again with that channel's own site URL | Each one has its own SharePoint site |
| 1:1 chats, group chats, meeting chats | No | Admin / Purview only |

## Detection rules and confidence

| Type | Rule | Confidence |
|---|---|---|
| SSN | 3-2-4 digits, valid by SSA issuing rules | High with "SSN"/"SSAN"/"social security" near; Medium with dashes; Low for 9 bare digits |
| DoD ID (EDIPI) | 10 digits with "DoD ID", "DODID", "EDIPI" or "CAC" near | High |
| Date of birth | A date with "DOB", "date of birth" or "born" near | High |
| Payment card | 13–19 digits passing the Luhn check | High with "card"/"GTC"/"Visa" etc. near; else Medium |
| Passport | 6–9 letters/digits with "passport" near | Medium |
| Phone | US 10-digit number | Medium with "cell"/"mobile"/"home"/"personal" near; else Low |
| Personal e-mail | Any address not ending in .gov or .mil | Medium |
| Home address | House number + street name + street type (St, Ave, Rd ...), optional city, state, ZIP | High with "home address"/"residence" near; Medium with state + ZIP; Low for a street only |
| Name | A name after a label: Name:, Full name:, Member:, Spouse:, Dependent:, NOK:, POC: | High when an SSN, DoD ID, date of birth, passport or card number is near; else Low |
| PII in a file or folder name | Any of the above in the path | As above; reported with Source "File name" |

Confidence is a pattern guess. Open each item and check it before you act. A name **without** a label is not detected: that would need a list of real names.

## Run it (gov device)

Needs: Windows PowerShell 5.1 and PnP.PowerShell 1.12.0. No other modules, no admin rights, nothing to download or unblock. You paste one text into PowerShell ISE: pasted text is not a script file, so the execution policy does not apply (ISSM ruling 2026-11-13).

1. Get the paste text: **Copy the script** on the site page, or open [paste/33332Z-pii-spot-check.ps1](https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/paste/33332Z-pii-spot-check.ps1), select all (Ctrl+A) and copy (Ctrl+C). Its first lines repeat these steps.
2. Open **Windows PowerShell ISE**, choose **File > New**, paste into the script pane (Ctrl+V) and press **F5**.
3. When asked, paste the team or site (see "The team or site to scan" below), then sign in with your CAC. If you do not see the sign-in window, press Alt+Tab. A wrong paste is explained and asked again; a site you cannot open says "No access" (close the sign-in window if it shows an access page) and nothing is saved for it.
4. Watch the progress bar (files done of the total, read, not scanned, possible PII per type, time used and time left, and the file in work, PII masked). Under it, a coloured line of live counts appears every 10 seconds and after each folder; each count is 5 characters wide, so the columns stay in place. Colours: cyan labels and headings, white values, green when clean or done, yellow for files to check by hand, red for possible PII. Files of 1 MB or more get a second bar: their size and step (download, read, check). PnP gives no byte count while it downloads, so a large file moves by step.
5. Read the summary on screen: a boxed table of hits per PII type (red with High hits, yellow with Medium), the files with the most serious hits, the files not scanned per reason, and the time spent downloading and checking. Press Enter to open the workbook (or `n`). Then paste another site to scan it, or press Enter to exit. The path of the results workbook is printed at the end: `Documents\Ion\<site>\<yyyy-MM-dd>\Ion Results <HH-mm-ss>.xlsx`.

The team or site to scan, one of:

- **Teams link:** in Teams, open the team or one of its channels, click the **...** next to its name, then **Copy link** (`https://<teams-host>/l/team/...?groupId=...`). The script then asks for the team's site address (the team's Shared tab > **...** > **Open in SharePoint**) and checks that the site belongs to the team in the link (by its group ID) before it reads anything.
- **SharePoint address:** in the team's **Shared** tab, click **...** > **Open in SharePoint** and copy the address (`https://usaf.dps.mil/teams/<team-site>`; some sites use `/sites/`). A file's **Copy link** from the **Shared** tab works too. Anything after the site name is cut off. `https://` may be left out.

The results workbook (one per run; a run never overwrites an earlier one):

| Tab | Holds |
|---|---|
| Summary (first) | The site, the folders, the run time; totals (files read, files with possible PII, hits, High / Medium / Low, not scanned); hits per PII type, worst first; files not scanned per reason; how to read the tabs; the breach step |
| One tab per PII type | High and Medium hits of that type, in this order of tabs: SSN, Date of birth, Passport, Payment card, DoD ID, Home address, Name, Phone, Personal e-mail (the PAM's order, 2026-11-20; AFI 33-332 gives no ranking). Columns: Confidence, Channel (a team) or Folder (a plain site), File, Masked value, Context, Modified, Link (click to open the file), Found in, then who is accountable (the PAM, 2026-11-21): Modified by, Created, Created by, PII added, PII added by. SharePoint and Teams give e-mail addresses, and the version history shows who added each hit and when; a drive gives the file's owner (a Windows logon name), its times and, for Office files, the names saved in the file. Header frozen, filter buttons on |
| Low confidence | Every Low hit (for example 9 bare digits), with its type: kept apart to cut noise |
| Not scanned | Files not read (images, PDFs without text, too large, empty, damaged, encrypted or labelled with encryption) and a scan that stopped early: check these by hand. An Office file with a password or an encrypting sensitivity label is stored in a form the script cannot read |

**Fewer false positives** (the PAM, 2026-11-20): Word field codes (a table of contents' `_Toc123456789` bookmarks) are not read; a number glued to letters (`ABC123456789`) is no SSN, phone or card; a card number must also have a card network's prefix and length; an e-mail at a personal provider (Gmail, Yahoo, Outlook ...) is Medium, at any other non-government domain Low. **Excel files:** each cell is read on its own, so values in neighbouring cells never join into one number. Formula text, decimals (a time such as `45231.4166…`) and exponents are skipped: they made false positives.

Optional settings at the top of the paste text (edit them in ISE before F5):

| Setting | Meaning |
|---|---|
| `$SiteUrl` | The team's Teams link, or its site `https://usaf.dps.mil/teams/<team-site>`; empty = asked |
| `$OutputFolder` | Where to save the results: a folder per site, then per day (default `Documents\Ion`) |
| `$MaxFileMB` | Size limit per file (default 25) |

## Safety

- **Read-only:** it never changes, moves or deletes anything in Teams or SharePoint.
- **One site only:** it touches only the team or site you give. It never lists other sites or teams.
- **No stored credentials:** your own CAC sign-in each run.
- **Masked results:** the screen and the workbook never hold a full value. The context snippet is cut from text that is already masked (digits become `#`, found values and e-mail names become `*`), so a cut cannot expose part of a value. Site, folder and file names show found values masked and every e-mail name as `***`; other digits stay, so the file can be found (`FY27 Plan.docx` stays as it is; the PAM, 2026-11-20). Error messages get a stricter mask: every digit also becomes `#`. A link goes to the file itself when its path holds no PII; else, for an Office file, to the file by its ID (the link names nothing); else to the deepest clean folder.
- **No formulas:** every text cell of the workbook is plain text, so text from a scanned file cannot run as a formula in Excel. The site folder name is masked like any other name.
- **Stops cleanly:** if something fails part way or you press Ctrl+C, the workbook is still saved and says INCOMPLETE; the script signs out even if it cannot be saved.
- **Handle the workbook as CUI anyway:** file names and masked fragments still point to where PII is.
- PII where it should not be is a potential or actual breach; report it to your servicing Privacy Manager/Monitor without delay (AFI 33-332 para 3.1.2, as changed by AFI33-332_DAFGM2026-01, 7 July 2026, page 1, second paragraph).

## Files

Each file starts with a header: purpose, inputs, outputs, how to run, run target, and the three safety lines `Network:`, `Writes:`, `Credentials:`. Read those before you run anything.

| File | What it does |
|---|---|
| `src/Find-ChannelPII.ps1` | Entry point: the five steps (address, sign-in, folders, files, results) |
| `src/ChannelScan.psm1` | SharePoint and Teams steps: check the address, team or plain site, list folders, read files |
| `src/TextExtract.psm1` | Text from Office files, text files and HTML, in memory |
| `src/PiiDetectors.psm1` | One detector per PII type; finds positions and raw values |
| `src/PiiMasking.psm1` | Masks every value and snippet; the only way results leave the detectors |
| `src/Report.psm1` | Screen summary, the severity order, the workbook tabs and where the workbook is saved |
| `src/XlsxWriter.psm1` | Writes the .xlsx workbook with .NET only (no Excel, no module) |
| `src/ScanMeter.psm1` | The two progress bars and the line per folder: counts, sizes and steps only; folder names masked |
| `tests/Run-Tests.ps1` | Runs all tests (invented data, fake PnP cmdlets, no network) |
| `tests/TestKit.ps1`, `tests/FakePnP.ps1` | The test assertion and samples; the fake PnP cmdlets |
| `tests/*.Tests.ps1` | One test file per module, a start-to-end run, and a run of the paste text |
| `tests/Lint.ps1` | Windows PowerShell 5.1 syntax check (PSScriptAnalyzer) |
| `paste/33332Z-pii-spot-check.ps1` | GENERATED: the one text you paste into ISE (the run steps, then a Settings block, the seven modules and the script) |
| `scripts/build_paste.py` | Builds the paste text from `src/`; `--check` fails when it is out of date |
| `ref/afi33-332.pdf`, `ref/afi33-332.txt` | The rule source: AFI 33-332 with AFI33-332_DAFGM2026-01 (edition: docs/DECISIONS.md) |

## Standards and policy it follows

| Source | What it means here |
|---|---|
| [docs/code-standard-33-332.md](docs/code-standard-33-332.md): kit rules C1–C11 with `code-c1-safety: on`, `code-c2: full`, `code-max-line-chars: 200`, and the RADE extras X1–X7 | Headers with safety lines and run target; full help for each function; functions ≤ 60 lines, files ≤ 600, lines ≤ 200; named values with their source; every discarded error says why; `Set-StrictMode -Version 2`. Checked on every pull request by `scripts/check.sh` (the `checks` job) |
| Windows PowerShell 5.1, PnP.PowerShell 1.12.0, no other modules | Lint checks 5.1 syntax; the `ps51` job runs the tests on Windows PowerShell 5.1; only .NET built-ins are used |
| Paste and run (ISSM ruling 2026-11-13) | One generated text pasted into ISE; no download, no unblocking, no execution-policy change |
| No network reconnaissance | Only the one team given (its site); nothing lists other sites or teams |
| Government data stays in DoD365 | File content is read in memory only; reports stay on your computer; the tests use invented data only |
| Credentials | Your own CAC sign-in each run; nothing stored |
| No scripts in results | Script files are not read; only their masked names can appear in a report |
| Decisions in comments | Each names who decided and the date (for example "the PAM, 2026-11-20") |

## Develop

```bash
scripts/check.sh
```

Pass = last line `CHECK PASS`. It runs the kit checks, the content check and `scripts/check.local.sh`: the paste-bundle check, the PowerShell tests (`tests/Run-Tests.ps1`: invented data, fake PnP cmdlets, no network; the last test runs the paste bundle itself) and the lint (`tests/Lint.ps1`, Windows PowerShell 5.1 syntax). After a change in `src/`, rebuild the paste text:

```bash
python3 scripts/build_paste.py
```

On pull requests, the `checks` job runs the same, and the `ps51` job runs the tests on Windows PowerShell 5.1. PSScriptAnalyzerSettings.psd1 says which style rules are off and why. Test data is invented; PII-shaped samples are built from parts at run time, so no complete PII-shaped string is stored in this public repo.
