![33-332.Z: Privacy Act Management](https://repository-images.githubusercontent.com/1465745904/63428ac3-60d0-44c2-be58-475984d11817)

# Ion
**DoD365 PII Reporter**

<!-- site:intro -->
Keep your eyes on the files in local, network, and cloud directories you can access. Ion checks them for Personally Identifiable Information (PII) and gives you a comprehensive report: an Excel workbook that summarizes and categorizes its findings. It only reads, and every value it finds is masked.
<!-- /site:intro -->

---

### 🔺 Key Facts

<!-- site:key-facts -->
| | |
|---|---|
| **Universal:** | Ion scans every type of file in the chosen directory. |
| **Access:** | uses your CAC sign-on to scan, and saves nothing between sessions. |
| **Read-only:** | it makes no changes to anything it scans. |
| **Accountable:** | see who added or changed PII, and when. |
| **Boundary:** | nothing leaves DoD365; the export is saved to your computer. |
| **Verifiable:** | inspect the commented code yourself within a .txt before execution. |
| **Accessible:** | paste and run via ISE; no download, no admin rights, no policy change. |

> Have concerns? Speak with your WCO.
<!-- /site:key-facts -->

---

[**Copy the script**](https://github.com/rhope-ops/33-332.Z/blob/main/paste/33332Z-pii-spot-check.ps1) · [**Open the script**](https://raw.githubusercontent.com/rhope-ops/33-332.Z/main/paste/33332Z-pii-spot-check.ps1)

The steps to run it are at the top of the script.

---

### 🔺 How it works

<kbd>Sign in</kbd> → <kbd>Read files in memory</kbd> → <kbd>Find PII</kbd> → <kbd><b>Mask it</b></kbd> → <kbd>Report</kbd>

#### What it looks for

<!-- site:looks-for -->
| Type | Found when | Shown as |
|---|---|---|
| SSN | 3-2-4 digits the SSA could issue | `***-**-1234` |
| DoD ID | 10 digits after "DoD ID", "DODID", "EDIPI" or "CAC" | `******7890` |
| Date of birth | A date after "DOB", "date of birth" or "born" | `[date masked]` |
| Payment card | 13 to 19 digits with a card network's prefix that pass the card check | `************1111` |
| Passport | 6 to 9 letters and digits after "passport" | `*****567` |
| Phone | A US 10-digit number | `***-***-0142` |
| Personal e-mail | Any address not ending in .gov or .mil (Gmail, Yahoo and the like rank higher) | `***@domain` |
| Home address | Number, street name and street type (St, Ave, Rd ...) | `[address masked]` |
| Name | After "Name:", "Full name:", "Member:", "Spouse:", "Dependent:", "NOK:" or "POC:" | Initials |
| PII in a file or folder name | Any of the above in the path | As above |
<!-- /site:looks-for -->

---

### 🔺 Limitations

| Not scanned | What to do |
|---|---|
| Channel posts and replies | Reading posts needs a Microsoft Graph permission that a tenant admin must approve; ask your privacy or cybersecurity office for a Microsoft Purview content search |
| Private and shared channels | Each has its own SharePoint site: run Ion again with that site's link |
| Subsites | Run Ion again with the subsite's link |
| Encrypted Office files | A password or an encrypting sensitivity label stops Ion reading the file; it is listed as NOT SCANNED: check it by hand |
| Files over 25 MB | Listed as NOT SCANNED: check them by hand |
| Old .doc .xls .ppt and Outlook .msg | Read in part (numbers in an old .xls are not read): check these files by hand |

---

### 🔺 If it finds PII

<!-- site:breach -->
> [!CAUTION]
> PII where it should not be is a potential or actual breach; report it to your servicing Privacy Manager/Monitor without delay.
>
> AFI 33-332 para 3.1.2, as changed by AFI33-332_DAFGM2026-01 (7 July 2026), page 1, second paragraph.
<!-- /site:breach -->

---

### 🔺 Safety

- **Read-only:** it never changes, moves or deletes anything in any directory it scans.
- **One location only:** it only scans the directory you identify.
- **No stored credentials:** your own CAC sign-in each run.
- **Masked results:** every value Ion finds is masked. A link opens the file, or its folder, without showing PII.
- **No formulas:** every text cell of the workbook is plain text, so text from a scanned file cannot run as a formula in Excel. The site folder name is masked like any other name.
- **Stops cleanly:** if something fails part way or you press Ctrl+C, the workbook is still saved and says INCOMPLETE; the script signs out even if it cannot be saved.
- **Handle the workbook as CUI anyway:** file names and masked fragments still point to where PII is.

---

### 🔺 Files

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
| `ref/afi33-332.pdf`, `ref/afi33-332.txt` | The rule source: AFI 33-332 with AFI33-332_DAFGM2026-01 |
