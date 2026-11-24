#!/usr/bin/env python3
"""Build the paste bundle of Ion (DoD365 PII Reporter): one PowerShell text that a person pastes into Windows
PowerShell ISE and runs with F5 (AGENTS.md "Platform": paste and run, never an execution-policy change).

Why: a gov computer runs pasted text in ISE, not downloaded files (ISSM ruling 2026-11-13). A pasted text loads no
module file and runs no script file, so the execution policy does not apply to it.
What it writes: paste/Ion-script.txt, the one bundle file (a .txt: a .ps1 link causes trouble on gov computers, the PAM,
2026-11-21; the .ps1 copy is retired, fleet v2 DEC 8, 2026-11-23), in this order:
  1. a header (how to run first, then GENERATED, the sources with their SHA-256, Network/Writes/Credentials, a contents list);
  2. the Settings block: plain variables a person may set in ISE (site, results folder, file size limit);
  3. the modules of src/, inline (no module file is loaded), in dependency order: the reader modules inside one
     script block ($IonReaderBlock), loaded as an in-memory module by the script and by each reader worker
     (ReaderPool.psm1 gets the block's text as $IonWorkerInit), then the scan and report modules;
  4. the body of src/Find-ChannelPII.ps1 (from its "# ==== Body ====" line to the end);
  5. a Clean-up block: parts 3 and 4 run inside try/finally, which turns strict mode off and puts the ISE session's
     error setting back when the run ends.
Inputs:  the files in READERS and MODULES below (src/*.psm1) and src/Find-ChannelPII.ps1.
Outputs: the bundle file; with --check, nothing: exit 1 when it differs from a fresh build or the retired .ps1 is back.
Usage:   python3 scripts/build_paste.py            build
         python3 scripts/build_paste.py --check    check (scripts/check.local.sh runs it)
Network: none.
Writes: paste/Ion-script.txt (not with --check).
Credentials: none.
"""

import argparse
import hashlib
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
# The one bundle file: the README's link opens it as plain text, with no .ps1 in the address.
BUNDLE = ROOT / 'paste' / 'Ion-script.txt'
# The retired .ps1 copy (fleet v2 DEC 8, 2026-11-23): --check fails when it is back.
RETIRED = ROOT / 'paste' / '33332Z-pii-spot-check.ps1'
MAIN = 'src/Find-ChannelPII.ps1'
# Dependency order: each module uses only those listed before it (their headers say so).
# The reader modules come first; the reader workers (ReaderPool.psm1) load exactly these.
READERS = ('src/PiiPatterns.psm1', 'src/PiiSignedIds.psm1', 'src/PiiDetectors.psm1', 'src/PiiMasking.psm1', 'src/OcrText.psm1', 'src/OleText.psm1', 'src/ContainerText.psm1',
           'src/PdfCore.psm1', 'src/PdfObjects.psm1', 'src/PdfFonts.psm1', 'src/PdfShown.psm1', 'src/PdfLayout.psm1', 'src/PdfContent.psm1', 'src/PdfText.psm1', 'src/TextExtract.psm1',
           'src/ReaderPool.psm1')
MODULES = READERS + ('src/XlsxWriter.psm1', 'src/Report.psm1', 'src/ReportBook.psm1', 'src/ScanMeter.psm1', 'src/ScanQueue.psm1', 'src/Accountability.psm1',
                     'src/SiteTarget.psm1', 'src/ChannelScan.psm1', 'src/DriveScan.psm1')
READER_START = """# ==== Reader modules ==== (one script block: this script loads it as an in-memory module, and each reader
# worker loads the same text; nothing is read from a file)
$IonReaderBlock = {
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
"""
READER_END = """}
$null = New-Module -Name IonReaders -ScriptBlock $IonReaderBlock
$IonWorkerInit = @{ Code = $IonReaderBlock.ToString() }
"""
BODY_MARK = '# ==== Body ===='
# The end of the bundle: strict mode off and the session's error setting back, also after an error or Ctrl+C, so
# the ISE session the person keeps working in is as it was before F5.
CLEANUP = """# ==== Clean-up ====
} finally {
    Set-StrictMode -Off
    $ErrorActionPreference = $previousErrorAction
}
"""

# Module lines that do not belong in one pasted script: the export list (a module-only command, which may run on
# to the next lines) is dropped; Set-StrictMode is set once in the Settings block.
EXPORT_START = 'Export-ModuleMember'
STRICT_LINE = 'Set-StrictMode -Version 2.0'

# The Settings block: a person may edit these values in ISE before F5; empty means the script asks. The scan
# always covers every folder of the site's library and saves one workbook (the PAM, 2026-11-20).
SETTINGS = """# ==== Settings ==== (optional: edit the values between the quotes; empty = the script asks)
$SiteUrl = ''            # the team's Teams link, or its site https://usaf.dps.mil/teams/<team-site>; empty = asked
$OutputFolder = (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Ion')   # results: a folder per site, then per day
$MaxFileMB = 25          # files larger than this are reported as NOT SCANNED

# ISE runs pasted text in your session: keep its error setting so the Clean-up block at the end can put it back.
$previousErrorAction = $ErrorActionPreference
try {
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
"""


def sha256(text):
    """Return the SHA-256 hex digest of a text (UTF-8)."""
    return hashlib.sha256(text.encode('utf-8')).hexdigest()


def read(rel):
    """Return the text of a repo file with Unix line ends."""
    return (ROOT / rel).read_text(encoding='utf-8').replace('\r\n', '\n')


def module_code(rel):
    """Return a module's code for the bundle: without its export list and its own Set-StrictMode line.

    Args:
        rel: the module path relative to the repo root.
    Returns:
        The code as a string ending in one line break.
    """
    lines = read(rel).split('\n')
    if STRICT_LINE not in lines:
        sys.exit('FAIL %s: no "%s" line (every module sets strict mode)' % (rel, STRICT_LINE))
    start = next((i for i, line in enumerate(lines) if line.startswith(EXPORT_START)), None)
    if start is None:
        sys.exit('FAIL %s: no %s line' % (rel, EXPORT_START))
    end = start + 1
    # The export list may run on to indented or plain continuation lines until the first blank line or the file end.
    while end < len(lines) and lines[end].strip() and not lines[end].startswith(('function ', '#', '$')):
        end += 1
    kept = [line for line in lines[:start] + lines[end:] if line != STRICT_LINE]
    return '\n'.join(kept).rstrip('\n') + '\n'


def body_code():
    """Return the body of the main script: from its "# ==== Body ====" line to the end."""
    text = read(MAIN)
    at = text.find(BODY_MARK)
    if at < 0:
        sys.exit('FAIL %s: no "%s" line' % (MAIN, BODY_MARK))
    return text[at:].rstrip('\n') + '\n'


def check_names(parts):
    """Stop when two modules define the same function or $script: variable (they share one scope in the bundle).

    Args:
        parts: (path, code) pairs of the modules.
    Returns:
        None.
    """
    seen = {}
    for rel, code in parts:
        names = set(re.findall(r'^function ([\w-]+)', code, re.M)) | set(re.findall(r'^\$script:(\w+) =', code, re.M))
        for name in names:
            if name in seen:
                sys.exit('FAIL %s and %s both define %s; rename one (the bundle has one scope)' % (seen[name], rel, name))
            seen[name] = rel


def header(sources, contents):
    """Return the bundle header: how to run, GENERATED, sources, safety lines and the contents list.

    Args:
        sources: (path, sha256) pairs of the source files.
        contents: (line number, title) pairs of the bundle's parts.
    Returns:
        The header text.
    """
    out = ['<#',
           '  Ion - DoD365 PII Reporter (33-332.Z; GENERATED from src/ by scripts/build_paste.py: do not edit this text)',
           '',
           '  HOW TO RUN',
           '    1. Optional: read this text first (Open the script shows it as a .txt in your browser).',
           '    2. Open Windows PowerShell ISE and choose File > New.',
           '    3. Paste ALL of this text (Ctrl+V) and press F5.',
           '    4. When asked, paste a Teams link (... next to the team or channel name > Copy link) or a SharePoint',
           '       site address (in Teams: the Shared tab > ... > Open in SharePoint). Sign in with your CAC.',
           '       If you do not see a sign-in window, press Alt+Tab. Or paste a folder on a drive (X:\\<folder> or',
           '       \\\\<server>\\<share>\\<folder>): it is read with your Windows logon, no sign-in.',
           '    5. Watch the two progress bars (the whole scan, and the file in work) and the line per folder.',
           '       At the end, read the summary of hits per PII type. The results workbook is saved in',
           '       Documents\\Ion\\<site>\\<date>: a Summary tab, then Hits (one row per hit) and Files.',
           '       Handle it as CUI: it shows where PII is.',
           '',
           '  It scans every visible document library of the site (every folder, in a team one per standard channel, and',
           '  the files at each library\'s top), or the drive folder and all its subfolders. Channel posts are not read.',
           '',
           '  Built from these files; to change it, change them and rebuild with: python3 scripts/build_paste.py']
    out += ['    %s, sha256 %s' % (rel, digest[:12]) for rel, digest in sources]
    out += ['',
            '  Read-only: it never changes, moves or deletes anything in Teams or SharePoint. Every value found is masked.',
            '  Run target: Windows PowerShell 5.1 with PnP.PowerShell 1.12.0; no admin rights, no file to unblock.',
            '  Network: only the one site or drive folder given (a folder: through your Windows logon).',
            '  Writes: only the results workbook under $OutputFolder (default Documents\\Ion).',
            '  Credentials: your own CAC sign-in in the PnP sign-in window (a drive folder: your Windows logon); nothing is stored.',
            '',
            '  Contents (line numbers of this file):']
    out += ['    line %5d  %s' % (line, title) for line, title in contents]
    out += ['#>', '']
    return '\n'.join(out) + '\n'


def build():
    """Build the bundle text from the sources.

    Returns:
        The whole bundle text: header, Settings, the modules and the body.
    """
    modules = [(rel, module_code(rel)) for rel in MODULES]
    check_names(modules)
    parts = [('Settings', SETTINGS), ('Reader block', READER_START)]
    for rel, code in modules:
        parts.append(('Module %s' % Path(rel).name, '# ==== Module %s ====\n%s' % (Path(rel).name, code)))
        if rel == READERS[-1]:
            parts.append(('Reader block end', READER_END))
    parts.append(('Body of Find-ChannelPII.ps1', body_code()))
    parts.append(('Clean-up', CLEANUP))
    sources = [(rel, sha256(read(rel))) for rel in MODULES + (MAIN,)]
    # Two passes: the header's own length fixes the line numbers of the parts.
    contents = [(0, title) for title, _ in parts]
    for _ in range(2):
        head = header(sources, contents)
        line, contents = head.count('\n') + 1, []
        for title, text in parts:
            contents.append((line, title))
            line += text.count('\n') + 1
    return head + '\n'.join(text for _, text in parts)


def main():
    """Build the bundle file, or with --check compare it with the committed file. Exit 1 on a difference."""
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('--check', action='store_true', help='fail when paste/ is not the fresh build')
    args = ap.parse_args()
    text = build()
    rel = BUNDLE.relative_to(ROOT)
    if args.check:
        if RETIRED.exists():
            print('PASTE FAIL: paste/%s is retired (DEC 8); delete it' % RETIRED.name)
            sys.exit(1)
        if not BUNDLE.is_file() or BUNDLE.read_text(encoding='utf-8') != text:
            print('PASTE FAIL: %s is not the fresh build; run python3 scripts/build_paste.py and commit it' % rel)
            sys.exit(1)
        print('PASTE PASS: %s is up to date' % rel)
        return
    BUNDLE.parent.mkdir(parents=True, exist_ok=True)
    BUNDLE.write_text(text, encoding='utf-8', newline='\n')
    print('PASTE BUILT: %s (%d lines)' % (rel, text.count('\n')))


if __name__ == '__main__':
    main()
