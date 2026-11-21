#!/usr/bin/env python3
"""33-332.Z's own parts of the site build: the section kinds looks-for and breach, and the tool-text check.

Why: scripts/build_site.py is a kit file (kit 7.0.0) that builds intro and key-facts; a project adds its own kinds in
this file, which a kit sync never touches (the PAM, 2026-11-21: the project's own parts moved out of build_site.py).
Run target: any machine with Python 3 (a build tool, never run on a gov device).
What it does: the kit's build_site.py loads this file and uses
  looks-for  a table Type | Found when | Shown as | Confidence   -> the table's rows (the first three columns)
  breach     the advice, a blank line, its citation               -> the breach note
  extra_checks: the tool's breach text (src/Report.psm1, $script:BreachAdvice) must say the README's breach section,
             word for word (the advice, then its citation in brackets).
Inputs:  README.md, index.html, src/Report.psm1 (through the kit's build_site.py); with --test, invented files only.
Outputs: with --test, one line per case and "SITE LOCAL TEST PASS" (or "... FAIL"); the checks run in build_site.py.
Usage:   python3 scripts/build_site_local.py --test    self-test (scripts/check.sh runs it); the build itself:
         python3 scripts/build_site.py --write | --check
Network: none.
Writes: nothing in the repo (the self-test: temporary files).
Credentials: none.
"""

import argparse
import re
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # the kit's helpers: scripts/build_site.py
from build_site import COMMAND, README, SITE, inline, paragraphs, readme_sections, run  # noqa: E402

TOOL = 'src/Report.psm1'
# The site shows the first three columns of the README's "What it looks for" table (Confidence stays README only);
# their names must be these, as the site's hand-written table head says them (index.html <thead>).
SITE_HEAD = ('Type', 'Found when', 'Shown as')
SITE_COLUMNS = len(SITE_HEAD)
# Where the tool keeps its breach text: one PowerShell string, written as single-quoted parts joined with "+".
TOOL_VAR = re.compile(r"\$script:BreachAdvice\s*=\s*((?:'(?:[^']|'')*'\s*\+?\s*)+)")
TOOL_PART = re.compile(r"'((?:[^']|'')*)'")
# A table's delimiter row (GFM): cells of dashes with optional alignment colons.
DELIMITER_ROW = r'\|(\s*:?-+:?\s*\|)+'


def split_row(line):
    """Split one Markdown table row into cell texts; raise ValueError on an escaped pipe.

    Args:
        line: one table row, for example "| SSN | 3-2-4 digits |". GitHub splits on every "|", even inside `code`.
    Returns:
        The cell texts, stripped.
    """
    if '\\|' in line:
        raise ValueError(f'{README}: an escaped "|" is not supported in a site table: {line}')
    return [cell.strip() for cell in line.strip()[1:-1].split('|')]


def table(markdown):
    """Return a section's table (header cells and body rows) and the text after it.

    Args:
        markdown: the section: a table first (every line starts and ends with "|", the second line is the
            delimiter row), then optional plain text after a blank line.
    Returns:
        A triple: the header cells, the body rows as lists of cell texts, and the text after the table.
    Raises:
        ValueError: when the table is missing, malformed, or a row's cell count differs from the header's.
    """
    lines = markdown.splitlines()
    end = next((n for n, line in enumerate(lines) if not line.strip().startswith('|')), len(lines))
    rows_text, rest = lines[:end], '\n'.join(lines[end:])
    if len(rows_text) < 3 or any(not line.strip().endswith('|') for line in rows_text):
        raise ValueError(f'{README}: a site section must start with a table: every row starts and ends with "|"')
    if not re.fullmatch(DELIMITER_ROW, rows_text[1].strip()) or len(split_row(rows_text[1])) != len(split_row(rows_text[0])):
        raise ValueError(f'{README}: the second table line must be the delimiter row, one |---| cell per header cell')
    if '|' in rest or (rest.strip() and not rest.startswith('\n')):
        raise ValueError(f'{README}: text after a site table needs a blank line, and no other table rows')
    head, rows = split_row(rows_text[0]), [split_row(line) for line in rows_text[2:]]
    for row in rows:
        if len(row) != len(head):
            raise ValueError(f'{README}: a table row has {len(row)} cells, its header {len(head)}: {row}')
    return head, rows, rest


def build_looks_for(markdown, indent):
    """Return the table rows of "What it looks for" (Type, Found when, Shown as).

    Args:
        markdown: the looks-for section: a table with at least three columns.
        indent: the indent of the block's start mark on the page (the rows use it).
    Returns:
        The HTML rows, one per line.
    """
    head, rows, rest = table(markdown)
    if rest.strip():
        raise ValueError(f'{README}: site:looks-for holds only its table')
    if tuple(head[:SITE_COLUMNS]) != SITE_HEAD:
        raise ValueError(f'{README}: site:looks-for must start with the columns {" | ".join(SITE_HEAD)} (the site\'s table head)')
    return '\n'.join(indent + '<tr>' + ''.join(f'<td>{inline(cell)}</td>' for cell in row[:SITE_COLUMNS]) + '</tr>'
                     for row in rows)


def breach_parts(markdown):
    """Return the breach advice and its citation; raise ValueError unless the section has exactly these two paragraphs."""
    parts = paragraphs(markdown)
    if len(parts) != 2:
        raise ValueError(f'{README}: site:breach must hold the advice, a blank line, then its citation')
    return parts[0], parts[1]


def build_breach(markdown, indent):
    """Return the site's breach note: the advice, then its citation in small text.

    Args:
        markdown: the breach section: the advice, a blank line, its citation.
        indent: the indent of the block's start mark on the page.
    Returns:
        The HTML paragraph.
    """
    advice, citation = breach_parts(markdown)
    return f'{indent}<p class="note">{inline(advice)}<br><span class="mut small">{inline(citation)}</span></p>'


# The section kinds this project adds to the kit's intro and key-facts (the interface of scripts/build_site.py).
BUILDERS = {'looks-for': build_looks_for, 'breach': build_breach}


def tool_text(advice, citation):
    """Return the breach text the tool must show: the advice with its citation in brackets."""
    return f'{advice.removesuffix(".")} ({citation.removesuffix(".")}).'


def extra_checks(root, readme):
    """Return findings when the tool's breach text differs from the README's breach section.

    Args:
        root: the repository folder.
        readme: the text of README.md.
    Returns:
        A list of finding lines, empty when they match.
    """
    try:
        tool = (root / TOOL).read_text(encoding='utf-8')
        wanted = tool_text(*breach_parts(readme_sections(readme)['breach']))
    except (OSError, KeyError, ValueError) as err:
        return [f'{TOOL}: breach text cannot be checked: {type(err).__name__}: {err}']
    match = TOOL_VAR.search(tool)
    if not match:
        return [f'{TOOL}: no $script:BreachAdvice string found']
    said = ''.join(part.replace("''", "'") for part in TOOL_PART.findall(match.group(1)))
    return [] if said == wanted else [f'{TOOL}: $script:BreachAdvice differs from {README} site:breach; it must say: {wanted}']


# The self-test's invented README and tool text, and its cases: (label, README edit as (old, new) or None, tool edit
# or None, the words each finding must hold (none: the case must pass), and whether the case rebuilds the site first).
TEST_README = ('<!-- site:intro -->\nA safe_tool & more.\n<!-- /site:intro -->\n'
               '<!-- site:looks-for -->\n| Type | Found when | Shown as | Confidence |\n|---|---|---|---|\n'
               '| SSN | 3-2-4 digits | `***-**-1234` | High |\n<!-- /site:looks-for -->\n'
               '<!-- site:breach -->\nReport it now.\n\nRule para 1.\n<!-- /site:breach -->\n')
TEST_TOOL = "$script:BreachAdvice = 'Report it now ' +\n    '(Rule para 1).'\n"
UNSUPPORTED_WORDS = 'unsupported Markdown'
TEST_CASES = (
    ('a fresh build passes', None, None, [], True),
    ('a stale site fails', None, None, ['out of date'], False),
    ('a tool text that differs fails', None, ('now', 'later'), ['BreachAdvice differs'], True),
    ('a tool with no breach string fails', None, ('BreachAdvice', 'Other'), ['no $script:BreachAdvice'], True),
    ('a renamed site column fails', ('| Shown as |', '| Shown |'), None, ['must start with the columns'], True),
    ('a row with a wrong cell count fails', ('| High |', '| High | extra |'), None, ['cells, its header'], True),
    ('a "|" inside code fails, as GitHub splits the cell there', ('| 3-2-4 digits |', '| 3-2-4 `a|b` digits |'), None,
     ['cells, its header'], True),
    ('a table without its delimiter row fails', ('|---|---|---|---|\n', '| a | b | c | d |\n'), None, ['delimiter row'], True),
    ('a delimiter row with the wrong cell count fails', ('|---|---|---|---|', '|---|---|---|'), None, ['delimiter row'], True),
    ('text after the looks-for table fails', ('| High |\n', '| High |\n\nmore\n'), None, ['holds only its table'], True),
    ('a breach with one paragraph fails', ('\n\nRule para 1.', ''), None, ['advice, a blank line'], True),
    ('unsupported Markdown in the breach fails', ('Report it now.', 'Report it *now*.'), None, [UNSUPPORTED_WORDS], True),
    ('a link in the breach fails', ('Rule para 1.', 'Rule [para 1](x).'), None, [UNSUPPORTED_WORDS], True),
    ('a quote and a GitHub Caution box are read as their plain text',
     ('Report it now.\n\nRule para 1.', '> [!CAUTION]\n> Report it now.\n>\n> Rule para 1.'), None, [], True))
# What the fresh build must write: the row's first three cells with code kept, the Confidence column left out.
TEST_BUILT = ('<td>3-2-4 digits</td><td><code>***-**-1234</code></td></tr>', '<p class="note">Report it now.<br>')


def test_case(edit, tool_edit, expect, write):
    """Run one self-test case in a temporary repository; return whether it behaved as expected, and the findings.

    Args:
        edit: an (old, new) change to the invented README, or None for the README as it is.
        tool_edit: an (old, new) change to the invented tool text, or None.
        expect: the words each finding must hold; empty when the case must pass.
        write: True to rebuild the site before the check.
    Returns:
        A pair: True when the case behaved as expected, and the list of findings.
    """
    readme = TEST_README.replace(*edit) if edit else TEST_README
    tool = TEST_TOOL.replace(*tool_edit) if tool_edit else TEST_TOOL
    marks = ''.join(f'<!-- >>> site:{name}: generated from {README} by {COMMAND} (do not edit) -->\nold\n'
                    f'<!-- <<< site:{name} -->\n' for name in ('intro', 'looks-for', 'breach'))
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        (root / 'src').mkdir()
        (root / 'scripts').mkdir()
        shutil.copy(Path(__file__).resolve(), root / 'scripts' / Path(__file__).name)  # the kit's build loads it from here
        (root / README).write_text(readme, encoding='utf-8')
        (root / SITE).write_text(marks, encoding='utf-8')
        (root / TOOL).write_text(tool, encoding='utf-8')
        found, _ = run(root, write)
        if expect:
            return all(any(word in line for line in found) for word in expect), found
        built = (root / SITE).read_text(encoding='utf-8')
        return not found and all(part in built for part in TEST_BUILT) and '>High<' not in built, found


def self_test():
    """Run every self-test case on invented files and print one line each.

    Returns:
        0 when every case passes, else 1.
    """
    failed = 0
    for label, edit, tool_edit, expect, write in TEST_CASES:
        ok, found = test_case(edit, tool_edit, expect, write)
        failed += not ok
        print(('ok   ' if ok else 'FAIL ') + label + ('' if ok else f': {found}'))
    print('SITE LOCAL TEST PASS' if not failed else 'SITE LOCAL TEST FAIL')
    return 1 if failed else 0


def main():
    """Parse the arguments and run the self-test.

    Returns:
        The exit code: 0 on pass, else 1.
    """
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument('--test', action='store_true', required=True, help='self-test')
    parser.parse_args()
    return self_test()


if __name__ == '__main__':
    sys.exit(main())
