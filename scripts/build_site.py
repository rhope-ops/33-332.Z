#!/usr/bin/env python3
"""Build the shared parts of the site page (index.html) from README.md, so the README is the one source of truth.

Why: the README and the site said the same facts in two hand-written copies, and they drifted (the PAM, 2026-11-20:
"one needs to be source of truth"; the PAM chose the README). Run target: any machine with Python 3 (a build tool,
never run on a gov device).
What it does: README.md holds four marked sections, "<!-- site:NAME -->" to "<!-- /site:NAME -->". Each one is
turned into HTML and written into index.html between the generated-block marks
"<!-- >>> site:NAME: generated from README.md by scripts/build_site.py --write (do not edit) -->" and
"<!-- <<< site:NAME -->" (the kit's code standard C5 form). It also checks that the tool's breach text
(src/Report.psm1, $script:BreachAdvice) says the README's breach section, word for word.
  intro      one paragraph                               -> the paragraph under the tool's name
  key-facts  an HTML table of rows <tr><td><b>Label:</b></td><td>text</td></tr> (no header row; the PAM, 2026-11-21),
             then one line                                -> the aligned list of key facts, then that line
  looks-for  a table Type | Found when | Shown as | Confidence -> the table's rows (the first three columns)
  breach     the advice, a blank line, its citation      -> the breach note
Inputs:  README.md, index.html, src/Report.psm1.
Outputs: with --write, index.html; then, or with --check alone, "SITE PASS", or one line per finding and "SITE FAIL".
Usage:   python3 scripts/build_site.py --write    rebuild the site's shared parts, then check
         python3 scripts/build_site.py --check    check only (scripts/check.local.sh runs it)
         python3 scripts/build_site.py --test     self-test; prints SITE TEST PASS
Network: none.
Writes: index.html with --write only (the self-test: temporary files).
Credentials: none.
"""

import argparse
import html
import re
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
README = 'README.md'
SITE = 'index.html'
TOOL = 'src/Report.psm1'
COMMAND = 'scripts/build_site.py --write'
# The four shared sections, in README order (the PAM, 2026-11-20, README Source Preview v3).
SECTIONS = ('intro', 'key-facts', 'looks-for', 'breach')
# The site shows the first three columns of the README's "What it looks for" table (Confidence stays README only);
# their names must be these, as the site's hand-written table head says them (index.html <thead>).
SITE_HEAD = ('Type', 'Found when', 'Shown as')
SITE_COLUMNS = len(SITE_HEAD)
# Where the tool keeps its breach text: one PowerShell string, written as single-quoted parts joined with "+".
TOOL_VAR = re.compile(r"\$script:BreachAdvice\s*=\s*((?:'(?:[^']|'')*'\s*\+?\s*)+)")
TOOL_PART = re.compile(r"'((?:[^']|'')*)'")
# Plain text in a site section may hold only these characters (an allow list): letters, digits, spaces and ordinary
# punctuation that GitHub shows as typed. Any other character (* _ at a word's edge, ` ~ [ ] < > @ # = \ | ^ and the
# like) could be Markdown that GitHub renders differently from this builder, so the check fails on it. An underscore
# inside a word is plain text on GitHub (AFI33-332_DAFGM2026-01) and is allowed.
PLAIN_TEXT = r"[A-Za-z0-9 .,;:'\"()\-/?!%+&\u2018\u2019\u201c\u201d\u2013\u2014\u2026]*"
IN_WORD_UNDERSCORE = r'(?<=[A-Za-z0-9])_(?=[A-Za-z0-9])'
# Strong text GitHub renders, in the forms this builder allows: "**" after the line start, a space, "(", a quote or
# "-", before a non-space; the closing "**" after a non-space, before the line end, a space, punctuation, a quote or
# "-". ("**x:**y" stays refused: GitHub cannot close it before a letter.)
BOLD = (r'(?:(?<=^)|(?<=[\s("\'“‘-]))\*\*(?=[^\s*`])([^*`]*?[^\s*`])\*\*'
        r'(?=$|[\s.,;:!?)"\'”’-])')
# Rendered by GitHub although every character is on the allow list: HTML entities (also with digits, &frac12;),
# emoji codes (:warning:; a name that starts with a letter, not right after a letter or digit, so 09:00-17:00 and
# 1:2:3 pass), and web addresses of any scheme or case (ftp://, HTTP://, WWW.).
ENTITY_OR_ADDRESS = (r'&(#\d+|#x[0-9A-Fa-f]+|[A-Za-z][A-Za-z0-9]*);|(?<![A-Za-z0-9]):[A-Za-z_+][A-Za-z0-9_+-]*:'
                     r'|[A-Za-z][A-Za-z0-9+.-]*://|(?i:www\.)')
# A line that starts a list (also an empty item: "+" or "1." alone), is indented (an indented code block, or a
# continuation GitHub may read differently), ends with a hard break, or underlines a heading ("===" or "---").
LIST_START = r'([-+]|\d+[.)])(\s|$)'
# Inside a paragraph, only a bullet or "1." starts a list on GitHub ("2026. That year." stays text).
LIST_INTERRUPT = r'([-+]|1[.)])(\s|$)'
INDENTED = r'[ \t]'
SETEXT_LINE = r'\s*(=+|-+)\s*'
HARD_BREAK = r'.*(  |\\)$'
# A quote mark at a line's start, and a GitHub alert line; both are taken off before a section is built (unquote).
QUOTE_MARK = r'^>[ \t]?'
ALERT_LINE = r'\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]'
# The key-facts table: plain HTML, because a Markdown table always shows a header row and the PAM wants none
# (2026-11-21). One row per line, the label bold and ending in ":"; cells hold plain text only (no tags, no Markdown).
FACT_TABLE = r'<table>\n((?:<tr>.*</tr>\n)+)</table>(.*)'
# No "&" either: in raw HTML a browser reads "&copy" as a symbol even without its ";", and the site would not.
FACT_ROW = r'<tr><td><b>([^<>*`&]+:)</b></td><td>([^<>*`&]+)</td></tr>'
# A table's delimiter row (GFM): cells of dashes with optional alignment colons.
DELIMITER_ROW = r'\|(\s*:?-+:?\s*\|)+'
# HTML indents inside the generated blocks, matching the hand-written page around them.
INDENT_BLOCK = '        '
INDENT_ROW = '              '


def readme_section(text, name):
    """Return the text between a README section's marks, stripped; raise ValueError when it is missing.

    Args:
        text: the whole README.md.
        name: the section name, for example "key-facts".
    Returns:
        The section's Markdown, without its marks.
    """
    match = re.search(rf'<!-- site:{re.escape(name)} -->\n(.*?)\n<!-- /site:{re.escape(name)} -->', text, re.S)
    if not match:
        raise ValueError(f'{README}: no section <!-- site:{name} --> ... <!-- /site:{name} -->')
    return unquote(match.group(1)).strip()


def unquote(markdown):
    """Return Markdown with quote marks ("> ") taken off and GitHub alert lines ("[!CAUTION]") dropped.

    The README shows the WCO line as a quote and the breach note as a GitHub Caution box (the PAM, 2026-11-21); the
    site shows the same words in its own style, so the builder reads them as the plain text inside.

    Args:
        markdown: the text of one README section.
    Returns:
        The same text without quote marks and alert lines.
    """
    lines = [re.sub(QUOTE_MARK, '', line) for line in markdown.splitlines()]
    return '\n'.join(line for line in lines if not re.fullmatch(ALERT_LINE, line.strip()))


def refuse_marks(markdown):
    """Raise ValueError unless plain text (outside `code`, **bold** marks removed) holds only allow-listed characters.

    Args:
        markdown: one line of text, without its `code` spans.
    """
    text = re.sub(IN_WORD_UNDERSCORE, 'x', re.sub(BOLD, r'\1', markdown))
    if not re.fullmatch(PLAIN_TEXT, text) or re.search(ENTITY_OR_ADDRESS, text):
        raise ValueError(f'{README}: unsupported Markdown or character in a site section; only plain text, **bold** and `code`: '
                         f'{markdown}')


def inline(markdown):
    """Turn one line of Markdown (text, **bold**, `code`) into escaped HTML; raise ValueError on any other Markdown.

    Args:
        markdown: one line of text; text inside `code` is kept as is and never made bold.
    Returns:
        The escaped HTML.
    """
    out = []
    for piece in re.split(r'(`[^`]+`)', markdown):
        if piece.startswith('`') and piece.endswith('`') and len(piece) > 1:
            out.append('<code>' + html.escape(piece[1:-1], quote=False) + '</code>')
            continue
        refuse_marks(piece)
        out.append(re.sub(BOLD, r'<strong>\1</strong>', html.escape(piece, quote=False)))
    return ''.join(out)


def paragraphs(markdown):
    """Return the paragraphs of plain Markdown text (blank lines split them; lines inside one are joined).

    Args:
        markdown: text without a table.
    Returns:
        The paragraphs, each on one line; raise ValueError when a line starts a list, heading or quote.
    """
    lines = markdown.splitlines()
    for number, line in enumerate(lines):
        starts = number == 0 or not lines[number - 1].strip()
        if re.match(LIST_START if starts else LIST_INTERRUPT, line) or re.match(INDENTED, line) or re.fullmatch(SETEXT_LINE, line) or re.fullmatch(HARD_BREAK, line):
            raise ValueError(f'{README}: a list, an indented line, a heading underline or a hard line break is not supported '
                             f'in a site section: {line}')
    return [' '.join(part.split()) for part in re.split(r'\n\s*\n', markdown) if part.strip()]


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


def build_intro(markdown):
    """Return the site's intro paragraph."""
    parts = paragraphs(markdown)
    if len(parts) != 1:
        raise ValueError(f'{README}: site:intro must be one paragraph')
    return f'{INDENT_BLOCK}<p>{inline(parts[0])}</p>'


def build_key_facts(markdown):
    """Return the site's aligned key-facts list and the line after it.

    Args:
        markdown: the key-facts section: an HTML table (FACT_TABLE, FACT_ROW), a blank line, one closing line.
    Returns:
        The HTML: a definition list (label, text) and a small paragraph.
    Raises:
        ValueError: when the table or a row has another form, or a cell holds more than plain text.
    """
    match = re.fullmatch(FACT_TABLE, markdown, re.S)
    lines = match.group(1).splitlines() if match else []
    if not match or any(not re.fullmatch(FACT_ROW, line) for line in lines):
        raise ValueError(f'{README}: site:key-facts must be <table>, then rows "<tr><td><b>Label:</b></td><td>text</td></tr>", '
                         f'one per line, then </table>')
    rest = match.group(2)
    if rest.strip() and not rest.startswith('\n\n'):
        raise ValueError(f'{README}: text after the site:key-facts table needs a blank line')
    rows = [re.fullmatch(FACT_ROW, line).groups() for line in lines]
    for cell in (cell for row in rows for cell in row):
        refuse_marks(cell)
    items = ''.join(f'\n{INDENT_BLOCK}  <dt>{html.escape(label.strip(), quote=False)}</dt>'
                    f'<dd>{html.escape(text.strip(), quote=False)}</dd>' for label, text in rows)
    tail = ''.join(f'\n{INDENT_BLOCK}<p class="small mut">{inline(part)}</p>' for part in paragraphs(rest))
    return f'{INDENT_BLOCK}<dl class="facts">{items}\n{INDENT_BLOCK}</dl>{tail}'


def build_looks_for(markdown):
    """Return the table rows of "What it looks for" (Type, Found when, Shown as).

    Args:
        markdown: the looks-for section: a table with at least three columns.
    Returns:
        The HTML rows, one per line.
    """
    head, rows, rest = table(markdown)
    if rest.strip():
        raise ValueError(f'{README}: site:looks-for holds only its table')
    if tuple(head[:SITE_COLUMNS]) != SITE_HEAD:
        raise ValueError(f'{README}: site:looks-for must start with the columns {" | ".join(SITE_HEAD)} (the site\'s table head)')
    return '\n'.join(INDENT_ROW + '<tr>' + ''.join(f'<td>{inline(cell)}</td>' for cell in row[:SITE_COLUMNS]) + '</tr>'
                     for row in rows)


def breach_parts(markdown):
    """Return the breach advice and its citation; raise ValueError unless the section has exactly these two paragraphs."""
    parts = paragraphs(markdown)
    if len(parts) != 2:
        raise ValueError(f'{README}: site:breach must hold the advice, a blank line, then its citation')
    return parts[0], parts[1]


def build_breach(markdown):
    """Return the site's breach note: the advice, then its citation in small text."""
    advice, citation = breach_parts(markdown)
    return f'{INDENT_BLOCK}<p class="note">{inline(advice)}<br><span class="mut small">{inline(citation)}</span></p>'


BUILDERS = {'intro': build_intro, 'key-facts': build_key_facts, 'looks-for': build_looks_for, 'breach': build_breach}


def tool_text(advice, citation):
    """Return the breach text the tool must show: the advice with its citation in brackets."""
    return f'{advice.removesuffix(".")} ({citation.removesuffix(".")}).'


def block_pattern(name):
    """Return the regex of a generated block in index.html: its two marks and what lies between them."""
    start = f'<!-- >>> site:{name}: generated from {README} by {COMMAND} (do not edit) -->'
    return re.compile(rf'({re.escape(start)}\n)(.*?)(\n[ \t]*<!-- <<< site:{re.escape(name)} -->)', re.S)


def render(readme, site):
    """Return the site page with every shared block rebuilt from the README.

    Args:
        readme: the text of README.md.
        site: the text of index.html.
    Returns:
        The new text of index.html.
    Raises:
        ValueError: when a README section or a site block is missing or has the wrong form.
    """
    for name in SECTIONS:
        pattern = block_pattern(name)
        if len(pattern.findall(site)) != 1:
            raise ValueError(f'{SITE}: needs exactly one generated block for site:{name}')
        body = BUILDERS[name](readme_section(readme, name))
        site = pattern.sub(lambda match, body=body: match.group(1) + body + match.group(3), site)
    return site


def check_tool(readme, tool):
    """Return findings when the tool's breach text differs from the README's breach section.

    Args:
        readme: the text of README.md.
        tool: the text of src/Report.psm1.
    Returns:
        A list of finding lines, empty when they match.
    """
    match = TOOL_VAR.search(tool)
    if not match:
        return [f'{TOOL}: no $script:BreachAdvice string found']
    said = ''.join(part.replace("''", "'") for part in TOOL_PART.findall(match.group(1)))
    wanted = tool_text(*breach_parts(readme_section(readme, 'breach')))
    return [] if said == wanted else [f'{TOOL}: $script:BreachAdvice differs from {README} site:breach; it must say: {wanted}']


def run(root, write):
    """Check (and with write, rebuild) the site under root; return the list of findings.

    Args:
        root: the repository folder.
        write: True to rewrite index.html first.
    Returns:
        The findings, empty when everything passes.
    """
    try:
        readme = (root / README).read_text(encoding='utf-8')
        site = (root / SITE).read_text(encoding='utf-8')
        fresh = render(readme, site)
        findings = check_tool(readme, (root / TOOL).read_text(encoding='utf-8'))
    except (OSError, ValueError) as err:
        return [str(err)]
    if write and fresh != site:
        (root / SITE).write_text(fresh, encoding='utf-8')
        site = fresh
    if fresh != site:
        findings.append(f'{SITE}: out of date with {README}; run: python3 {COMMAND}')
    return findings


# The self-test's invented README, its tool text, and its cases: (label, README edit as (old, new) or None,
# the words each finding must hold (none: the case must pass), and whether the case rebuilds the site first).
TEST_README = ('<!-- site:intro -->\nA safe_tool & more.\n<!-- /site:intro -->\n'
               '<!-- site:key-facts -->\n<table>\n<tr><td><b>Read-only:</b></td><td>it changes nothing.</td></tr>\n</table>\n\nFirst line.\nAsk & learn.\n'
               '<!-- /site:key-facts -->\n'
               '<!-- site:looks-for -->\n| Type | Found when | Shown as | Confidence |\n|---|---|---|---|\n'
               '| SSN | 3-2-4 digits | `***-**-1234` | High |\n<!-- /site:looks-for -->\n'
               '<!-- site:breach -->\nReport it now.\n\nRule para 1.\n<!-- /site:breach -->\n')
# The words that the findings for unsupported text and for unsupported lines start with.
UNSUPPORTED_WORDS = 'unsupported Markdown'
BLOCK_WORDS = 'is not supported in a site section'
TEST_TOOL = "$script:BreachAdvice = 'Report it now ' +\n    '(Rule para 1).'\n"
TEST_CASES = (('a fresh build passes', None, None, [], True),
              ('a stale site fails', None, None, ['out of date'], False),
              ('a tool text that differs fails', None, ('now', 'later'), ['BreachAdvice differs'], True),
              ('a missing README section fails', ('<!-- site:intro -->', '<!-- site:gone -->'), None, ['no section'], True),
              ('a key fact without a bold label fails', ('<b>Read-only:</b>', 'Read-only:'), None, ['<b>Label:</b>'], True),
              ('a key fact with Markdown fails', ('changes nothing.', 'changes **nothing**.'), None, ['<b>Label:</b>'], True),
              ('a key fact with "&" fails', ('changes nothing.', 'changes &copy nothing.'), None, ['<b>Label:</b>'], True),
              ('text right after the key-facts table fails', ('</table>\n\nFirst line.', '</table>\nFirst line.'), None,
               ['needs a blank line'], True),
              ('a key fact with a link fails', ('changes nothing.', 'changes www.x.example'), None, [UNSUPPORTED_WORDS], True),
              ('a key-facts table as Markdown fails', ('<table>\n<tr><td><b>Read-only:</b></td><td>it changes nothing.</td></tr>\n</table>',
               '| | |\n|---|---|\n| **Read-only:** | it changes nothing. |'), None, ['<b>Label:</b>'], True),
              ('a renamed site column fails', ('| Shown as |', '| Shown |'), None, ['must start with the columns'], True),
              ('a row with a wrong cell count fails', ('| High |', '| High | extra |'), None, ['cells, its header'], True),
              ('unsupported Markdown fails', ('Report it now.', 'Report it *now*.'), None, [UNSUPPORTED_WORDS], True),
              ('a link fails', ('Rule para 1.', 'Rule [para 1](x).'), None, [UNSUPPORTED_WORDS], True),
              ('a "|" inside code fails, as GitHub splits the cell there', ('| 3-2-4 digits |', '| 3-2-4 `a|b` digits |'), None,
               ['cells, its header'], True),
              ('a table without its delimiter row fails', ('|---|---|---|---|\n', '| a | b | c | d |\n'), None, ['delimiter row'], True),
              ('a list in a section fails', ('First line.', '- First line.'), None, [BLOCK_WORDS], True),
              ('an empty list item fails', ('First line.', '1.\nFirst line.'), None, [BLOCK_WORDS], True),
              ('an indented code block fails', ('Ask & learn.', 'Ask & learn.\n\n    code here'), None, [BLOCK_WORDS], True),
              ('a heading underline fails', ('Ask & learn.', 'Ask & learn.\n==='), None, [BLOCK_WORDS], True),
              ('bold with inner spaces fails', ('Report it now.', 'Report it ** now **.'), None, [UNSUPPORTED_WORDS], True),
              ('bold GitHub would not close fails', ('Report it now.', 'Report **it:**now.'), None, [UNSUPPORTED_WORDS], True),
              ('a delimiter row with the wrong cell count fails', ('|---|---|---|---|', '|---|---|---|'), None, ['delimiter row'], True),
              ('an e-mail address fails', ('Rule para 1.', 'Rule para 1, ask a@b.example.'), None, [UNSUPPORTED_WORDS], True),
              ('an entity with digits fails', ('Rule para 1.', 'Rule para &frac12;.'), None, [UNSUPPORTED_WORDS], True),
              ('an emoji code fails', ('Report it now.', 'Report it now :warning:'), None, [UNSUPPORTED_WORDS], True),
              ('plain times, ratios and bold in brackets or quotes pass',
               ('A safe_tool & more.', 'A safe_tool (**Microsoft Purview**), "**Stop**", **PII**-free, 09:00-17:00, 10:30:45, 1:2:3,\n'
                '2026. That year.'), None, [], True),
              ('a quote and a GitHub Caution box are read as their plain text',
               ('Report it now.\n\nRule para 1.', '> [!CAUTION]\n> Report it now.\n>\n> Rule para 1.'), None, [], True),
              ('a web address of another scheme fails', ('Rule para 1.', 'Rule para 1, see ftp://x.example'), None,
               [UNSUPPORTED_WORDS], True))
# What the fresh build must write: a key fact, code kept, the Confidence column left out, two lines joined, an escaped "&".
TEST_BUILT = ('<dt>Read-only:</dt><dd>it changes nothing.</dd>', '<td>3-2-4 digits</td><td><code>***-**-1234</code></td></tr>',
              'First line. Ask &amp; learn.</p>')


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
                    f'<!-- <<< site:{name} -->\n' for name in SECTIONS)
    with tempfile.TemporaryDirectory() as tmp:
        root = Path(tmp)
        (root / 'src').mkdir()
        (root / README).write_text(readme, encoding='utf-8')
        (root / SITE).write_text(marks, encoding='utf-8')
        (root / TOOL).write_text(tool, encoding='utf-8')
        found = run(root, write)
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
    print('SITE TEST PASS' if not failed else 'SITE TEST FAIL')
    return 1 if failed else 0


def main():
    """Parse the arguments and run the rebuild, the check or the self-test.

    Returns:
        The exit code: 0 on pass, else 1.
    """
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument('--write', action='store_true', help='rebuild the shared parts of index.html, then check')
    mode.add_argument('--check', action='store_true', help='check only')
    mode.add_argument('--test', action='store_true', help='self-test')
    args = parser.parse_args()
    if args.test:
        return self_test()
    found = run(ROOT, args.write)
    for line in found:
        print(line)
    print('SITE PASS' if not found else 'SITE FAIL')
    return 1 if found else 0


if __name__ == '__main__':
    sys.exit(main())
