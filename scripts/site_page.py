#!/usr/bin/env python3
"""Checks of the site page index.html that are not about the README blocks: size, security metas, accessibility parts
and the WCAG contrast of its colour tokens (fleet v2 Stage 3B, 2026-11-23; DEC 5 Option A tokens).

Why: the page is mostly hand-written; these checks keep the fleet v2 fixes in place. Contrast is computed here from the
page's own CSS tokens (WCAG 2.x relative luminance), never by hand.
What it checks (check_page): the page is under MAX_BYTES; a Content-Security-Policy meta with connect-src 'none', a
no-referrer meta, two theme-color metas (light and dark); a skip link to #main; a polite live region; every external
link opens with rel="noopener noreferrer"; each <img> has alt text and its file exists; the reduced-motion, [hidden] and
.sr-only !important rules; each contrast pair in PAIRS, per theme, at its minimum.
Inputs:  index.html and the files it links (through scripts/build_site_local.py extra_checks); with --test, invented
         pages only.
Outputs: a list of problem lines (empty = pass); --test prints "SITE PAGE TEST PASS".
Usage:   python3 scripts/site_page.py [--test]     (scripts/build_site_local.py runs both)
Network: none.
Writes: none.
Credentials: none.
"""

import argparse
import re
import sys
from pathlib import Path

PAGE = 'index.html'
# The page alone (the logos are files): a budget with room for text, far below the 119 KB of inline logos it had.
MAX_BYTES = 35000
# WCAG 2.x relative luminance: the sRGB threshold and the flare term of the contrast ratio.
SRGB_LINEAR_LIMIT = 0.03928
FLARE = 0.05
TEXT_MIN = 4.5   # SC 1.4.3, normal text
UI_MIN = 3.0     # SC 1.4.11, UI borders and focus; SC 1.4.3 large text
# (foreground token, background token, minimum, scope). Scope "page" reads :root (and its dark block); "tool" reads the
# .tool card's own tokens on top (the card is dark in both themes; its background is --card).
PAIRS = (
    ('fg', 'bg', TEXT_MIN, 'page'), ('focus', 'bg', UI_MIN, 'page'), ('ctl', 'bg', UI_MIN, 'page'), ('band-fg', 'band', TEXT_MIN, 'page'),
    ('band-mist', 'band', TEXT_MIN, 'page'), ('gold', 'band', UI_MIN, 'page'),
    ('fg', 'card', TEXT_MIN, 'tool'), ('mut', 'card', TEXT_MIN, 'tool'), ('link', 'card', TEXT_MIN, 'tool'),
    ('err', 'card', TEXT_MIN, 'tool'), ('focus', 'card', UI_MIN, 'tool'), ('ctl', 'card', UI_MIN, 'tool'),
    ('btn-fg', 'btn', TEXT_MIN, 'tool'), ('btn-fg', 'btn-hover', TEXT_MIN, 'tool'), ('btn', 'card', UI_MIN, 'tool'),
    ('btn-hover', 'card', UI_MIN, 'tool'),
)
HEX = r'#[0-9A-Fa-f]{6}'
REQUIRED = (
    (r'<meta http-equiv="Content-Security-Policy" content="[^"]*connect-src \'none\'', "a Content-Security-Policy meta with connect-src 'none'"),
    (r'<meta name="referrer" content="no-referrer">', 'a no-referrer meta'),
    (r'<meta name="theme-color" content="' + HEX + r'" media="\(prefers-color-scheme: light\)">', 'a light theme-color meta'),
    (r'<meta name="theme-color" content="' + HEX + r'" media="\(prefers-color-scheme: dark\)">', 'a dark theme-color meta'),
    (r'<a class="skip-link" href="#main">', 'a skip link to #main'),
    (r'<main id="main"', 'a main element with id="main"'),
    (r'aria-live="polite"', 'a polite live region'),
    (r'@media \(prefers-reduced-motion:reduce\)\{[^}]*transition:none!important', 'the reduced-motion !important rule'),
    (r'\[hidden\]\{display:none!important\}', 'the [hidden] !important rule'),
    (r'\.sr-only\{position:absolute!important', 'the .sr-only !important rule'),
)


def luminance(hex_colour):
    """Return the WCAG relative luminance of a #RRGGBB colour."""
    channels = [int(hex_colour[i:i + 2], 16) / 255 for i in (1, 3, 5)]
    linear = [c / 12.92 if c <= SRGB_LINEAR_LIMIT else ((c + 0.055) / 1.055) ** 2.4 for c in channels]
    return 0.2126 * linear[0] + 0.7152 * linear[1] + 0.0722 * linear[2]


def contrast(fore, back):
    """Return the WCAG contrast ratio of two #RRGGBB colours."""
    light, dark = sorted((luminance(fore), luminance(back)), reverse=True)
    return (light + FLARE) / (dark + FLARE)


def tokens(block):
    """Return the hex custom properties of one CSS block as a dict (name without "--" -> #RRGGBB, upper case)."""
    return {name: value.upper() for name, value in re.findall(r'--([\w-]+):\s*(' + HEX + r')\b', block)}


def themes(html):
    """Return the token sets per theme and scope: {(theme, scope): {name: hex}}.

    Args:
        html: the page text.
    Returns:
        Keys ('light'|'dark', 'page'|'tool'); a missing block gives a set without its tokens (the pair then fails).
    """
    root = re.search(r':root\{(.*?)\}', html, re.S)
    dark = re.search(r':root\[data-theme="dark"\]\{(.*?)\}', html, re.S)
    tool = re.search(r'\n\.tool\{(.*?)\}', html, re.S)
    light_page = tokens(root.group(1)) if root else {}
    dark_page = dict(light_page, **(tokens(dark.group(1)) if dark else {}))
    tool_tokens = tokens(tool.group(1)) if tool else {}
    out = {}
    for theme, page in (('light', light_page), ('dark', dark_page)):
        out[(theme, 'page')] = page
        out[(theme, 'tool')] = dict(page, **tool_tokens)
    return out


def contrast_problems(html):
    """Return one line per contrast pair below its minimum, or with a token missing.

    Args:
        html: the page text.
    Returns:
        A list of problem lines (empty = every pair passes).
    """
    sets, problems = themes(html), []
    for fore, back, minimum, scope in PAIRS:
        for theme in ('light', 'dark'):
            values = sets[(theme, scope)]
            if fore not in values or back not in values:
                problems.append(f'{PAGE}: contrast {fore} on {back} ({scope}, {theme}): a token is missing')
                continue
            ratio = contrast(values[fore], values[back])
            if ratio < minimum:
                problems.append(f'{PAGE}: contrast {fore} {values[fore]} on {back} {values[back]} ({scope}, {theme}) is '
                                f'{ratio:.2f}:1, below {minimum}:1')
    return problems


def check_page(root):
    """Return the problems of the page under root (an empty list = pass).

    Args:
        root: the repo folder (a Path).
    Returns:
        A list of problem lines.
    """
    path = Path(root) / PAGE
    data = path.read_bytes()
    html = data.decode('utf-8')
    problems = [] if len(data) < MAX_BYTES else [f'{PAGE}: {len(data):,} bytes; keep it under {MAX_BYTES:,} (images go in files)']
    problems += [f'{PAGE}: missing {what}' for pattern, what in REQUIRED if not re.search(pattern, html)]
    for tag in re.findall(r'<a [^>]*target="_blank"[^>]*>', html):
        if 'rel="noopener noreferrer"' not in tag:
            problems.append(f'{PAGE}: a new-tab link without rel="noopener noreferrer": {tag[:80]}')
    for tag in re.findall(r'<img [^>]*>', html):
        src, alt = re.search(r'src="([^"]*)"', tag), re.search(r'alt="([^"]+)"', tag)
        if not alt:
            problems.append(f'{PAGE}: an image without alt text: {tag[:80]}')
        if not src or not (Path(root) / src.group(1)).is_file():
            problems.append(f'{PAGE}: an image file is missing: {tag[:80]}')
    return problems + contrast_problems(html)


def self_test():
    """Run the cases on invented pages and print one line each.

    Returns:
        True when every case passes.
    """
    import tempfile
    good = Path(__file__).resolve().parent.parent / PAGE
    cases = (
        ('the page passes', lambda t: t, []),
        ('a low-contrast control line fails', lambda t: t.replace('--focus:#FF8C66;--ctl:#6E7382;', '--focus:#FF8C66;--ctl:#30343D;'), ['ctl #30343D on card']),
        ('no CSP fails', lambda t: t.replace("connect-src 'none'", "connect-src *"), ['Content-Security-Policy']),
        ('no skip link fails', lambda t: t.replace('class="skip-link"', 'class="x"'), ['skip link']),
        ('a missing image fails', lambda t: t.replace('assets/logo-core.svg', 'assets/none.svg'), ['image file is missing']),
        ('an inline logo over the size fails', lambda t: t.replace('</body>', '<!--' + 'x' * MAX_BYTES + '--></body>'), ['bytes']),
        ('no reduced-motion rule fails', lambda t: t.replace('prefers-reduced-motion:reduce', 'print'), ['reduced-motion']),
    )
    ok = abs(contrast('#F2916F', '#F3EEE4') - 2.01) < 0.01 and abs(contrast('#DED6C8', '#F3EEE4') - 1.25) < 0.01
    print(('pass' if ok else 'FAIL') + '  the ratio math gives the plan\'s 2.01 and 1.25')
    with tempfile.TemporaryDirectory() as tmp:
        (Path(tmp) / 'assets').mkdir()
        for name in ('logo-33-332z.svg', 'logo-core.svg'):
            (Path(tmp) / 'assets' / name).write_text('<svg/>')
        for label, edit, expect in cases:
            (Path(tmp) / PAGE).write_text(edit(good.read_text(encoding='utf-8')), encoding='utf-8')
            got = check_page(tmp)
            passed = (not got) if not expect else all(any(e in line for line in got) for e in expect)
            ok = ok and passed
            print(('pass' if passed else 'FAIL') + '  ' + label + ('' if passed else ': ' + '; '.join(got)))
    print('SITE PAGE TEST ' + ('PASS' if ok else 'FAIL'))
    return ok


def main():
    """Check the page, or run the self-test with --test."""
    ap = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    ap.add_argument('--test', action='store_true', help='run the self-test')
    if ap.parse_args().test:
        sys.exit(0 if self_test() else 1)
    problems = check_page(Path(__file__).resolve().parent.parent)
    for line in problems:
        print('SITE PAGE FAIL: ' + line)
    print('SITE PAGE PASS' if not problems else 'SITE PAGE FAIL')
    sys.exit(1 if problems else 0)


if __name__ == '__main__':
    main()
