#!/usr/bin/env python3
"""Self-test of scripts/build_paste.py and of the links to its output.

It proves: the build is the same text twice (no time stamp or random part); the header lists every source with the
first 12 characters of the SHA-256 of its text; each module appears once, in the READERS/MODULES order, and every src/*.psm1 is in MODULES; two
modules that define the same function stop the build; a module's export list and Set-StrictMode line are dropped; the
retired .ps1 copy (DEC 8) is absent; the README "Open the script" link and the site button point at the bundle named
in site.config.json (DEC 7: the one home of the owner, repo and branch).
Usage:   python3 scripts/build_paste_test.py        (prints "BUILD_PASTE TEST PASS" and exits 0, or names the failure)
Inputs:  scripts/build_paste.py, src/, README.md, index.html, site.config.json.
Network: none.
Writes: temporary folders only.
Credentials: none.
"""

import importlib.util
import json
import pathlib
import re
import sys
import tempfile
import unittest

ROOT = pathlib.Path(__file__).resolve().parent.parent
SPEC = importlib.util.spec_from_file_location('build_paste', ROOT / 'scripts' / 'build_paste.py')
bp = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(bp)


class BuildPasteTest(unittest.TestCase):
    """The bundle build and its published links."""

    @classmethod
    def setUpClass(cls):
        """Build the bundle once for all tests."""
        cls.text = bp.build()

    def assertIn(self, member, container, msg=None):  # noqa: N802 (unittest name)
        """Like unittest's assertIn, without printing the whole bundle on a failure."""
        if member not in container:
            self.fail(msg or '%r not found' % (member[:80],))

    def test_build_is_stable(self):
        """Two builds give the same text."""
        self.assertEqual(self.text, bp.build())

    def test_header_lists_each_source_hash(self):
        """The header names each source with its 12-character SHA-256 prefix."""
        for rel in bp.MODULES + (bp.MAIN,):
            self.assertIn('    %s, sha256 %s\n' % (rel, bp.sha256(bp.read(rel))[:12]), self.text, rel)

    def test_modules_once_in_order(self):
        """Each module is in the bundle once, in the MODULES order."""
        marks = [m.group(1) for m in re.finditer(r'^# ==== Module (\S+) ====$', self.text, re.M)]
        self.assertEqual(marks, [pathlib.Path(rel).name for rel in bp.MODULES])

    def test_every_src_module_is_bundled(self):
        """Every src/*.psm1 file is in MODULES (a split module is not left out)."""
        on_disk = sorted('src/' + p.name for p in (ROOT / 'src').glob('*.psm1'))
        self.assertEqual(on_disk, sorted(bp.MODULES))

    def test_reader_block_holds_the_readers(self):
        """The reader script block holds exactly the READERS modules."""
        start = self.text.index('$IonReaderBlock = {')
        end = self.text.index('$null = New-Module -Name IonReaders')
        inside = set(re.findall(r'^# ==== Module (\S+) ====$', self.text[start:end], re.M))
        self.assertEqual(inside, {pathlib.Path(rel).name for rel in bp.READERS})

    def test_duplicate_name_stops_the_build(self):
        """Two modules with the same function name stop the build; different names pass."""
        with self.assertRaises(SystemExit):
            bp.check_names([('a.psm1', 'function Get-X {}\n'), ('b.psm1', 'function Get-X {}\n')])
        bp.check_names([('a.psm1', 'function Get-X {}\n'), ('b.psm1', 'function Get-Y {}\n')])

    def test_module_code_drops_export_and_strict(self):
        """module_code drops the export list (with its continuation line) and the strict-mode line.

        Args:
            self: the test case.
        Returns:
            None."""
        with tempfile.TemporaryDirectory() as tmp:
            saved = bp.ROOT
            try:
                bp.ROOT = pathlib.Path(tmp)
                (bp.ROOT / 'm.psm1').write_text('Set-StrictMode -Version 2.0\nfunction Get-A {}\n'
                                                'Export-ModuleMember -Function Get-A,\n    Get-B\n\nfunction Get-B {}\n')
                code = bp.module_code('m.psm1')
            finally:
                bp.ROOT = saved
        self.assertNotIn('Export-ModuleMember', code)
        self.assertNotIn('    Get-B', code)
        self.assertNotIn(bp.STRICT_LINE, code)
        self.assertIn('function Get-B {}', code)

    def test_retired_copy_is_absent(self):
        """The retired .ps1 copy is not in paste/ and the bundle is the .txt."""
        self.assertFalse(bp.RETIRED.exists())
        self.assertEqual(bp.BUNDLE.name, 'Ion-script.txt')

    def test_links_follow_site_config(self):
        """The README raw link and the site button match site.config.json."""
        config = json.loads((ROOT / 'site.config.json').read_text(encoding='utf-8'))
        self.assertEqual(ROOT / config['script'], bp.BUNDLE)
        raw = 'https://raw.githubusercontent.com/%s/%s/%s/%s' % (config['owner'], config['repo'], config['branch'], config['script'])
        self.assertIn('(%s)' % raw, (ROOT / 'README.md').read_text(encoding='utf-8'))
        self.assertIn('href="%s"' % config['script'], (ROOT / 'index.html').read_text(encoding='utf-8'))


if __name__ == '__main__':
    result = unittest.main(exit=False, verbosity=0).result
    if result.wasSuccessful() and result.testsRun > 0:
        print('BUILD_PASTE TEST PASS: %d tests' % result.testsRun)
        sys.exit(0)
    sys.exit(1)
