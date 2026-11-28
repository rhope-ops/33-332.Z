# Administrator guide

This guide is for the person who keeps ion working. It tells you how to change Ion, rebuild the paste text, change the site page and check your work. You do these tasks on a development computer with Python 3, PowerShell 7 and PSScriptAnalyzer 1.23.0. You never do them on a gov computer.

## Facts that you must know

- The files in `src/` are the source of Ion. The file `paste/Ion-script.txt` is generated from them. Do not edit `paste/Ion-script.txt`.
- The page `index.html` is mostly written by hand. Some blocks in it come from `README.md`. These blocks start with `>>> site:`. Do not edit these blocks in `index.html`.
- Keep modules cohesive without arbitrary line-splitting; scripts/code_standard.py checks function length and headers, while ruff and PSScriptAnalyzer check syntax and complexity.
- All test data is invented. Do not put real names, e-mail addresses, DoD IDs, site paths or files in the repository.
- In each step, "the terminal" is a terminal window that is open in the repository folder.

## Change Ion

1. In your editor, change the file in `src/`.
2. In your editor, save the file.
3. In the terminal, do these actions:
   1. Type `python3 scripts/build_paste.py`.
   2. Press **Enter**.

   You see `PASTE BUILT: paste/Ion-script.txt`.
4. In the terminal, do these actions:
   1. Type `pwsh -NoProfile -File tests/Run-Tests.ps1`.
   2. Press **Enter**.

   Wait until you see the line `TESTS passed=`.
5. If the line does not show `failed=0`, do these actions in the terminal:
   1. Fix each item that the output names.
   2. Type `pwsh -NoProfile -File tests/Run-Tests.ps1`.
   3. Press **Enter**.

   Wait until you see the new output.
6. In the terminal, do these actions:
   1. Type `pwsh -NoProfile -File tests/Lint.ps1`.
   2. Press **Enter**.

   You see `findings=0`. The task is complete.

## Add a module to Ion

1. In your editor, save the new `.psm1` file in the folder `src/`.
2. In the file `scripts/build_paste.py`, add the file name to `READERS` or `MODULES`, after each module that it uses.
3. In the files `tests/Run-Tests.ps1` and `src/Find-ChannelPII.ps1`, add the file name to the module list, in the same position.
4. If the new module reads files, add its name to `$script:ReaderModules` in the file `src/ReaderPool.psm1`.
5. In the terminal, do these actions:
   1. Type `python3 scripts/build_paste_test.py`.
   2. Press **Enter**.

   You see `BUILD_PASTE TEST PASS`.
6. In the terminal, do these actions:
   1. Type `python3 scripts/build_paste.py`.
   2. Press **Enter**.

   You see `PASTE BUILT`.
7. In the terminal, do these actions:
   1. Type `pwsh -NoProfile -File tests/Run-Tests.ps1`.
   2. Press **Enter**.

   You see `failed=0`. The task is complete.

## Change the site page

1. In your editor, change the text of `index.html` outside the `>>> site:` blocks.
2. When a `>>> site:` block must change, edit its text between the `<!-- site:` lines in `README.md`.
3. In your editor, save each changed file.
4. In the terminal, do these actions:
   1. Type `python3 scripts/build_site.py --write`.
   2. Press **Enter**.

   You see a line that starts with `SITE`.
5. In the terminal, do these actions:
   1. Type `python3 scripts/build_site.py --check`.
   2. Press **Enter**.

   You see a last line `SITE PASS` or `SITE FAIL`.
6. If the last line is `SITE FAIL`, do these actions in the terminal:
   1. Fix each item that the output names.
   2. Type `python3 scripts/build_site.py --check`.
   3. Press **Enter**.

   Wait until you see the new output.
7. In the terminal, read the last line. You see `SITE PASS`. The task is complete.

## Check all work before a pull request

1. In the terminal, do these actions:
   1. Type `bash scripts/check.sh`.
   2. Press **Enter**.

   Wait until the last line shows (about 2 minutes).
2. If the last line starts with `CHECK FAIL`, do these actions in the terminal:
   1. Fix each item that the output names.
   2. Type `bash scripts/check.sh`.
   3. Press **Enter**.

   Wait until you see the new output.
3. In the terminal, read the last line. You see `CHECK PASS`. The task is complete.
