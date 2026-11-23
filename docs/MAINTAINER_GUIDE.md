# Maintainer guide

This guide is for the person who keeps 33-332.Z working. It tells you how to change Ion, rebuild the paste text, change the site page and check your work. You do these tasks on a development computer with Python 3, PowerShell 7 and PSScriptAnalyzer 1.23.0. You never do them on a gov computer.

## Facts that you must know

- The files in `src/` are the source of Ion. The file `paste/Ion-script.txt` is generated from them. Do not edit `paste/Ion-script.txt`.
- The page `index.html` is mostly written by hand. Some blocks in it come from `README.md`. These blocks start with `>>> site:`. Do not edit these blocks in `index.html`.
- Each hand-written code file has a limit of 450 lines (fleet v2 DEC 6, 2026-11-23). Count the lines with `wc -l`.
- All test data is invented. Do not put real names, e-mail addresses, DoD IDs, site paths or files in the repository.
- In each step, "the terminal" is a terminal window that is open in the repository folder.

## Change Ion

1. In your editor, change the file in `src/`. You see the change in the editor. Go on.
2. In your editor, save the file. You see the saved file. Go on.
3. In the terminal, do these actions:
   1. Type `python3 scripts/build_paste.py`.
   2. Press **Enter**.

   You see `PASTE BUILT: paste/Ion-script.txt`. Go on.
4. In the terminal, do these actions:
   1. Type `pwsh -NoProfile -File tests/Run-Tests.ps1`.
   2. Press **Enter**.

   You see the line `TESTS passed=`. Wait until the line shows.
5. If the line does not show `failed=0`, do these actions in the terminal:
   1. Correct each item that the output names.
   2. Type `pwsh -NoProfile -File tests/Run-Tests.ps1` again.
   3. Press **Enter**.

   You see the new output. Go on.
6. In the terminal, do these actions:
   1. Type `pwsh -NoProfile -File tests/Lint.ps1`.
   2. Press **Enter**.

   You see `findings=0`. The task is complete.

## Add a module to Ion

1. In your editor, save the new `.psm1` file in the folder `src/`. You see the file in `src/`. Go on.
2. In the file `scripts/build_paste.py`, add the file name to `READERS` or `MODULES`, after each module that it uses. You see the name in the list. Go on.
3. In the files `tests/Run-Tests.ps1` and `src/Find-ChannelPII.ps1`, add the file name to the module list, in the same position. You see the name in each list. Go on.
4. If the new module reads files, add its name to `$script:ReaderModules` in the file `src/ReaderPool.psm1`. You see the name in the list. Go on.
5. In the terminal, do these actions:
   1. Type `python3 scripts/build_paste_test.py`.
   2. Press **Enter**.

   You see `BUILD_PASTE TEST PASS`. Go on.
6. In the terminal, do these actions:
   1. Type `python3 scripts/build_paste.py`.
   2. Press **Enter**.

   You see `PASTE BUILT`. Go on.
7. In the terminal, do these actions:
   1. Type `pwsh -NoProfile -File tests/Run-Tests.ps1`.
   2. Press **Enter**.

   You see `failed=0`. The task is complete.

## Change the site page

1. In your editor, change the text of `index.html` outside the `>>> site:` blocks. You see the change in the editor. Go on.
2. If you must change a `>>> site:` block, change its text between the `<!-- site:` lines in `README.md`. You see the change in the editor. Go on.
3. In your editor, save each changed file. You see the saved files. Go on.
4. In the terminal, do these actions:
   1. Type `python3 scripts/build_site.py --write`.
   2. Press **Enter**.

   You see a line that starts with `SITE`. Go on.
5. In the terminal, do these actions:
   1. Type `python3 scripts/build_site.py --check`.
   2. Press **Enter**.

   You see a last line `SITE PASS` or `SITE FAIL`. Go on.
6. If the last line is `SITE FAIL`, do these actions in the terminal:
   1. Correct each item that the output names.
   2. Type `python3 scripts/build_site.py --check` again.
   3. Press **Enter**.

   You see the new output. Go on.
7. In the terminal, read the last line. You see `SITE PASS`. The task is complete.

## Check all work before a pull request

1. In the terminal, do these actions:
   1. Type `bash scripts/check.sh`.
   2. Press **Enter**.

   You see the lines of each check. Wait until the last line shows (about 30 seconds; 25 seconds on 2026-11-23).
2. If the last line starts with `CHECK FAIL`, do these actions in the terminal:
   1. Correct each item that the output names.
   2. Type `bash scripts/check.sh` again.
   3. Press **Enter**.

   You see the new output. Go on.
3. In the terminal, read the last line. You see `CHECK PASS`. The task is complete.
