# PowerGUIAutomation

GUI automation (FlaUI / UIA3, net10.0-windows) for Power Rider Studio. This file is updated continuously; for now it is general to the whole project, and per-test details will be added later.

## Before every run (general)

1. **Power Rider Studio is open and a unit is connected.** The tests attach to the already-running GUI.
   - **Windows display scaling must be 100%** (Settings > Display > Scale) and the whole GUI window must stay visible on a screen (not minimized, covered or partly off-screen): the red X / green V icons are read from a screenshot (`RowIcons.cs`), and the test warns if the scaling is not 100%. The screen resolution and the number of screens do not matter.
   - **The GUI must run with the UP-TO-DATE Studio DB.** Tests 2 and 3 read Min / Max / Default of each parameter from the GUI tooltip (`Constraints: ...`). If the GUI loaded an old DB, parameters appear without limits (or with wrong ones) and the tests skip them or pick wrong values. Limits that are still only in the *prerelease* DB are visible only when the GUI runs in debug mode with the prerelease channel: check a tooltip (e.g. `Log. I TH Event`) before the run.
   - **Keep the Power Rider Studio window in FRONT and do not use the mouse/keyboard during the run.** The tests type into the GUI (paste, Enter): the keystrokes go to whichever window is in front. If the terminal is in front, the first parameter can fail (see "Known issues" below).
2. **Select the matching unit.** Choose the unit by part number in the main menu (e.g. `RD152`); there is no default. `U` changes the unit. The test verifies that the connected unit (part number and channel count) matches the selection and stops if it does not.
   Scripted run: `PowerGUIAutomation.exe unit=RD152 2`.
3. **Check that the Excel is valid.** The unit's `<PN> Parameters` sheet in the master workbook must have:
   - Columns: `Channel` | `Name Automation ID` | `Parameter Name` | `Value Automation ID` | `Value (default)`.
   - One row per parameter, with both ids and a default value.
   - A unit without its sheet makes the tests stop with a message (e.g. RD250 - no INI file yet).
4. **The master workbook path must be correct.** The path is fixed in code: `ExcelPath` in `Program.cs`
   (`...\QA\Power\scripts\PowerGUIAutomation\Power GUI AU-Tests Master.xlsx` (the master lives in the project folder)). The program checks that the file exists; if it was moved or renamed, update `ExcelPath`.
5. **The master workbook must be closed in Excel while a test runs.** The program checks this before every run and aborts if the file is open.
6. **Building:** the build fails while the `PowerGUIAutomation.exe` console is open (the exe is locked). Close the console, then `dotnet build`.

The unit must be **calibrated** before any test. Choosing a test shows a "Before test N" frame with the requirements of that test (calibrated unit, GUI in front, up-to-date DB, closed workbook, reset to default, what the test writes to the unit, what Ctrl+C leaves behind); Enter starts the test, `Q` goes back to the menu.

## Before tests 1 and 2 (Configure)

**Reset the unit to default before the run.** The tests assume every parameter starts at its Excel default. The program no longer asks: the "Before test N" frame of tests 1 and 2 says it (Enter starts the test).

## Tests (menu)

The Operate - Second Bar test was removed on 2026-10-13 (a copy is in `scripts\_removed`); the `Operate` sheet of the master workbook is kept and is not used by any test now.

| # | Test | Notes |
|---|---|---|
| 1 | Configure - Compare parameters to Excel (`ConfigureCompareParametersToExcelTest`) | compares GUI values with the Excel defaults |
| 2 | Configure - Change and restore every parameter (`ConfigureChangeAndRestoreAllParametersTest`) | **writes to the unit**: one parameter after the other (found by tree page scan): change, Save, read the value back from the unit with the row's own refresh button (global Fetch only as fallback), verify, then put the parameter back to the value it had before the test (its own default after the reset) and check that this value is the Excel default. A ComboBox parameter tries every option in list order and is restored only after the last one. No confirmation question (the instructions frame tells the user to reset the unit first) |
| 3 | Configure - Input validation (`ConfigureInputValidationTest`) | types invalid values (letters, empty, Min-1, Max+1, Hebrew, too-long text) into the text fields of every System and channel parameter (all rows, or the rows / ranges / failures file given at the start, like test 2) and checks that the GUI refuses them (red X) and accepts valid ones. Pressing Enter applies a value to the unit at once; this test **never presses Save** and types the original value back / refreshes the row after every input. Developer switches: environment variables `VALIDATION_ONLY` (name contains) and `VALIDATION_TEXT` (sample of valid text) |
| 4 | Debug tests (`DebugTests`) | `1` (`DebugCompareSubsetTest`): test 1 only on chosen channels (`15`, `13,15`, `S,15`) and/or parameters whose name contains a text. `2` (`DebugAutomationIdScanTest`): AutomationId scan of all parameters; the report goes to a new workbook in the `Excel` folder. `3` (`DebugLoopTest`): **writes to the unit** - runs test 2's change-and-restore on a group of parameters (rows / ranges / failures file, like test 2) again and again, round after round, until you stop it (stop key), the first failure, or one hour has passed (`TimeLimit` in `DebugLoopTest.cs`); the log and the failures file say after how many writes the first failure came (`RESULT| loop`, and "Round r, after W writes to the unit" in the failure type), to tell a failure caused by the quantity from a parameter that really cannot be set |

**Test 2 options (asked at the start):** (1) on a failure either *stop* (default; the closing lines print `RESUME` with the rows to answer next time) or *continue* with the next parameter and list every failure at the end - in the log and in a workbook `Failures\ChangeAndRestore_<PN>_<time>.xlsx` next to the logs (Excel row, channel, parameter, failed step, short reason, whether the default was restored); in continue mode 5 failures in a row stop the run (problem with the unit/GUI). (2) Which Excel rows to run: Enter = all, or `54-70, 72, 100-` (ranges, single rows, `100-` = to the end).

**Loop / overnight run (test 2):** after the rows question the test asks `Loop`: Enter = once, `N` = N rounds, `0` = round after round until you stop it (then an optional time limit in hours). Every round starts with a framed `ROUND | LOOP ROUND r ... STARTED` banner in the console and the log (`=====` lines), the `>>>` line of each parameter says `round r/N`, a `ROUND ... FINISHED` line closes it, every failure type starts with `Round r ...`, and the failures workbook has a `Loop round` column (and a `Loop` line in the `Run` sheet). The PC and the display are kept awake while a test runs (`KeepAwake.cs`); a locked screen is not prevented - disable the lock screen / screen saver / sleep before an overnight run. `Ctrl+Shift+F12` stops after the parameter that is being checked.

**Pass rule of test 2 for one change (checked right after the change AND again after the refresh from the unit):** the requested value is what the field / combo box shows, and the chip next to the field (what the unit reports) is not DIFFERENT from the requested value (a chip equal to it is fine - some GUI builds show it after every set; a red X comes with a different value, `-1` = the unit shows nothing valid). Every failing attempt is its own line in the failures file. A choice parameter (ComboBox) keeps trying its other options after one fails. The red X and the green V are icons that UI Automation does not expose, so they are found by their colours in a screenshot of the area right of the value field (`INDICATOR` lines in the log show what was seen). A red X is a failure and is only reported when it was seen. The green V must appear right after a change (once a V was seen in the run); after the refresh from the unit the GUI does not show it consistently, so it is not required there. Verified on RD279 (GUI 2.1.2-beta.2): `Digital Output 1 Mode/Inverse` > `Blink Slow` and `Blink S. & Inverse` give a red X + chip `-1` and are reported; accepted changes show a green V. The GUI log lines (`Failed to save parameters to unit flash`, `Failed to set ...`) are only copied to the test log (APPLOG), never used as a pass/fail rule. Every `FAIL` line in the .txt log has a row of `!` above and below it.

**Failures file and re-running only the failures (tests 2 and 3):** each run has its own workbook `Failures\<ChangeAndRestore|InputValidation>_<PN>_<run start>.xlsx` next to the exe; its full path is logged at the start (`this run's failures file`). It is written when the run starts (status *Running*), rewritten after every failure and at the end - also after Ctrl+C, closing the console, a stuck run or a crash (the parameter that was being tested is listed as failed, its value on the unit is unknown). A run that finishes with no failures deletes it. Sheets: `Failed` (test 2: one line per failing attempt - Excel row, Channel, Parameter, Failure type in words, Previous value, Value we tried to write, Value read from the unit after refresh, Unit, GUI version, Run time; test 3: one line per failing input), `Run` (status, start/end, run time, unit, rows of the run). At **Excel rows to run** give the path of any failures file - drag the file into the console window - to run ONLY its failed rows (the file must be from a run on the same unit); the log then says `RE-RUN of the failed rows listed in: <file>`, and warns if the parameter name at a row no longer matches the file (master workbook edited). Every test ends with a separate ruler line `════ TEST N (PN) TOOK 1 h 12 min 05 s (start - end) ════` (not in the `TAG |` format, so it is easy to spot in the .txt log); it counts from the real start of the test - the menu, the instructions and the questions are not included. Tests 2 and 3 also print `RESULT| total run time`.

Test 1 (and 4) compare by **page scan**: they select the tree node "System Parameters" / "Channel N Parameters", scroll that scope's parameter list and compare every wanted row as it appears (a row not seen is looked up with the search box). This replaced one search per parameter, which took hours.

Test 1 (and 4) log, for every channel parameter, what the GUI itself shows as the displayed channel (`breadcrumb='… Channel 15 Parameters' | result 15 of 16`). If the wanted channel cannot be reached the test fails; it never compares against another channel's value.

Per-test details: to be added.

## Conventions for every test (current and future)

These rules were set by the test owner; every new test follows them, and tests 2 and 3 already do.

- **Failures workbook** (`Failures\<Test>_<PN>_<run start>.xlsx`): created when the run starts, rewritten after every failure and at the end, also after a stop, a crash or a stuck run; deleted when nothing failed. Sheet `Failed`: Excel row, Channel, Parameter, Failure type, Previous value, Value we tried to write, Value read from the unit, Unit, GUI version, Run time (no "step" column). Sheet `Run`: the test (number and menu name, e.g. `3 - Configure - Change and restore every parameter (writes to the unit)`, because the Failures folder collects the files of every test), Unit, Rows of this run (+ status and time); no counters (the log has them). A later run can use the file: give it at "Excel rows to run".
- **Failure type** = a plain English sentence with a capital letter only at the start, the real values in curly quotes (“X”); a reason that is not classified: `Other - missing information, see the log for details.` A stopped / crashed / stuck test says so and asks to verify the value.
- **Pass rule for a written value:** after the change and again after the refresh the requested value is shown, and no chip that DIFFERS from the requested value appears next to the field (the chip shows what the unit reports; some GUI builds show it after every set, equal to the requested value, and that is fine; a red X comes with a different value such as `-1`) (green V when it can be detected). Anything else is a failure with its own type.
- **Log:** every `FAIL` line has a ruler of `!` above and below it in the .txt log; each test ends with the `TEST N TOOK ...` line (time without the menu and the questions); row ranges are explained in words (`40-end = from row 40 to the end of the sheet`).
- **Before a test:** the "Before test N" frame (cyan header, normal gray body) lists the requirements; no Y/N question for something the frame already says.
- **Stopping:** `Ctrl+Shift+F12` works while the GUI is in front (Ctrl+C reaches only the console when it is in front). Tests that write to the unit stop after the parameter that is being checked; a second press stops at once.
- **One place for every test:** the banner title and the "Before test N" frame come from the catalog in `TestInstructions.cs` (the same rows for every test: unit connected + calibrated + reset to default, GUI in front, closed workbook, what it does, how to stop). A change to those rows is made once there. Boxes are as wide as their longest line.

## Helper files

- The run log is saved next to the exe; console colors are display only.

## Known issues

- **The first parameter fails when the terminal is in front of the GUI.** The tests answer their prompts in the console, so at the start the console is the front window. Keyboard input (paste, Enter) goes to the front window, and setting focus on a GUI element does not always bring the GUI forward, so the first typed value does not reach the GUI (e.g. `Save to Unit Flash` stays disabled). Until this is fixed in the code: after answering the last prompt, click the Power Rider Studio window once so it is in front, and leave it there.
