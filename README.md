# PowerGUIAutomation

GUI automation (FlaUI / UIA3, net10.0-windows) for Power Rider Studio. This file is updated continuously; for now it is general to the whole project, and per-test details will be added later.

## Before every run (general)

1. **Power Rider Studio is open and a unit is connected.** The tests attach to the already-running GUI.
   - **The GUI must run with the UP-TO-DATE Studio DB.** Tests 3 and 4 read Min / Max / Default of each parameter from the GUI tooltip (`Constraints: ...`). If the GUI loaded an old DB, parameters appear without limits (or with wrong ones) and the tests skip them or pick wrong values. Limits that are still only in the *prerelease* DB are visible only when the GUI runs in debug mode with the prerelease channel: check a tooltip (e.g. `Log. I TH Event`) before the run.
   - **Keep the Power Rider Studio window in FRONT and do not use the mouse/keyboard during the run.** The tests type into the GUI (paste, Enter): the keystrokes go to whichever window is in front. If the terminal is in front, the first parameter can fail (see "Known issues" below).
2. **Select the matching unit.** Choose the unit by part number in the main menu (e.g. `RD152`); there is no default. `U` changes the unit. The test verifies that the connected unit (part number and channel count) matches the selection and stops if it does not.
   Scripted run: `PowerGUIAutomation.exe unit=RD152 3`.
3. **Check that the Excel is valid.** The unit's `<PN> Parameters` sheet in the master workbook must have:
   - Columns: `Channel` | `Name Automation ID` | `Parameter Name` | `Value Automation ID` | `Value (default)`.
   - One row per parameter, with both ids and a default value.
   - A unit without its sheet makes the tests stop with a message (e.g. RD250 - no INI file yet).
4. **The master workbook path must be correct.** The path is fixed in code: `ExcelPath` in `Program.cs`
   (`...\QA\Power\scripts\PowerGUIAutomation\Power GUI AU-Tests Master.xlsx` (the master lives in the project folder)). The program checks that the file exists; if it was moved or renamed, update `ExcelPath`.
5. **The master workbook must be closed in Excel while a test runs.** The program checks this before every run and aborts if the file is open.
6. **Building:** the build fails while the `PowerGUIAutomation.exe` console is open (the exe is locked). Close the console, then `dotnet build`.

The unit must be **calibrated** before any test. Choosing a test shows a "Before test N" frame with the requirements of that test (calibrated unit, GUI in front, up-to-date DB, closed workbook, reset to default, what the test writes to the unit, what Ctrl+C leaves behind); Enter starts the test, `Q` goes back to the menu.

## Before tests 2 and 3 (Configure)

**Reset the unit to default before the run.** The tests assume every parameter starts at its Excel default. The program no longer asks: the "Before test N" frame of tests 2 and 3 says it (Enter starts the test).

## Tests (menu)

| # | Test | Notes |
|---|---|---|
| 1 | Operate - Second Bar (`OperateSecondBarTest`) | reads ids from the `Operate` sheet |
| 2 | Configure - Compare parameters to Excel (`ConfigureCompareParametersToExcelTest`) | compares GUI values with the Excel defaults |
| 3 | Configure - Change and restore every parameter (`ConfigureChangeAndRestoreAllParametersTest`) | **writes to the unit**: one parameter after the other (found by tree page scan): change, Save, read the value back from the unit with the row's own refresh button (global Fetch only as fallback), verify, then restore the Excel default the same way. A ComboBox parameter tries every option in list order and is restored only after the last one. No confirmation question (the instructions frame tells the user to reset the unit first) |
| 4 | Configure - Input validation (`ConfigureInputValidationTest`) | types invalid values (letters, empty, Min-1, Max+1, Hebrew, too-long text) into the text fields of the System and channel 1 parameters and checks that the GUI refuses them (red X) and accepts valid ones. Pressing Enter applies a value to the unit at once; this test **never presses Save** and types the original value back / refreshes the row after every input. Developer switches: environment variables `VALIDATION_ONLY` (name contains) and `VALIDATION_TEXT` (sample of valid text) |
| 5 | Debug tests (`DebugTests`; read-only) | `1` (`DebugCompareSubsetTest`): test 2 only on chosen channels (`15`, `13,15`, `S,15`) and/or parameters whose name contains a text. `2` (`DebugAutomationIdScanTest`): AutomationId scan of all parameters; the report goes to a new workbook in the `Excel` folder |

**Test 3 options (asked at the start):** (1) on a failure either *stop* (default; the closing lines print `RESUME` with the rows to answer next time) or *continue* with the next parameter and list every failure at the end - in the log and in a workbook `Failures\ChangeAndRestore_<PN>_<time>.xlsx` next to the logs (Excel row, channel, parameter, failed step, short reason, whether the default was restored); in continue mode 5 failures in a row stop the run (problem with the unit/GUI). (2) Which Excel rows to run: Enter = all, or `54-70, 72, 100-` (ranges, single rows, `100-` = to the end).

**Failures file and re-running only the failures (tests 3 and 4):** each run has its own workbook `Failures\<ChangeAndRestore|InputValidation>_<PN>_<run start>.xlsx` next to the exe; its full path is logged at the start (`this run's failures file`). It is written when the run starts (status *Running*), rewritten after every failure and at the end - also after Ctrl+C, closing the console, a stuck run or a crash (the parameter that was being tested is listed as failed, its value on the unit is unknown). A run that finishes with no failures deletes it. Sheets: `Failed` (one line per failing parameter in test 3, one line per failing input in test 4; column A = Excel row, column D = why it failed), `Run` (status, start/end, run time, mode, counters) and `Not run` (rows the run never reached). At **Excel rows to run** give the path of any failures file - drag the file into the console window - to run ONLY its failed rows (the file must be from a run on the same unit); the log then says `RE-RUN of the failed rows listed in: <file>`, and warns if the parameter name at a row no longer matches the file (master workbook edited). Every test ends with a separate ruler line `════ TEST N (PN) TOOK 1 h 12 min 05 s (start - end) ════` (not in the `TAG |` format, so it is easy to spot in the .txt log); it counts from the real start of the test - the menu, the instructions and the questions are not included. Tests 3 and 4 also print `RESULT| total run time`.

Test 2 (and 5) compare by **page scan**: they select the tree node "System Parameters" / "Channel N Parameters", scroll that scope's parameter list and compare every wanted row as it appears (a row not seen is looked up with the search box). This replaced one search per parameter, which took hours.

Test 2 (and 5) log, for every channel parameter, what the GUI itself shows as the displayed channel (`breadcrumb='… Channel 15 Parameters' | result 15 of 16`). If the wanted channel cannot be reached the test fails; it never compares against another channel's value.

Per-test details: to be added.

## Helper files

- The run log is saved next to the exe; console colors are display only.

## Known issues

- **The first parameter fails when the terminal is in front of the GUI.** The tests answer their prompts in the console, so at the start the console is the front window. Keyboard input (paste, Enter) goes to the front window, and setting focus on a GUI element does not always bring the GUI forward, so the first typed value does not reach the GUI (e.g. `Save to Unit Flash` stays disabled). Until this is fixed in the code: after answering the last prompt, click the Power Rider Studio window once so it is in front, and leave it there.
