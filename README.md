# PowerGUIAutomation

GUI automation (FlaUI / UIA3, net10.0-windows) for Power Rider Studio. This file is updated continuously; for now it is general to the whole project, and per-test details will be added later.

## Before every run (general)

1. **Power Rider Studio is open and a unit is connected.** The tests attach to the already-running GUI.
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

## Before tests 2 and 3 (Configure)

**Reset the unit to default before the run.** The tests assume every parameter starts at its Excel default. The program asks `Was the unit RESET TO DEFAULT before this test? (Y/N)` and does not continue without `Y` (in a scripted run it only logs a warning).

## Tests (menu)

| # | Test | Notes |
|---|---|---|
| 1 | Operate - Second Bar (`OperateSecondBarTest`) | reads ids from the `Operate` sheet |
| 2 | Configure - Compare parameters to Excel (`ConfigureCompareParametersToExcelTest`) | compares GUI values with the Excel defaults |
| 3 | Configure - Change and restore every parameter (`ConfigureChangeAndRestoreAllParametersTest`) | **writes to the unit**: one parameter after the other (found by tree page scan): change, Save, Fetch, verify, then restore the Excel default, Save, Fetch, verify. A ComboBox parameter tries every option in list order and is restored only after the last one. Asks only whether the unit was reset to default |
| 5 | Debug tests (`DebugTests`; read-only) | `1` (`DebugCompareSubsetTest`): test 2 only on chosen channels (`15`, `13,15`, `S,15`) and/or parameters whose name contains a text. `2` (`DebugAutomationIdScanTest`): AutomationId scan of all parameters; the report goes to a new workbook in the `Excel` folder |

Test 2 (and 5) compare by **page scan**: they select the tree node "System Parameters" / "Channel N Parameters", scroll that scope's parameter list and compare every wanted row as it appears (a row not seen is looked up with the search box). This replaced one search per parameter, which took hours.

Test 2 (and 5) log, for every channel parameter, what the GUI itself shows as the displayed channel (`breadcrumb='… Channel 15 Parameters' | result 15 of 16`). If the wanted channel cannot be reached the test fails; it never compares against another channel's value.

Per-test details: to be added.

## Helper files

- The run log is saved next to the exe; console colors are display only.
