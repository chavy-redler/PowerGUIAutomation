using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using ClosedXML.Excel;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;

namespace PowerGUIAutomation.Tests;

// "Change and restore EVERY parameter" test of the Configure tab, one parameter after the other, in the order of the unit's Excel sheet
// (System, then channel 1..N, then the Group Control page). Parameters are found the same fast way as in ConfigureCompareParametersToExcelTest:
// select the tree node of the scope ("System Parameters" / "Channel N Parameters" / "Group Channel Control") and scroll its list
// (see ConfigureNav); the search box is only the fallback.
//
// For every parameter:
//   1. find it and read its current value + the Constraints text of its row (Min / Max / Default)
//   2. pick a different valid value and set it
//   3. Save (Save to Unit Flash), Fetch Parameters, verify the new value is really what the unit returns
//   4. set the Excel default (the "Value (default)" column) back, Save, Fetch, verify it is back at the default
// The restore is attempted even if the change failed. The test stops at the first failure.
//
// New value: numeric = random inside the row's [Min, Max] (different from the current one); no range = by unit ([msec]=10, [sec]=1,
// [%] = current +-1); text = QA.... A choice parameter (ComboBox / check box) is different: EVERY other option is tried in list order
// (each one: set, Save, Fetch, verify) and the default is restored only once, after the last option. Parameters that cannot be tested are skipped with the reason:
// CAN Bus Baud Rate, Serial Baud Rate, Destination Address, Unit ID Address, Logic Control / Logic Expression, Channel N Group Control, read-only controls
// and parameters with no range and no unit rule.
//
// SAFETY: this test WRITES parameters to the connected unit and saves them to flash (Reset Unit only if DoReset is on). There is no
// confirmation question; the "Before test 3" frame of the main menu (TestInstructions.cs) tells the user to reset the unit to default first.
public class ConfigureChangeAndRestoreAllParametersTest
{
    private const int MaxFindRetries = 12;
    private const int RetryDelayMs = 400;
    // Reset Unit takes a long time (the unit restarts and the GUI reconnects) and is OFF. Without it the test proves that the unit
    // ACCEPTED the saved value (Fetch reads it back from the unit) but not that it survives a restart.
    private const bool DoReset = false;
    private const int SaveTimeoutMs = 20000;
    private const int ResetTimeoutMs = 45000;
    private const int FetchTimeoutMs = 25000;

    private static readonly HashSet<string> NeverChange = new(StringComparer.OrdinalIgnoreCase)
    {
        "CAN Bus Baud Rate", "Serial Baud Rate", "Destination Address", "Unit ID Address",
    };

    // never changed by this test: the NeverChange names, Logic Control / Logic Expression, and every "Channel N Group Control" (Group Channel Control page)
    private static bool NotChanged(Param p) => NeverChange.Contains(p.Name) || IsLogicParam(p.Name) || p.Scope == ConfigureNav.GroupScope;

    private static readonly Regex BracketedRegex = new(@"^\[\s*([^,\]]+?)\s*[,\]]");

    private readonly string _excelPath;
    private readonly Action<string> _log;
    private readonly UnitProfile _unit;
    private readonly Random _random = new();

    // First failure of the run, for the closing "RESULT" explanation.
    private string _failParam = "", _failStep = "", _failWhy = "", _failCheck = "";
    private string _guiVersion = "";
    private readonly Dictionary<int, double> _lastPct = new();      // per scope: scroll position where the previous parameter was found
    private Dictionary<int, List<int>> _scopeRows = new();           // per scope: the Excel rows in list order (position of a row in the page = its index here)

    private AutomationElement? _logElement;

    public ConfigureChangeAndRestoreAllParametersTest(string excelPath, Action<string> log, UnitProfile unit)
    {
        _excelPath = excelPath;
        _log = log;
        _unit = unit;
    }

    private string[] Sheets => new[] { _unit.ParameterSheet };
    private bool _guiSaveNoteShown;

    // "stop on fail" (default) halts at the first failure; "continue" goes on with the next parameter and lists every failure at the end
    // (log + a workbook in the Failures folder next to the logs).
    private bool _stopOnFail = true;
    private const int MaxFailuresInARow = 5;      // continue mode only: this many in a row means a problem with the unit/GUI, not with one parameter

    private sealed record FailedParam(Param P, string Step, string Why, bool Restored);

    // What the test is doing right now (shown by the heartbeat and in the failure when nothing moves for NoProgressMinutes).
    private volatile string _where = "";
    private long _progressAt = Environment.TickCount64;
    private const int NoProgressMinutes = 3;

    private void At(string what)
    {
        _where = what;
        _progressAt = Environment.TickCount64;
    }

    // Runs one parameter on a worker thread: every 30 s without a new step it logs where it is; after NoProgressMinutes without any progress the
    // run is declared stuck (a UIA call that never returns cannot be interrupted, so the run stops instead of hanging for an hour).
    private Outcome RunWatched(Func<Outcome> action, out bool stuck)
    {
        stuck = false;
        At("starting");
        var task = Task.Run(action);
        long lastBeat = Environment.TickCount64;
        while (!task.Wait(1000))
        {
            long now = Environment.TickCount64;
            if (now - _progressAt > NoProgressMinutes * 60_000L)
            {
                stuck = true;
                return Outcome.Failed;
            }

            if (now - lastBeat >= 30_000 && now - _progressAt >= 30_000)
            {
                lastBeat = now;
                _log("INFO  | still working - " + ((now - _progressAt) / 1000) + " s on: " + _where);
            }
        }

        return task.GetAwaiter().GetResult();
    }

    private readonly List<FailedParam> _failures = new();
    private string _curStep = "", _curWhy = "";      // first failure of the parameter that is being tested

    // The failures workbook of THIS run (fixed name, rewritten after every failure and at the end - see RunReport) and what it needs to say about
    // a run that did not reach its end: the parameters of the run, how many are finished, and the one that is being tested right now.
    private const string FilePrefix = "ChangeAndRestore";
    private RunReport? _report;
    private List<Param> _todo = new();
    private int _finished;
    private Param? _current;
    private string _rowsDescription = "all";
    private int _checked, _passedCount, _skippedCount;

    // Recorded under the report's lock: the Ctrl+C / console-close handler writes the workbook from another thread.
    private void AddFailure(FailedParam f)
    {
        if (_report == null)
        {
            _failures.Add(f);
            return;
        }

        lock (_report.Sync)
        {
            _failures.Add(f);
        }

        WriteFailuresFile();      // the file is current after every failure, not only at the end of a (possibly hours long) run
    }

    // Ctrl+C or closing the console: the run ends here, so the file must say so. The parameter that was being tested is listed as failed -
    // it may have been left on a changed value on the unit (the test did not get to restore it).
    private void Interrupted(string status)
    {
        var report = _report;
        if (report == null)
        {
            return;
        }

        lock (report.Sync)
        {
            if (report.Finished)
            {
                return;
            }

            if (_current != null)
            {
                _failures.Add(new FailedParam(_current, "interrupted", "the run was stopped while this parameter was being changed or restored - its value on the unit is UNKNOWN, check it", false));
            }

            report.Finish(status);
        }

        _log("STOPPED | " + status + (_current != null ? " | interrupted while testing: " + _current.Label + " - check this parameter on the unit" : ""));
        WriteFailuresFile();
        _log("RESULT| run time        : " + RunReport.Format(report.Elapsed));
        _log("RESULT| failures file   : " + report.Path);
    }

    // ------------------------------------------------------------------ model

    private sealed class Param
    {
        public string Sheet = "", Name = "", NameId = "", ValueId = "", Default = "";
        public int ExcelRow, Channel;                 // Channel 0 = system parameter
        public int Scope => ConfigureNav.ScopeOf(Channel, Name);     // tree page: 0 = System, 1..N = channel, ConfigureNav.GroupScope = Group Channel Control
        public string Label => (Channel > 0 ? "Channel " + Channel + " | " : "System | ") + Name + " (" + Sheet + " row " + ExcelRow + ")";
    }

    private sealed class Tip
    {
        public string Text = "";
        public double? Min, Max;
        public string DefaultText = "";
        public string TypeToken = "";
        public bool Found;
    }

    // One place that reports a failure the same way every time: which parameter, which STEP of the test,
    // WHY it failed and what to check. The first one is repeated in the closing RESULT lines.
    private void Fail(Param? p, string step, string why, string check)
    {
        string who = p?.Label ?? "(test setup)";
        _log("FAIL  | " + who + " | STEP: " + step + " | WHY: " + why + " | CHECK: " + check);
        if (p != null && _curStep.Length == 0)
        {
            _curStep = step;
            _curWhy = why;
        }

        if (_failStep.Length == 0)
        {
            _failParam = who;
            _failStep = step;
            _failWhy = why;
            _failCheck = check;
        }
    }

    // 'status' = how the run ended, for the log and the Run sheet of the failures file; the default covers the usual endings.
    private void LogResult(bool passed, int done, int passedCount, bool restored, int skipped, string? status = null)
    {
        _checked = done;
        _passedCount = passedCount;
        _skippedCount = skipped;
        status ??= passed ? "Completed - no failures"
                 : done == 0 ? "Test setup failed - nothing was checked"
                 : _stopOnFail ? "Stopped after the first failure" : "Completed - with failures";
        _report?.Finish(status);
        _log("RESULT| " + new string('=', 90));
        _log("RESULT| " + (passed ? "PASSED" : "FAILED") + " | unit=" + _unit.Name + " | parameters changed+restored=" + passedCount + " of " + done + " checked | failed=" + _failures.Count + " | skipped=" + skipped
             + " | mode=" + (_stopOnFail ? "stop on fail" : "continue"));
        _log("RESULT| status          : " + status);
        if (_report != null)
        {
            _log("RESULT| total run time  : " + RunReport.Format(_report.Elapsed) + "  (" + _report.Started.ToString("HH:mm:ss") + " - " + (_report.Ended ?? DateTime.Now).ToString("HH:mm:ss") + ")");
        }

        if (!passed && _stopOnFail && _failStep.Length > 0)
        {
            _log("RESULT| failed parameter : " + _failParam);
            _log("RESULT| failed at step   : " + _failStep);
            _log("RESULT| reason           : " + _failWhy);
            _log("RESULT| what to check    : " + _failCheck);
            _log("RESULT| unit state       : " + (_failParam == "(test setup)" ? "nothing was changed on the unit"
                                                       : restored ? "parameter is back at its Excel default (verified)" : "NOT verified - check this parameter on the unit manually"));
        }

        if (_failures.Count > 0 && !_stopOnFail)
        {
            _log("RESULT| failed parameters (" + _failures.Count + "):");
            foreach (var f in _failures)
            {
                _log("RESULT|   row " + f.P.ExcelRow + " | " + (f.P.Channel > 0 ? "Channel " + f.P.Channel : "System") + " | " + f.P.Name + " | " + f.Step + " | " + Short(f.Why, 150)
                     + (f.Restored ? "" : " | DEFAULT NOT VERIFIED - check on the unit"));
            }
        }

        string? file = WriteFailuresFile();
        if (file != null)
        {
            _log("RESULT| failures file : " + file + (_failures.Count > 0 ? "   (to re-run only these rows give this file at 'Excel rows to run')" : ""));
        }

        _log("RESULT| " + new string('=', 90));
    }

    private static string Short(string text, int max) => text.Length <= max ? text : text.Substring(0, max) + "...";

    // One row per failed parameter (Excel row first, so a later run can take its parameters from this file - give the file at 'Excel rows to run'):
    // Failures\ChangeAndRestore_<PN>_<run start>.xlsx next to the log files. Sheets: Failed, Run (status / time), Not run (rows never reached).
    // Called after every failure and at the end, whatever the end is. Written when something failed or the run did not complete normally.
    private string? WriteFailuresFile()
    {
        var report = _report;
        if (report == null)
        {
            return null;
        }

        try
        {
            lock (report.Sync)
            {
                bool normalEnd = report.Finished && _failures.Count == 0 && report.Status.StartsWith("Completed", StringComparison.Ordinal);
                if (normalEnd)
                {
                    try { File.Delete(report.Path); } catch { /* the "running" version written at the start - nothing failed, nothing to keep */ }
                    return null;
                }

                using var wb = new XLWorkbook();
                var ws = wb.AddWorksheet(RunReport.SheetFailed);
                string[] headers = { "Excel row", "Channel", "Parameter", "Why it failed", "Failed step", "Default restored", "Unit", "GUI version", "Run time" };
                for (int i = 0; i < headers.Length; i++)
                {
                    ws.Cell(1, i + 1).Value = headers[i];
                    ws.Cell(1, i + 1).Style.Font.Bold = true;
                }

                int r = 2;
                foreach (var f in _failures)
                {
                    ws.Cell(r, 1).Value = f.P.ExcelRow;
                    ws.Cell(r, 2).Value = f.P.Channel > 0 ? f.P.Channel.ToString() : "System";
                    ws.Cell(r, 3).Value = f.P.Name;
                    ws.Cell(r, 4).Value = Short(f.Why, 300);
                    ws.Cell(r, 5).Value = f.Step;
                    ws.Cell(r, 6).Value = f.Restored ? "yes" : "NO";
                    ws.Cell(r, 7).Value = _unit.Name;
                    ws.Cell(r, 8).Value = _guiVersion;
                    ws.Cell(r, 9).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                    r++;
                }

                ws.Columns().AdjustToContents();

                // parameters of this run that were never reached (a finished run has none; a stopped one lists where to go on)
                var failedRows = new HashSet<int>(_failures.Select(f => f.P.ExcelRow));
                var notRun = _todo.Skip(_finished).Where(p => !failedRows.Contains(p.ExcelRow))
                                  .Select(p => (p.ExcelRow, p.Channel > 0 ? p.Channel.ToString() : "System", p.Name)).ToList();
                report.AddRunSheet(wb, new[]
                {
                    ("Test", "Configure - change and restore every parameter"),
                    ("Unit", _unit.Name),
                    ("GUI version", _guiVersion),
                    ("On failure", _stopOnFail ? "stop at the first failure" : "continue"),
                    ("Rows of this run", _rowsDescription),
                    ("Parameters in this run", _todo.Count.ToString()),
                    ("Checked / passed / skipped / failed", _checked + " / " + _passedCount + " / " + _skippedCount + " / " + _failures.Count),
                    ("Not reached", notRun.Count.ToString()),
                });
                RunReport.AddNotRunSheet(wb, notRun);
                report.Save(wb);
                return report.Path;
            }
        }
        catch (Exception ex)
        {
            _log("WARN  | could not write the failures file: " + ex.Message);
            return null;
        }
    }

    // ------------------------------------------------------------------ entry point

    // The run is wrapped so that Ctrl+C, closing the console and a crash all leave the failures file with the right status (see Interrupted).
    public bool Run()
    {
        ConsoleCancelEventHandler onCancel = (_, _) => Interrupted("Stopped by the user (Ctrl+C)");
        EventHandler onExit = (_, _) => Interrupted("Stopped - the console was closed or the program was ended");
        Console.CancelKeyPress += onCancel;
        AppDomain.CurrentDomain.ProcessExit += onExit;
        try
        {
            return RunCore();
        }
        catch (Exception ex)
        {
            if (_report != null && !_report.Finished)
            {
                // an exception that nothing caught: "test failed" in the middle - the file says so and names the parameter that was being tested
                if (_current != null)
                {
                    lock (_report.Sync)
                    {
                        _failures.Add(new FailedParam(_current, "unexpected error", ex.GetType().Name + ": " + ex.Message + " - the value on the unit is UNKNOWN, check it", false));
                    }
                }

                _log("FAIL  | the test crashed: " + ex.GetType().Name + ": " + ex.Message);
                LogResult(false, _checked, _passedCount, false, _skippedCount, "Test failed - crashed: " + ex.GetType().Name + ": " + Short(ex.Message, 120));
            }

            throw;
        }
        finally
        {
            Console.CancelKeyPress -= onCancel;
            AppDomain.CurrentDomain.ProcessExit -= onExit;
        }
    }

    private bool RunCore()
    {
        _guiVersion = ReadGuiVersion();
        if (_guiVersion.Length == 0)
        {
            _log("FAIL | Power Rider Studio is not running (cannot read its version).");
            return false;
        }

        Console.WriteLine();
        Console.WriteLine("Configure change-and-restore test | GUI version: " + _guiVersion);
        Console.WriteLine();
        Console.WriteLine("WARNING: this test WRITES parameters to the connected unit and saves them to flash (change, Save, Fetch, restore default).");
        Console.WriteLine();

        Console.WriteLine("When a parameter fails:");
        Console.WriteLine("  1. stop at the first failure  [default]");
        Console.WriteLine("  2. continue with the next parameter; every failure is listed at the end (and in the Failures folder)");
        Console.Write("Choice [1]: ");
        _stopOnFail = (Console.ReadLine() ?? "").Trim() != "2";

        var selection = RowSelection.Ask(FilePrefix, _unit.PartNumber, out _rowsDescription);
        Console.WriteLine();

        _report = new RunReport(FilePrefix, _unit.PartNumber);
        RunClock.Start();      // the menu's "took" time starts here: the questions above are not counted
        _log("START | Configure change-and-restore test | unit=" + _unit.Name + " (sheet '" + _unit.ParameterSheet + "') | GUI version=" + _guiVersion
             + " | on failure: " + (_stopOnFail ? "stop" : "continue") + " | rows: " + _rowsDescription);
        if (selection?.SourceFile != null)
        {
            _log("INFO  | RE-RUN of the failed rows listed in: " + selection.SourceFile);
        }

        _log("INFO  | this run's failures file (rewritten after every failure and at the end, also after Ctrl+C or a crash): " + _report.Path);
        WriteFailuresFile();      // exists from the first minute: status "Running" until the run ends (deleted again if nothing fails)

        List<Param> all;
        try
        {
            all = ReadParams();
        }
        catch (Exception ex)
        {
            Fail(null, "read the Excel", "could not read parameters from sheet '" + _unit.ParameterSheet + "': " + ex.Message,
                 "the workbook path, and that the sheet exists for unit '" + _unit.Name + "'");
            LogResult(false, 0, 0, true, 0);
            return false;
        }

        if (all.Count == 0)
        {
            Fail(null, "read the Excel", "no parameters were found for unit " + _unit.PartNumber + ": the sheet '" + _unit.ParameterSheet + "' is missing or empty in the workbook",
                 "create this sheet for " + _unit.PartNumber + " (same columns as the RD152 sheet) or pick another unit");
            LogResult(false, 0, 0, true, 0);
            return false;
        }

        all = all.Where(p => p.Channel <= _unit.Channels).ToList();      // a 12-channel unit has no channel 13..16
        _scopeRows = all.GroupBy(p => p.Scope).ToDictionary(g => g.Key, g => g.OrderBy(p => p.ExcelRow).Select(p => p.ExcelRow).ToList());
        var excluded = all.Where(NotChanged).ToList();
        var todo = all.Where(p => !NotChanged(p)).ToList();
        _log("INFO  | parameters in Excel=" + all.Count + " | never changed (baud rates, addresses, logic, group control)=" + excluded.Count + " | to check in this run=" + todo.Count);
        if (selection != null)
        {
            todo = todo.Where(p => selection.Contains(p.ExcelRow)).ToList();
            _log("INFO  | row selection '" + selection + "': " + todo.Count + " parameter(s) (rows of parameters that are never changed are ignored)");
            foreach (var mismatch in selection.NameMismatches(todo.Select(p => (p.ExcelRow, p.Name))))
            {
                _log("WARN  | " + mismatch + " - was the master workbook edited since that run?");
            }
        }

        _todo = todo;

        if (todo.Count == 0)
        {
            _log("FAIL  | nothing to check - no changeable parameter in the selected rows");
            return false;
        }

        using var automation = new UIA3Automation();
        var window = AppConnection.Attach(automation);

        string? unitProblem = VerifyUnit(window);
        if (unitProblem != null)
        {
            Fail(null, "check the connected unit", unitProblem, "connect the unit '" + _unit.Name + "' (Manage Units) or choose the right unit in the main menu");
            LogResult(false, 0, 0, true, 0);
            return false;
        }

        AppConnection.GoToTab(window, "ConfigureTab");
        if (!WaitFetchDone(window, 300000, "Configure tab just opened", quick: false))
        {
            Fail(null, "open the Configure tab", "the GUI kept 'Fetching parameters...' for more than 5 minutes", "is the unit connected and responding? (status bar shows BUS CONFLICT?)");
            LogResult(false, 0, 0, true, 0);
            return false;
        }

        var searchElement = window.FindFirstDescendant(cf => cf.ByAutomationId("searchTextBox"));
        if (searchElement == null)
        {
            Fail(null, "open the Configure tab", "the search box (searchTextBox) was not found", "is the Configure tab open and a unit connected?");
            LogResult(false, 0, 0, true, 0);
            return false;
        }

        var search = searchElement.AsTextBox();
        _logElement = window.FindFirstDescendant(cf => cf.ByAutomationId("ConfigureLog"));

        int done = 0, passed = 0, skipped = 0, index = 0, inARow = 0;
        var watch = Stopwatch.StartNew();
        foreach (var p in todo)
        {
            index++;
            _current = p;
            _log("DRAW  | " + index + "/" + todo.Count + " | " + p.Label);
            Outcome outcome;
            bool stuck = false;
            try
            {
                outcome = RunWatched(() => TestOne(window, search, p), out stuck);
            }
            catch (Exception ex)
            {
                Fail(p, "unexpected error", ex.GetType().Name + ": " + ex.Message, "see the message");
                _lastRestored = false;
                outcome = Outcome.Failed;
            }

            if (stuck)
            {
                Fail(p, "stuck", "no progress for " + NoProgressMinutes + " minutes - last action: " + _where, "the GUI may be frozen or a UI call is not returning; look at the GUI, restart it if needed and run again from this row");
                _lastRestored = false;
                _current = null;
                _finished = index;
                _failures.Add(new FailedParam(p, _curStep, _curWhy, false));
                _log("STOPPED | the run is stuck and cannot continue safely");
                _log("RESUME| to continue later answer 'Excel rows to run' with:  " + p.ExcelRow + "-  (after checking this parameter on the unit)   or give the failures file (only its failed rows)");
                LogResult(false, done + 1, passed, false, skipped, "Stopped - stuck (no progress for " + NoProgressMinutes + " minutes)");
                return false;
            }

            _current = null;
            _finished = index;
            if (outcome == Outcome.Skipped)
            {
                skipped++;
                continue;
            }

            done++;
            if (outcome == Outcome.Failed)
            {
                AddFailure(new FailedParam(p, _curStep, _curWhy, _lastRestored));
                inARow++;
                if (_stopOnFail)
                {
                    _log("STOPPED | halted after first failure - remaining parameters were not checked");
                    var nextParam = index < todo.Count ? todo[index] : null;
                    _log("RESUME| to continue later answer 'Excel rows to run' with:  " + p.ExcelRow + "-  (this parameter again)"
                         + (nextParam != null ? "   or   " + nextParam.ExcelRow + "-  (skip it, continue with " + nextParam.Name + ")" : ""));
                    ClearSearch(window);
                    LogResult(false, done, passed, _lastRestored, skipped);
                    return false;
                }

                _log("INFO  | failure " + _failures.Count + " recorded - continuing with the next parameter");
                if (inARow >= MaxFailuresInARow)
                {
                    _log("STOPPED | " + inARow + " parameters failed in a row - that looks like a problem with the unit or the GUI, not with the parameters");
                    ClearSearch(window);
                    LogResult(false, done, passed, _lastRestored, skipped, "Stopped - " + inARow + " parameters failed in a row (unit or GUI problem?)");
                    return false;
                }

                continue;
            }

            inARow = 0;
            passed++;
        }

        ClearSearch(window);
        _log("DONE  | checked=" + done + " passed=" + passed + " failed=" + _failures.Count + " skipped=" + skipped + " | " + RunReport.Format(watch.Elapsed)
             + " | tree pages selected: " + ConfigureNav.SelectScopeCalls + " | scroll positions tried: " + ConfigureNav.ScrollSteps);
        LogResult(_failures.Count == 0, done, passed, true, skipped);
        return _failures.Count == 0;
    }

    // ------------------------------------------------------------------ one parameter

    // Passed = changed and restored. Failed = anything else (the run stops at the first failure, the parameter is put back to its default first;
    // the closing lines say how to continue from this parameter with the "Start from" option).
    private enum Outcome { Passed, Failed, Skipped }

    private enum ChangeResult { Ok, NotAccepted, Hard }

    // a value that was not accepted / not kept is a failure of the test (first one is repeated in the closing RESULT lines)
    private void Finding(Param p, string step, string why)
    {
        Fail(p, step, why, "this value/option cannot be set or is not kept - a GUI/unit problem? (start from this row again after it is fixed)");
    }

    private bool _lastRestored = true;

    // The unit under test must be the unit that is connected: the Excel defaults are for THIS unit and the run writes to it.
    private string? VerifyUnit(Window window) => UnitCheck.Verify(window, _unit, _log);

    // Finds the parameter's name + value elements: select the scope's tree page and scroll to the row (the previous parameter's position is
    // the starting point, the rows come in list order); if that fails, the search box.
    private (AutomationElement? name, AutomationElement? value) LocateRow(Window window, TextBox search, Param p, bool refresh)
    {
        if (!WaitFetchDone(window, 300000, "before locating " + p.Name))
        {
            Fail(p, "1 locate the parameter", "the GUI kept 'Fetching parameters...' for more than 5 minutes", "is the unit connected and responding? (status bar shows BUS CONFLICT?)");
            return (null, null);
        }

        var found = LocateOnPage(window, p);
        if (found.name != null)
        {
            return found;
        }

        _log("INFO  | " + p.Label + " | not found by the page scan - using the search box");
        return Locate(window, search, p, refresh);
    }

    // Fast path of the locate: the parameter's row is already visible on the right page (the same row after Save / refresh, or the next row of the list).
    // The breadcrumb proves the page (ids repeat in every channel), the search box must be empty.
    private (AutomationElement? name, AutomationElement? value) TryFindVisible(Window window, Param p)
    {
        try
        {
            var list = window.FindFirstDescendant(cf => cf.ByAutomationId("ParametersList"));
            var nameEl = list?.FindFirstDescendant(cf => cf.ByAutomationId(p.NameId));
            var valueEl = list?.FindFirstDescendant(cf => cf.ByAutomationId(p.ValueId));
            if (nameEl == null || valueEl == null || (nameEl.Name ?? "").Trim() != p.Name || nameEl.Properties.IsOffscreen.ValueOrDefault || valueEl.BoundingRectangle.Width <= 0)
            {
                return (null, null);
            }

            var box = window.FindFirstDescendant(cf => cf.ByAutomationId("searchTextBox"));
            if (box != null && box.Patterns.Value.TryGetPattern(out var vp) && (vp.Value.ValueOrDefault ?? "").Length > 0)
            {
                return (null, null);
            }

            string crumb = ConfigureNav.ReadBreadcrumb(window);
            bool scopeOk = p.Scope == ConfigureNav.GroupScope ? crumb.Contains("Group Channel Control", StringComparison.Ordinal)
                : p.Scope > 0 ? crumb.Contains("Channel " + p.Scope + " Parameters", StringComparison.Ordinal)
                : crumb.Contains("System Parameters", StringComparison.Ordinal) && !crumb.Contains("Channel", StringComparison.Ordinal);
            return scopeOk ? (nameEl, valueEl) : (null, null);
        }
        catch (Exception ex) when (ex is FlaUI.Core.Exceptions.PropertyNotSupportedException or FlaUI.Core.Exceptions.ElementNotAvailableException or System.Runtime.InteropServices.COMException or InvalidOperationException)
        {
            return (null, null);
        }
    }

    // 1) the row is visible already; 2) the right page is shown: only scroll (from where the previous row was found, rows come in list order); 3) select the
    // tree page (this rebuilds the list) and scroll. Steps 1-2 skip the tree selection, which is the slow part.
    private (AutomationElement? name, AutomationElement? value) LocateOnPage(Window window, Param p)
    {
        var visible = TryFindVisible(window, p);
        if (visible.name != null)
        {
            return visible;
        }

        try
        {
            if (PageIsShown(window, p))
            {
                var scanned = ScanPages(window, p);
                if (scanned.name != null)
                {
                    return scanned;
                }
            }

            string? crumb = ConfigureNav.SelectScope(window, p.Scope);      // re-selecting rebuilds the list
            if (crumb == null)
            {
                return (null, null);
            }

            return ScanPages(window, p);
        }
        catch (Exception ex) when (ex is FlaUI.Core.Exceptions.PropertyNotSupportedException or FlaUI.Core.Exceptions.ElementNotAvailableException or System.Runtime.InteropServices.COMException or InvalidOperationException)
        {
            _log("NOTE  | " + p.Label + " | page scan interrupted (" + ex.GetType().Name + ")");
        }

        return (null, null);
    }

    // The breadcrumb says the parameter's page is shown and the search box is empty (so the list is that page, not a search result).
    private bool PageIsShown(Window window, Param p)
    {
        try
        {
            var box = window.FindFirstDescendant(cf => cf.ByAutomationId("searchTextBox"));
            if (box != null && box.Patterns.Value.TryGetPattern(out var vp) && (vp.Value.ValueOrDefault ?? "").Length > 0)
            {
                return false;
            }

            string crumb = ConfigureNav.ReadBreadcrumb(window);
            return p.Scope == ConfigureNav.GroupScope ? crumb.Contains("Group Channel Control", StringComparison.Ordinal)
                : p.Scope > 0 ? crumb.Contains("Channel " + p.Scope + " Parameters", StringComparison.Ordinal)
                : crumb.Contains("System Parameters", StringComparison.Ordinal) && !crumb.Contains("Channel", StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
    }

    // Scrolls the shown page in half-page steps until the row is found.
    private (AutomationElement? name, AutomationElement? value) ScanPages(Window window, Param p)
    {
        var list = ConfigureNav.WaitListReady(window, 20000, _log, p.Label);
        if (list == null)
        {
            _log("NOTE  | " + p.Label + " | the parameter list showed no parameters within 20 s");
            return (null, null);
        }

        double view = 100;
        var scroll = list.Patterns.Scroll.PatternOrDefault;
        bool scrollable = scroll != null && scroll.VerticallyScrollable.ValueOrDefault;
        if (scrollable)
        {
            view = Math.Max(5, scroll!.VerticalViewSize.ValueOrDefault);
        }

        double estimate = _scopeRows.TryGetValue(p.Scope, out var rowsOfScope)
            ? ConfigureNav.EstimateScrollPercent(rowsOfScope.Count, rowsOfScope.IndexOf(p.ExcelRow), view) : -1;
        double lastPct = _lastPct.TryGetValue(p.Scope, out double lp) ? lp : -1;
        foreach (double pct in ConfigureNav.ScrollTries(estimate, lastPct, view))
        {
            list = window.FindFirstDescendant(cf => cf.ByAutomationId("ParametersList")) ?? list;
            scroll = list.Patterns.Scroll.PatternOrDefault;
            scrollable = scroll != null && scroll.VerticallyScrollable.ValueOrDefault;      // decided again on every step (the list may become scrollable late)
            if (scrollable)
            {
                view = Math.Max(5, scroll!.VerticalViewSize.ValueOrDefault);
                ConfigureNav.ScrollSteps++;
                scroll!.SetScrollPercent(-1, pct);
                Thread.Sleep(150);
            }

            var nameEl = list.FindFirstDescendant(cf => cf.ByAutomationId(p.NameId));
            var valueEl = list.FindFirstDescendant(cf => cf.ByAutomationId(p.ValueId));
            if (nameEl != null && valueEl != null && (nameEl.Name ?? "").Trim() == p.Name)
            {
                _lastPct[p.Scope] = pct;
                return (nameEl, valueEl);
            }

            if (!scrollable)
            {
                break;
            }
        }

        return (null, null);
    }

    private Outcome TestOne(Window window, TextBox search, Param p)
    {
        _curStep = "";
        _curWhy = "";
        _lastRestored = true;

        _log("STEP  | 1 find the parameter (tree page scan)");
        At("1 locating the parameter");
        var (nameEl, valueEl) = LocateRow(window, search, p, refresh: false);
        if (nameEl == null || valueEl == null)
        {
            Fail(p, "1 locate the parameter", "the parameter row was not found on screen (ids " + p.NameId + " / " + p.ValueId + ")",
                 "is the Configure tab open? does the page '" + (p.Channel > 0 ? "Channel " + p.Channel : "System") + "' list the parameter?");
            return Outcome.Failed;
        }

        string type = valueEl.ControlType.ToString();

        _log("STEP  | 2 read the current value and the constraints (range)");
        At("2 reading the current value");
        string before = ReadParam(window, p, out var freshValueEl, 4);
        valueEl = freshValueEl ?? valueEl;
        if (before.Length == 0)
        {
            // an empty field can simply be a value that was not loaded yet: use the parameter's own refresh button once
            _log("INFO  | " + p.Label + " | the field is empty - pressing the parameter's own 'Refresh parameter value from unit' button once");
            if (RefreshParam(valueEl))
            {
                Thread.Sleep(2500);
                before = ReadParam(window, p, out freshValueEl, 6);
                valueEl = freshValueEl ?? valueEl;
            }
        }

        if (before.Length == 0)
        {
            Fail(p, "2 read the current value", "the GUI shows an EMPTY value for this " + type + " (the field is blank, also after the parameter's refresh button); the Excel default is '" + p.Default + "'",
                 "is this value missing in the GUI/unit, or not shown for this channel? compare with the other channels of the same parameter");
            return Outcome.Failed;
        }

        if (!SameValue(before, p.Default))
        {
            _log("WARN  | " + p.Label + " | before the test the value is '" + before + "' but the Excel default is '" + p.Default + "' (the unit is not at its default)");
        }

        // The constraints text is already inside the parameter's row (hidden) - read it from there; hovering is only the fallback.
        bool isChoiceControl = valueEl.ControlType == ControlType.ComboBox || valueEl.ControlType == ControlType.CheckBox;
        var tip = new Tip();
        if (!isChoiceControl)
        {
            At("2 reading the constraints text of the row");
            tip = ReadConstraintsFromRow(valueEl);
            if (!tip.Found)
            {
                At("2 hovering the parameter name to read its tooltip");
                tip = ReadTooltip(window, nameEl);
            }
        }

        _log("READ  | " + p.Label + " | control=" + type + " | current='" + before + "' | excel default='" + p.Default + "'"
             + (tip.Found ? " | constraints: Min=" + Fmt(tip.Min) + " Max=" + Fmt(tip.Max) + " Default=" + tip.DefaultText + " unit/type=" + tip.TypeToken : " | (no constraints text)"));

        if (IsReadOnly(valueEl))
        {
            _log("SKIP  | " + p.Label + " | value control is read-only");
            return Outcome.Skipped;
        }

        _log("STEP  | 3 change the value (choose a new value and set it in the GUI)");
        // Choice parameters (ComboBox / check box): EVERY other option is tried, in list (index) order, without restoring the default in
        // between; the default is restored once, after the last option. Any other parameter gets one new value.
        var plan = new List<(string value, string why)>();
        if (valueEl.ControlType == ControlType.ComboBox || valueEl.ControlType == ControlType.CheckBox)
        {
            var options = valueEl.ControlType == ControlType.ComboBox ? ReadOptions(valueEl) : new List<string> { "On", "Off" };
            string cur = Normalize(before);
            var targets = options.Where(o => !SameValue(o, cur)).ToList();
            if (targets.Count == 0)
            {
                _log("SKIP  | " + p.Label + " | choice parameter with fewer than 2 options");
                return Outcome.Skipped;
            }

            for (int i = 0; i < targets.Count; i++)
            {
                plan.Add((targets[i], "option " + (i + 1) + " of " + targets.Count + " (list order) | all options=" + string.Join(", ", options)));
            }
        }
        else
        {
            string? newValue = ChooseNewValue(valueEl, p, before, tip, out string why);
            if (newValue == null)
            {
                _log("SKIP  | " + p.Label + " | " + why);
                return Outcome.Skipped;
            }

            plan.Add((newValue, why));
        }

        bool ok = true;
        try
        {
            for (int i = 0; i < plan.Count && ok; i++)
            {
                if (i > 0)
                {
                    // the previous option was saved and fetched: find the row again
                    var (_, again) = LocateRow(window, search, p, refresh: true);
                    if (again == null)
                    {
                        Fail(p, "3 set the new value", "the parameter was not found again before trying option " + (i + 1), "check the Configure screen");
                        ok = false;
                        break;
                    }

                    valueEl = again;
                }

                var result = ChangeAndVerify(window, search, p, valueEl, before, plan[i].value, plan[i].why);
                if (result == ChangeResult.Hard)
                {
                    ok = false;
                }
                else if (result == ChangeResult.NotAccepted)
                {
                    ok = false;
                }
                else
                {
                    before = plan[i].value;
                }
            }
        }
        catch (Exception ex)
        {
            Fail(p, "3-5 change the parameter", "exception: " + ex.Message, "see the message");
            ok = false;
        }

        if (!ok)
        {
            // an unsaved/unregistered edit may still be pending: Fetch discards pending edits
            var fetch = FindByHelp(window, "Refresh parameters from unit");
            if (fetch != null && fetch.IsEnabled)
            {
                _log("CLICK | " + p.Label + " | Fetch (discard a pending edit after the failure)");
                fetch.AsButton().Invoke();
                Thread.Sleep(3000);
            }
        }

        _log("STEP  | 6 restore the Excel default '" + p.Default + "' and verify it");
        At("6 restoring the default '" + p.Default + "'");
        bool restored = RestoreDefault(window, search, p);
        _lastRestored = restored;
        return ok && restored ? Outcome.Passed : Outcome.Failed;
    }

    // One change: set `newValue`, Save, Fetch and verify that the unit returns it. False = failed (reason already logged with Fail).
    private ChangeResult ChangeAndVerify(Window window, TextBox search, Param p, AutomationElement valueEl, string before, string newValue, string why)
    {
        At("3 setting the value '" + newValue + "'");
        SetValue(window, valueEl, newValue);
        string shownNow = ReadParam(window, p, out _);
        _log("SET   | " + p.Label + " | '" + before + "' -> '" + newValue + "' | the GUI now shows '" + shownNow + "' | " + why
             + " | Save to Unit Flash is " + (FindByHelp(window, "Save to Unit Flash")?.IsEnabled == true ? "enabled" : "DISABLED"));
        if (!SameValue(shownNow, newValue))
        {
            Finding(p, "3 set the new value '" + newValue + "'", "after setting it the control shows '" + shownNow + "' - the value was rejected/not committed (range, focus)");
            return ChangeResult.NotAccepted;
        }

        if (!WaitSaveEnabled(window) && valueEl.ControlType == ControlType.Edit)
        {
            // the edit may not have been committed: click into the field again and press Enter once more
            _log("NOTE  | " + p.Label + " | Save to Unit Flash is still disabled - pressing Enter in the field again");
            try
            {
                valueEl.Focus();
            }
            catch (Exception ex) when (ex is InvalidOperationException or System.Runtime.InteropServices.COMException)
            {
                // focus not possible right now
            }

            Keyboard.Type(VirtualKeyShort.ENTER);
            Thread.Sleep(500);
        }

        if (!WaitSaveEnabled(window))
        {
            Finding(p, "3 set the new value '" + newValue + "'", "the value shows '" + shownNow + "' but 'Save to Unit Flash' stayed DISABLED - the GUI did not register a pending change");
            return ChangeResult.NotAccepted;
        }

        _log("STEP  | 4 Save to Unit Flash" + (DoReset ? ", Reset Unit" : "") + ", read the value back from the unit");
        At("4 Save and read back '" + newValue + "'");
        if (!SaveResetFetch(window, search, p))
        {
            return ChangeResult.Hard;
        }

        _log("STEP  | 5 verify the value after Save / Fetch");
        At("5 verifying '" + newValue + "'");
        var (_, v2) = LocateRow(window, search, p, refresh: true);
        string after = v2 == null ? "<NOT FOUND>" : ReadParam(window, p, out _);
        bool changed = SameValue(after, newValue);
        _log((changed ? "PASS  | " : "FAIL  | ") + p.Label + " | after Save/Fetch: expected='" + newValue + "' actual='" + after + "'");
        if (!changed)
        {
            Finding(p, "5 verify after Save/Fetch", "expected '" + newValue + "' but the unit returned '" + after + "' - the unit did not keep the value");
            return ChangeResult.NotAccepted;
        }

        return ChangeResult.Ok;
    }

    private bool WaitSaveEnabled(Window window)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < 6000)
        {
            var save = FindByHelp(window, "Save to Unit Flash");
            if (save != null && save.IsEnabled)
            {
                return true;
            }

            Thread.Sleep(400);
        }

        return false;
    }

    private bool RestoreDefault(Window window, TextBox search, Param p)
    {
        try
        {
            var (_, v) = LocateRow(window, search, p, refresh: true);
            if (v == null)
            {
                Fail(p, "6 restore default", "cannot restore the default: the parameter was not found again", "check the Configure screen manually");
                return false;
            }

            string now = ReadParam(window, p, out var freshV);
            v = freshV ?? v;
            if (now.Length == 0)
            {
                Fail(p, "6 restore default", "cannot read the current value to compare with the default", "check the parameter on the unit manually");
                return false;
            }

            if (SameValue(now, p.Default))
            {
                _log("PASS  | " + p.Label + " | at the Excel default '" + p.Default + "'");
                return true;
            }

            _log("RESTORE| " + p.Label + " | '" + now + "' -> default '" + p.Default + "'");
            SetValue(window, v, p.Default);
            if (!SaveResetFetch(window, search, p))
            {
                return false;
            }

            var (_, v2) = LocateRow(window, search, p, refresh: true);
            string after = v2 == null ? "<NOT FOUND>" : ReadParam(window, p, out _);
            bool isDefault = SameValue(after, p.Default);
            _log((isDefault ? "PASS  | " : "FAIL  | ") + p.Label + " | back to default: expected='" + p.Default + "' actual='" + after + "'");
            if (!isDefault)
            {
                Fail(p, "6 restore default", "after restoring, the value is '" + after + "' instead of the default '" + p.Default + "'", "the unit still holds a changed value - fix it manually");
            }

            return isDefault;
        }
        catch (Exception ex)
        {
            Fail(p, "6 restore default", "exception: " + ex.Message, "check the parameter on the unit manually");
            return false;
        }
    }

    // ------------------------------------------------------------------ value choice

    private string? ChooseNewValue(AutomationElement valueEl, Param p, string current, Tip tip, out string why)
    {
        why = "";
        if (Normalize(current).Length == 0)
        {
            why = "the current value is unknown, so a different value cannot be chosen";
            return null;
        }

        // Edit / text box
        string cleaned = Normalize(current);
        // A text value is a string whether or not the tooltip has Min/Max (for a string they are the allowed LENGTH) and whether or not the
        // data-type label ("String") could be read from the row.
        bool isString = tip.TypeToken.Equals("String", StringComparison.OrdinalIgnoreCase)
                        || !double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
        if (isString)
        {
            int min = (int)(tip.Min ?? 1), max = (int)(tip.Max ?? 12);
            max = Math.Min(max, 12);
            int len = Math.Max(min, Math.Min(max, 6));
            string candidate;
            do { candidate = "QA" + new string(Enumerable.Range(0, Math.Max(0, len - 2)).Select(_ => (char)('A' + _random.Next(26))).ToArray()); }
            while (candidate == current && len > 0 && max > 2 && _random.Next(100) != 0);
            why = "text of length " + candidate.Length + " (allowed " + min + ".." + max + ")";
            return candidate.Length >= min ? candidate : null;
        }

        if (!tip.Min.HasValue || !tip.Max.HasValue)
        {
            // no range in the tooltip: take a value from the UNIT shown there (table below); a different value than the current one
            if (!double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out double curNum))
            {
                why = "no Min/Max in the tooltip and the current value '" + current + "' is not numeric";
                return null;
            }

            // [%] without a range (Log. I TH Event, Soft-Start/Stop PWM Duty Cycle): current value +1 or -1 (never below 0)
            if (tip.TypeToken.Trim().Trim('[', ']') == "%")
            {
                int pctDec = cleaned.Contains('.') ? cleaned.Length - cleaned.IndexOf('.') - 1 : 0;
                double pctNew = curNum - 1 >= 0 && _random.Next(2) == 0 ? curNum - 1 : curNum + 1;
                why = "no Min/Max in the tooltip - [%] parameter: current value " + (pctNew > curNum ? "+1" : "-1");
                return pctNew.ToString("F" + pctDec, CultureInfo.InvariantCulture);
            }

            double? byUnit = UnitTestValue(tip.TypeToken);
            if (byUnit == null)
            {
                why = "no Min/Max in the tooltip and no test value is defined for the unit '" + tip.TypeToken + "' (see UnitTestValue)";
                return null;
            }

            double chosen = Math.Abs(byUnit.Value - curNum) < 1e-9 ? byUnit.Value * 2 : byUnit.Value;     // already equal -> another value
            int dec = cleaned.Contains('.') ? cleaned.Length - cleaned.IndexOf('.') - 1 : 0;
            why = "no Min/Max in the tooltip - value chosen by its unit " + tip.TypeToken;
            return chosen.ToString("F" + dec, CultureInfo.InvariantCulture);
        }

        double lo = tip.Min.Value, hi = tip.Max.Value;
        if (!double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out double cur0))
        {
            why = "current value '" + current + "' is not numeric";
            return null;
        }

        if (hi - lo < 1e-9)
        {
            why = "Min equals Max - nothing to change";
            return null;
        }

        // Some constraints depend on OTHER parameters: "This value must be lower then Vin2 High Threshold P1/P2/P3" (or "higher"). The GUI rejects
        // a value that breaks it (Save stays disabled), and the other parameters are sitting at their own (valid) values - so only move in the
        // allowed direction: "lower" = between Min and the current value, "higher" = between the current value and Max.
        string direction = "";
        var dep = Regex.Match(tip.Text ?? "", @"must be\s+(lower|less|smaller|higher|greater|bigger)", RegexOptions.IgnoreCase);
        if (dep.Success)
        {
            bool lower = dep.Groups[1].Value.ToLowerInvariant() is "lower" or "less" or "smaller";
            if (lower)
            {
                hi = Math.Min(hi, cur0);
            }
            else
            {
                lo = Math.Max(lo, cur0);
            }

            direction = " | the constraint says 'must be " + dep.Groups[1].Value.ToLowerInvariant() + " than ...' (another parameter), so only " + (lower ? "below" : "above") + " the current value";
            if (hi - lo < 1e-9)
            {
                why = "the value cannot move " + (lower ? "below" : "above") + " the current value inside " + Fmt(tip.Min) + ".." + Fmt(tip.Max) + " (constraint depends on another parameter)";
                return null;
            }
        }

        // The new value is CLOSE to the current one (about +-10 %, at least 3 steps), inside Min..Max. A value picked anywhere in a huge range (e.g. up to 4,294,967,295)
        // is rounded by the GUI (float precision above 16,777,216) or refused by the unit (limits that depend on other parameters) - that is not what this test is about;
        // the limits themselves are the job of the input-validation test.
        int decimals = cleaned.Contains('.') ? cleaned.Length - cleaned.IndexOf('.') - 1 : 0;
        double step = decimals == 0 ? 1 : Math.Pow(10, -decimals);
        double span = Math.Min(Math.Max(step * 3, Math.Abs(cur0) * 0.1), Math.Min(hi - lo, 1_000_000));
        double value = cur0;
        for (int i = 0; i < 60 && Math.Abs(value - cur0) < step / 2; i++)
        {
            double raw = cur0 + (_random.NextDouble() * 2 - 1) * span;
            value = decimals == 0 ? Math.Round(raw) : Math.Round(raw, decimals);
            value = Math.Max(lo, Math.Min(hi, value));
        }

        if (Math.Abs(value - cur0) < step / 2)
        {
            why = "could not find a different value near " + cleaned + " inside " + Fmt(lo) + ".." + Fmt(hi);
            return null;
        }

        why = "value near the current one (+-" + Fmt(span) + ") inside " + Fmt(lo) + ".." + Fmt(hi) + direction;
        return value.ToString("F" + decimals, CultureInfo.InvariantCulture);
    }

    // ------------------------------------------------------------------ logic parameters

    // "Digital Output N Logic Control" and "Logic Expression" take free logic text with no Min/Max - this test does not change them
    // (excluded from the draw, like the baud rates and addresses in NeverChange).
    private static bool IsLogicParam(string name) => name.Contains("Logic Control", StringComparison.OrdinalIgnoreCase) || name.Equals("Logic Expression", StringComparison.OrdinalIgnoreCase);

    // ------------------------------------------------------------------ UI helpers

    private static readonly Regex MatchesRegex = new(@"^(\d+) of (\d+) matches", RegexOptions.Compiled);

    // The Configure search shows ONE result at a time ("3 of 16 matches"); "Next Result" moves to the next one and the
    // breadcrumb ("... > Channel 3 Parameters > ...") tells which scope is shown. So: search, then press Next until the
    // breadcrumb is the wanted scope (System Parameters / Channel N Parameters).
    private static readonly Regex CrumbChannelRegex = new(@"Channel (\d+) Parameters", RegexOptions.Compiled);

    private static readonly Regex FetchingRegex = new(@"Fetching parameters\.\.\.\s*(\d+)\s*/\s*(\d+)", RegexOptions.Compiled);

    // After opening the Configure tab (or Fetch / Reset) the GUI keeps loading the unit's parameters in the background and
    // shows "Fetching parameters... 5/688" in the status bar. Until that is finished the channel values are BLANK and the
    // search is disabled - so every locate/read must wait for it to disappear.
    private string? _staleProgress;      // a status text already judged "stale" (same counter, search box usable): not waited for again

    private bool WaitFetchDone(Window window, int timeoutMs, string reason, bool quick = true)
    {
        var sw = Stopwatch.StartNew();
        long lastLog = -100000;
        long? absentSince = null;
        bool announced = false;
        string? lastProgress = null;
        long progressSince = 0;
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            string? progress = null;
            try
            {
                foreach (var t in Kids(window).Where(x => x.ControlType == ControlType.Text))
                {
                    var m = FetchingRegex.Match(t.Name ?? "");
                    if (m.Success)
                    {
                        progress = m.Groups[1].Value + "/" + m.Groups[2].Value;
                        break;
                    }
                }
            }
            catch
            {
                // status bar being re-created
            }

            if (progress == null)
            {
                _staleProgress = null;
                // nothing is loading (e.g. no Fetch was pressed): two looks 150 ms apart are enough. Right after opening the tab / a global Fetch the status text
                // flickers while the GUI works, so there it must stay away for 1 s in a row.
                if (quick && !announced)
                {
                    if (absentSince == null)
                    {
                        absentSince = sw.ElapsedMilliseconds;
                        Thread.Sleep(150);
                        continue;
                    }

                    return true;
                }

                absentSince ??= sw.ElapsedMilliseconds;
                if (sw.ElapsedMilliseconds - absentSince >= 1000)
                {
                    if (announced)
                    {
                        _log("INFO  | the GUI finished fetching parameters (" + reason + ") after " + sw.Elapsed.TotalSeconds.ToString("0") + " s");
                    }

                    return true;
                }

                Thread.Sleep(500);
                continue;
            }

            absentSince = null;
            if (progress == _staleProgress && FreshSearch(window) is { IsEnabled: true })
            {
                return true;       // the same stuck text as before: already judged, do not wait 8 s again
            }

            announced = true;
            if (progress != lastProgress)
            {
                lastProgress = progress;
                progressSince = sw.ElapsedMilliseconds;
            }
            else if (sw.ElapsedMilliseconds - progressSince >= 8000 && FreshSearch(window) is { IsEnabled: true })
            {
                // same counter for 8 s while the search box is usable = stale status text, not a real load (a real load disables the search)
                _log("WARN  | the status bar is stuck on 'Fetching parameters... " + progress + "' for 8 s but the search box is enabled - treating the load as finished (" + reason + "); not waited for again while it shows the same text");
                _staleProgress = progress;
                return true;
            }

            if (sw.ElapsedMilliseconds - lastLog > 20000)
            {
                _log("INFO  | waiting - the GUI is still fetching parameters: " + progress + " (" + reason + ")");
                lastLog = sw.ElapsedMilliseconds;
            }

            Thread.Sleep(1000);
        }

        return false;
    }

    // The Configure tab disables its search box while it (re)loads the parameters from the unit (tab switch, Fetch, Reset).
    private static bool WaitSearchEnabled(TextBox search, int timeoutMs = 40000)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            try
            {
                if (search.IsEnabled)
                {
                    return true;
                }
            }
            catch
            {
                // element being re-created
            }

            Thread.Sleep(400);
        }

        return false;
    }

    // Focus + clear the search box; the box can be disabled again at any moment while the GUI is loading, so retry.
    private static bool PrepareSearch(TextBox search)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < 40000)
        {
            try
            {
                if (search.IsEnabled)
                {
                    search.Focus();
                    search.Text = "";
                    return true;
                }
            }
            catch (Exception ex) when (ex is FlaUI.Core.Exceptions.ElementNotEnabledException or System.Runtime.InteropServices.COMException or InvalidOperationException)
            {
                // disabled right now - wait and retry
            }

            Thread.Sleep(500);
        }

        return false;
    }

    // The search box element goes stale when the GUI re-creates the list (after Save/Fetch) - SetFocus then throws
    // InvalidOperationException. So it is looked up again every time.
    private static TextBox? FreshSearch(Window window) => window.FindFirstDescendant(cf => cf.ByAutomationId("searchTextBox"))?.AsTextBox();

    // Press the search box's own X ("Clear Search") so the whole list is refreshed, as a user would; fall back to emptying the text.
    private void ClearSearch(Window window)
    {
        try
        {
            var x = FindByHelp(window, "Clear Search");
            if (x != null && x.IsEnabled)
            {
                x.AsButton().Invoke();
                Thread.Sleep(300);
                return;
            }

            var sb = FreshSearch(window);
            if (sb != null && WaitSearchEnabled(sb, 15000))
            {
                sb.Focus();
                sb.Text = "";
            }
        }
        catch (Exception ex)
        {
            _log("NOTE  | could not clear the search box: " + ex.Message);
        }
    }

    private (AutomationElement? name, AutomationElement? value) Locate(Window window, TextBox search, Param p, bool clearFirst = false)
    {
        // restore step only: press X first so the list refreshes, then search the same parameter again (the plain search got stuck there)
        if (clearFirst)
        {
            var swClear = Stopwatch.StartNew();
            ClearSearch(window);
            _log("INFO  | search cleared with X in " + swClear.ElapsedMilliseconds + " ms");
        }

        search = FreshSearch(window) ?? search;
        if (!WaitFetchDone(window, 300000, "before searching " + p.Name))
        {
            Fail(p, "1 locate the parameter", "the GUI kept 'Fetching parameters...' for more than 5 minutes", "is the unit connected and responding? (status bar shows BUS CONFLICT?)");
            return (null, null);
        }

        search = FreshSearch(window) ?? search;
        if (!WaitSearchEnabled(search))
        {
            Fail(p, "1 locate the parameter", "the search box stayed disabled for 40 s - the Configure tab is still loading parameters (or the unit is not connected)", "wait until the Configure tab finished loading, check the unit connection");
            return (null, null);
        }

        if (!PrepareSearch(search))
        {
            Fail(p, "1 locate the parameter", "the search box could not be used (it stayed disabled)", "the Configure tab is probably (re)loading parameters from the unit");
            return (null, null);
        }

        TextCopy.ClipboardService.SetText(p.Name);
        Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_V);

        string wantedScope = p.Channel > 0 ? "Channel " + p.Channel + " Parameters" : "System Parameters";
        AutomationElement? next = null;
        string lastCrumb = "";
        for (int step = 0; step <= 40; step++)
        {
            // wait for the parameter row of the current search result
            AutomationElement? nameEl = null, valueEl = null;
            string crumb = "";
            for (int attempt = 1; attempt <= MaxFindRetries; attempt++)
            {
                nameEl = window.FindFirstDescendant(cf => cf.ByAutomationId(p.NameId));
                valueEl = window.FindFirstDescendant(cf => cf.ByAutomationId(p.ValueId));
                if (nameEl != null && valueEl != null)
                {
                    crumb = Breadcrumb(window);
                    if (crumb.Length > 0 && (step > 0 || crumb != lastCrumb || attempt > 2))
                    {
                        break;
                    }
                }

                Thread.Sleep(RetryDelayMs);
            }

            if (nameEl == null || valueEl == null)
            {
                return (null, null);
            }

            bool inScope = p.Channel > 0
                ? crumb.Contains(wantedScope, StringComparison.Ordinal)
                : crumb.Contains("System Parameters", StringComparison.Ordinal) && !crumb.Contains("Channel", StringComparison.Ordinal);
            if (inScope)
            {
                return (nameEl, valueEl);
            }

            // results are ordered by channel: jump straight to the wanted one instead of stepping 1 by 1
            int presses = 1;
            var cm = CrumbChannelRegex.Match(crumb);
            if (p.Channel > 0 && cm.Success && int.Parse(cm.Groups[1].Value) < p.Channel)
            {
                presses = p.Channel - int.Parse(cm.Groups[1].Value);
            }
            else if (step >= 20)
            {
                _log("FAIL  | " + p.Label + " | no search result in scope '" + wantedScope + "' (breadcrumb of last result: " + crumb + ")");
                return (null, null);
            }

            next ??= FindByHelp(window, "Next Result");
            if (next == null)
            {
                return (null, null);
            }

            lastCrumb = crumb;
            for (int i = 0; i < presses; i++)
            {
                next.AsButton().Invoke();
                Thread.Sleep(350);
            }

            Thread.Sleep(400);
        }

        return (null, null);
    }

    private static string Breadcrumb(Window window)
    {
        var crumb = window.FindFirstDescendant(cf => cf.ByAutomationId("BreadCrumb"));
        if (crumb == null)
        {
            return "";
        }

        return string.Join("", Walk(crumb).Where(t => t.ControlType == ControlType.Text).Select(t => t.Name ?? ""));
    }

    private static (int index, int total)? Counter(Window window)
    {
        foreach (var t in Walk(window).Where(x => x.ControlType == ControlType.Text))
        {
            var m = MatchesRegex.Match(t.Name ?? "");
            if (m.Success)
            {
                return (int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value));
            }
        }

        return null;
    }

    // The toolbar buttons expose NO AutomationId in the running app - they are found by their help text (tooltip).
    private static AutomationElement? FindByHelp(Window window, string helpText)
        => window.FindFirstDescendant(cf => cf.ByHelpText(helpText));

    // raw-view walk (template parts included), as used for the capture
    private static IEnumerable<AutomationElement> Kids(AutomationElement e)
    {
        var walker = e.Automation.TreeWalkerFactory.GetRawViewWalker();
        for (var c = walker.GetFirstChild(e); c != null; c = walker.GetNextSibling(c))
        {
            yield return c;
        }
    }

    private static IEnumerable<AutomationElement> Walk(AutomationElement e, int depth = 0)
    {
        yield return e;
        if (depth > 12)
        {
            yield break;
        }

        foreach (var c in Kids(e))
        {
            foreach (var x in Walk(c, depth + 1))
            {
                yield return x;
            }
        }
    }

    // One read attempt. A ComboBox may answer ValuePattern with "" while the real selection is only in SelectedItem or its
    // child text, so every method is tried.
    private static string ReadOnce(AutomationElement element)
    {
        string result = "";
        try
        {
            // text boxes must be focused to expose their value; a ComboBox is read as it is (a click into it can blank its text)
            if (element.ControlType != ControlType.ComboBox)
            {
                element.Focus();
                Thread.Sleep(250);
            }
            if (element.Patterns.Value.TryGetPattern(out var vp))
            {
                result = (vp.Value.Value ?? "").Trim();
            }

            if (result.Length == 0 && element.ControlType == ControlType.ComboBox)
            {
                var selected = element.AsComboBox().SelectedItem;
                result = (selected?.Name ?? "").Trim();
                if (result.Length == 0)
                {
                    result = (selected?.Text ?? "").Trim();
                }

                if (result.Length == 0)
                {
                    result = Walk(element).Where(t => t.ControlType == ControlType.Text).Select(t => (t.Name ?? "").Trim()).FirstOrDefault(t => t.Length > 0) ?? "";
                }
            }

            if (result.Length == 0 && element.Patterns.Toggle.TryGetPattern(out var tp))
            {
                result = tp.ToggleState.Value == ToggleState.On ? "On" : "Off";
            }
        }
        catch
        {
            // the element may be re-created while the list refreshes - the caller retries with a fresh element
        }

        return result;
    }

    // Right after a search the row can still be the PREVIOUS parameter's recycled container (its binding not yet updated),
    // so the value element is looked up again by its AutomationId on every attempt instead of re-reading a stale element.
    private string ReadParam(Window window, Param p, out AutomationElement? element, int attempts = 15)
    {
        element = null;
        for (int attempt = 1; attempt <= attempts; attempt++)
        {
            var el = window.FindFirstDescendant(cf => cf.ByAutomationId(p.ValueId));
            if (el != null)
            {
                element = el;
                string value = ReadOnce(el);
                if (value.Length > 0)
                {
                    return value;
                }

                // A text editor exposes an EMPTY value to UIA; the user reads the value from a separate text of the same row. Use it right
                // away (waiting for retries does not help and costs ~10 s per read).
                if (el.ControlType == ControlType.Edit)
                {
                    string shown = DisplayedValue(window, p, el);
                    if (shown.Length > 0)
                    {
                        return shown;
                    }
                }
            }

            Thread.Sleep(RetryDelayMs);
        }

        return "";
    }

    // the row also shows the parameter's data type as a plain text ("String") - that is a label, not the value
    private static readonly HashSet<string> DataTypeTokens = new(StringComparer.OrdinalIgnoreCase) { "String", "Integer", "Float", "Number", "Boolean" };

    // The visible text of the row that is the value: not the parameter name, not the data-type label, not a [unit], not a counter.
    private static string DisplayedValue(Window window, Param p, AutomationElement valueEl)
    {
        try
        {
            string paramName = (window.FindFirstDescendant(cf => cf.ByAutomationId(p.NameId))?.Name ?? p.Name).Trim();
            var row = valueEl.Parent;
            return row == null ? "" : VisibleTexts(row, 0)
                .FirstOrDefault(t => t.Length > 0 && t.Length < 80 && t != paramName && !DataTypeTokens.Contains(t) && !t.StartsWith("[") && !Regex.IsMatch(t, @"^\d+ / \d+$")) ?? "";
        }
        catch
        {
            return "";
        }
    }

    private static IEnumerable<string> VisibleTexts(AutomationElement root, int depth)
    {
        foreach (var c in Kids(root))
        {
            if (c.ControlType == ControlType.Text && !c.Properties.IsOffscreen.ValueOrDefault && c.BoundingRectangle.Width > 0)
            {
                yield return (c.Name ?? "").Trim();
            }

            if (depth < 4)
            {
                foreach (var t in VisibleTexts(c, depth + 1))
                {
                    yield return t;
                }
            }
        }
    }

    // The parameter row has its own small refresh button (tooltip "Refresh parameter value from unit") next to the value.
    private bool RefreshParam(AutomationElement valueEl)
    {
        try
        {
            var scope = valueEl.Parent;
            for (int up = 0; up < 5 && scope != null; up++, scope = scope.Parent)
            {
                var button = Walk(scope).FirstOrDefault(b => b.ControlType == ControlType.Button && b.HelpText == "Refresh parameter value from unit");
                if (button != null && button.IsEnabled)
                {
                    button.AsButton().Invoke();
                    return true;
                }
            }
        }
        catch
        {
            // best effort
        }

        return false;
    }

    // ComboBox items expose a raw object ("[250, 3]"); the shown text is in a child Text.
    private static List<string> ReadOptions(AutomationElement comboElement)
    {
        var options = new List<string>();
        var combo = comboElement.AsComboBox();
        combo.Expand();
        Thread.Sleep(500);
        foreach (var item in combo.Items)
        {
            string text = (item.Text ?? "").Trim();
            if (text.Length == 0)
            {
                var child = item.FindFirstDescendant(cf => cf.ByControlType(ControlType.Text));
                text = (child?.Name ?? item.Name ?? "").Trim();
            }

            if (text.Length > 0)
            {
                options.Add(Normalize(text));
            }
        }

        combo.Collapse();
        Thread.Sleep(250);
        return options.Distinct().ToList();
    }

    private void SetValue(Window window, AutomationElement element, string newValue)
    {
        if (element.ControlType == ControlType.ComboBox)
        {
            var combo = element.AsComboBox();
            combo.Expand();
            Thread.Sleep(400);
            ComboBoxItem? target = null;
            foreach (var item in combo.Items)
            {
                string text = (item.Text ?? "").Trim();
                if (text.Length == 0)
                {
                    text = item.FindFirstDescendant(cf => cf.ByControlType(ControlType.Text))?.Name ?? item.Name ?? "";
                }

                if (SameValue(text, newValue))
                {
                    target = item;
                    break;
                }
            }

            if (target == null)
            {
                combo.Collapse();
                throw new InvalidOperationException("option '" + newValue + "' not found in the list");
            }

            target.Select();
            Thread.Sleep(500);
            return;
        }

        if (element.ControlType == ControlType.CheckBox)
        {
            bool wantOn = newValue.Equals("On", StringComparison.OrdinalIgnoreCase);
            var toggle = element.Patterns.Toggle.Pattern;
            if ((toggle.ToggleState.Value == ToggleState.On) != wantOn)
            {
                toggle.Toggle();
            }

            Thread.Sleep(400);
            return;
        }

        // text box: paste the whole value, then commit with Enter. The keystrokes go to the FRONT window - it must be the GUI.
        if (!AppConnection.EnsureGuiInFront(window))
        {
            throw new InvalidOperationException("Power Rider Studio is not the front window (the keystrokes would go to another window) and could not be brought forward");
        }

        element.Focus();
        Thread.Sleep(200);
        Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
        TextCopy.ClipboardService.SetText(newValue);
        Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_V);
        Thread.Sleep(200);
        Keyboard.Type(VirtualKeyShort.ENTER);       // commit the edit like a user does (Enter), then leave the field
        Thread.Sleep(300);
        Keyboard.Type(VirtualKeyShort.TAB);
        Thread.Sleep(500);
    }

    private static bool IsReadOnly(AutomationElement element)
    {
        try
        {
            return element.Patterns.Value.TryGetPattern(out var vp) && vp.IsReadOnly.Value;
        }
        catch
        {
            return false;
        }
    }

    // ------------------------------------------------------------------ tooltip

    private static readonly Regex ConstraintsRegex = new(@"Min:\s*(-?[\d.,]+)[^\w-]+Max:\s*(-?[\d.,]+)", RegexOptions.Compiled);
    private static readonly Regex DefaultRegex = new(@"Default:\s*(.+)$", RegexOptions.Compiled);

    // Reads "Constraints: Min: X • Max: Y • Default: Z" and the unit/data-type label straight from the row's UIA tree - no hovering.
    // The separator is the bullet character (U+2022), not a space. Found=false when the row has no Constraints text (e.g. the three % parameters).
    private static Tip ReadConstraintsFromRow(AutomationElement valueEl)
    {
        var tip = new Tip();
        try
        {
            var row = valueEl.Parent;
            for (int up = 0; up < 2 && row != null && tip.Text.Length == 0; up++, row = row.Parent)      // the row (ParamControlView) and its list item - never the whole list
            {
                foreach (var t in Walk(row).Where(x => x.ControlType == ControlType.Text))
                {
                    string text = (t.Name ?? "").Trim();
                    if (text.Contains("Constraints:", StringComparison.Ordinal))
                    {
                        tip.Text = text;
                        var m = ConstraintsRegex.Match(text);
                        if (m.Success && TryNum(m.Groups[1].Value.Replace(",", ""), out double mn) && TryNum(m.Groups[2].Value.Replace(",", ""), out double mx))
                        {
                            tip.Min = mn;
                            tip.Max = mx;
                        }

                        var d = DefaultRegex.Match(text);
                        if (d.Success)
                        {
                            tip.DefaultText = d.Groups[1].Value.Trim();
                        }

                        break;
                    }
                }

                if (tip.Text.Length > 0)
                {
                    // the visible unit ("[msec]", "[%]") or data type ("String") label of the same row
                    tip.TypeToken = Walk(row).Where(x => x.ControlType == ControlType.Text && !x.Properties.IsOffscreen.ValueOrDefault)
                        .Select(x => (x.Name ?? "").Trim())
                        .FirstOrDefault(x => Regex.IsMatch(x, @"^\[.+\]$") || x.Equals("String", StringComparison.OrdinalIgnoreCase)) ?? "";
                    tip.Found = true;
                }
            }
        }
        catch
        {
            // fall back to hovering
        }

        return tip;
    }

    // Hover the parameter name; the tooltip carries "Constraints: Min: .. Max: .. Default: .." plus MIN/MAX/DEFAULT cells
    // and a type token ("String", ...).
    private Tip ReadTooltip(Window window, AutomationElement nameEl)
    {
        var tip = new Tip();
        try
        {
            var r = nameEl.BoundingRectangle;
            if (r.IsEmpty)
            {
                return tip;
            }

            Mouse.MoveTo(new System.Drawing.Point((int)(r.X + r.Width / 2), (int)(r.Y + r.Height / 2)));
            int pid = window.Properties.ProcessId.ValueOrDefault;
            Thread.Sleep(1200);     // tooltip initial show delay
            for (int i = 0; i < 10 && !tip.Found; i++)
            {
                Thread.Sleep(300);
                foreach (var popup in Kids(window).Concat(Kids(window.Automation.GetDesktop())))
                {
                    if (popup.ClassName != "Popup")
                    {
                        continue;
                    }

                    var tt = Walk(popup).FirstOrDefault(x => x.ControlType == ControlType.ToolTip);
                    if (tt == null)
                    {
                        continue;
                    }

                    var texts = Walk(tt).Where(x => x.ControlType == ControlType.Text).Select(t => (t.Name ?? "").Trim()).Where(t => t.Length > 0).ToList();
                    tip.Text = string.Join(" || ", texts).Replace("\r", " ").Replace("\n", " ");
                    tip.Found = texts.Count > 0;
                    for (int k = 0; k + 1 < texts.Count; k++)
                    {
                        if (texts[k] == "MIN" && TryNum(texts[k + 1], out double mn)) tip.Min = mn;
                        else if (texts[k] == "MAX" && TryNum(texts[k + 1], out double mx)) tip.Max = mx;
                        else if (texts[k] == "DEFAULT") tip.DefaultText = texts[k + 1];
                    }

                    var m = Regex.Match(tip.Text, @"Min:\s*(-?[\d.]+)\s+Max:\s*(-?[\d.]+)");
                    if (m.Success)
                    {
                        tip.Min ??= double.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
                        tip.Max ??= double.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
                    }

                    tip.TypeToken = texts[^1];     // unit ("[V]") or data type ("String")
                    break;
                }
            }

            Mouse.MoveTo(new System.Drawing.Point(5, 5));
            Thread.Sleep(200);
        }
        catch
        {
            // tooltip is best-effort
        }

        return tip;
    }

    // ------------------------------------------------------------------ Save / (Reset) / Fetch

    private bool SaveResetFetch(Window window, TextBox search, Param p)
    {
        // "Save to Unit Flash" stays disabled until an edit is pending - if it never enables, the edit was not registered
        if (!Press(window, "Save to Unit Flash", "Save", SaveTimeoutMs, p)) return false;
        if (DoReset && !Press(window, "Reset Unit", "Reset", ResetTimeoutMs, p)) return false;
        // Read the value back from the unit with the ROW's own refresh button (~0.5 s) instead of the global Fetch Parameters (3-4 s, reloads every
        // parameter). The global Fetch stays as the fallback when the row's button cannot be found, and after a failure (it also drops pending edits).
        if (!RefreshRowFromUnit(window, p) && !Press(window, "Refresh parameters from unit", "Fetch", FetchTimeoutMs, p)) return false;
        return true;
    }

    // "Refresh parameter value from unit" of the parameter's row (AutomationId = name id + "Refresh"). False = button not available.
    private bool RefreshRowFromUnit(Window window, Param p)
    {
        var list = window.FindFirstDescendant(cf => cf.ByAutomationId("ParametersList"));
        var button = list?.FindFirstDescendant(cf => cf.ByAutomationId(p.NameId + "Refresh"));
        if (button == null || !button.IsEnabled)
        {
            return false;
        }

        string before = ReadAppLog();
        _log("CLICK | " + p.Label + " | Refresh this parameter from the unit");
        button.AsButton().Invoke();
        // the row refresh reads one value (~0.5 s) and not every GUI version writes a log line for it: wait for the log only briefly
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < 1500)
        {
            Thread.Sleep(150);
            if (ReadAppLog() != before)
            {
                string added = Diff(before, ReadAppLog());
                if (added.Length > 0)
                {
                    _log("APPLOG| " + added.Replace("\r", "").Replace("\n", " ## "));
                }

                break;
            }
        }

        Thread.Sleep(300);
        return true;
    }

    private bool Press(Window window, string helpText, string label, int timeoutMs, Param p)
    {
        // the unit may still be restarting after a Reset: wait until the button exists and is enabled
        AutomationElement? button = null;
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            button = FindByHelp(window, helpText);
            if (button != null && button.IsEnabled)
            {
                break;
            }

            button = null;
            Thread.Sleep(500);
        }

        if (button == null)
        {
            Fail(p, "4 " + label, "the '" + label + "' button (" + helpText + ") did not become available/enabled within " + timeoutMs / 1000 + " s",
                 label == "Save" ? "Save is disabled while no change is pending - the previous step probably did not register the edit"
                 : label == "Reset" ? "the unit/GUI may be busy or disconnected"
                 : "the unit may still be restarting after Reset (reconnect time)");
            return false;
        }

        string before = ReadAppLog();
        _log("CLICK | " + p.Label + " | " + label);
        button.AsButton().Invoke();

        // wait for the app log to show something new, then let it settle
        sw.Restart();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            Thread.Sleep(200);
            string now = ReadAppLog();
            if (now != before)
            {
                Thread.Sleep(300);
                string added = Diff(before, ReadAppLog());
                if (added.Length > 0)
                {
                    _log("APPLOG| " + added.Replace("\r", "").Replace("\n", " ## "));
                    if (Regex.IsMatch(added, @"Failed|Error|Timeout", RegexOptions.IgnoreCase))
                    {
                        // known GUI bug: "Failed to save..." appears although the value is stored - NOT a failure here, the read-back after Fetch decides
                        if (!_guiSaveNoteShown)
                        {
                            _guiSaveNoteShown = true;
                            _log("NOTE  | the GUI log says 'Failed to ... after " + label + "' although the value is saved (known GUI bug). Not a failure of this test: the value read back from the unit decides. Not repeated.");
                        }
                    }
                }

                if (HasPasswordDialog(window))
                {
                    Fail(p, "4 " + label, "the app asks for the configuration-mode password", "the password is not automated - enter configuration mode manually first");
                    return false;
                }

                return true;
            }
        }

        // no log line appeared - the action may simply not log; continue after a pause but say so
        _log("NOTE  | " + p.Label + " | no new app-log line after '" + label + "' within " + timeoutMs / 1000 + " s - continuing");
        Thread.Sleep(2000);
        if (HasPasswordDialog(window))
        {
            Fail(p, "4 " + label, "the app asks for the configuration-mode password", "the password is not automated - enter configuration mode manually first");
            return false;
        }

        return true;
    }

    private string ReadAppLog()
    {
        try
        {
            var log = _logElement;
            return log == null ? "" : string.Join("\n", Walk(log).Where(t => t.ControlType == ControlType.Text).Select(t => t.Name ?? ""));
        }
        catch
        {
            return "";
        }
    }

    private static string Diff(string before, string after)
    {
        var old = before.Split('\n').ToHashSet();
        return string.Join("\n", after.Split('\n').Where(l => l.Length > 0 && !old.Contains(l)));
    }

    private static bool HasPasswordDialog(Window window)
    {
        return window.FindFirstDescendant(cf => cf.ByAutomationId("ConfigurationModePasswordOkButton")) != null
               || window.FindFirstDescendant(cf => cf.ByAutomationId("PasswordBox")) != null;
    }

    // ------------------------------------------------------------------ compare / format

    private static string Normalize(string s)
    {
        s = (s ?? "").Trim();
        var m = BracketedRegex.Match(s);
        return m.Success ? m.Groups[1].Value.Trim() : s;
    }

    private static bool SameValue(string a, string b)
    {
        a = Normalize(a);
        b = Normalize(b);
        if (double.TryParse(a, NumberStyles.Float, CultureInfo.InvariantCulture, out double x)
            && double.TryParse(b, NumberStyles.Float, CultureInfo.InvariantCulture, out double y))
        {
            return Math.Abs(x - y) < 1e-6;
        }

        return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
    }

    // Test value per unit for parameters whose tooltip has no Min/Max. EDIT HERE to add units; unknown units are skipped.
    private static double? UnitTestValue(string unitToken)
    {
        switch (unitToken.Trim().Trim('[', ']').ToLowerInvariant())
        {
            case "msec": case "ms": return 10;
            case "sec": case "s": return 1;
            default: return null;
        }
    }

    private static bool TryNum(string s, out double d) => double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d);
    private static string Fmt(double? d) => d.HasValue ? d.Value.ToString("0.###", CultureInfo.InvariantCulture) : "?";

    private static string ReadGuiVersion()
    {
        try
        {
            var p = Process.GetProcessesByName("PowerRiderStudio").FirstOrDefault();
            var info = p?.MainModule?.FileVersionInfo;
            return info == null ? "" : (info.ProductVersion ?? info.FileVersion ?? "");
        }
        catch
        {
            return "";
        }
    }

    // ------------------------------------------------------------------ Excel: parameters

    private List<Param> ReadParams()
    {
        string temp = CopyToTemp();
        try
        {
            using var wb = new XLWorkbook(temp);
            var result = new List<Param>();
            foreach (var name in Sheets)
            {
                if (!wb.Worksheets.Contains(name))
                {
                    continue;
                }

                var ws = wb.Worksheet(name);
                foreach (var row in ws.RowsUsed().Skip(1))
                {
                    string paramName = row.Cell(3).GetString().Trim();
                    string nameId = ExtractId(row.Cell(2).GetString());
                    string valueId = ExtractId(row.Cell(4).GetString());
                    if (paramName.Length == 0 || nameId.Length == 0 || valueId.Length == 0)
                    {
                        continue;
                    }

                    // column A: "System" or the channel number 1..16
                    int.TryParse(row.Cell(1).GetString().Trim(), out int channel);
                    result.Add(new Param
                    {
                        Sheet = name,
                        ExcelRow = row.RowNumber(),
                        Name = paramName,
                        NameId = nameId,
                        ValueId = valueId,
                        Default = row.Cell(5).GetString().Trim(),
                        Channel = channel,
                    });
                }
            }

            return result;
        }
        finally
        {
            try { File.Delete(temp); } catch { /* temp copy */ }
        }
    }

    // The id columns of the workbook hold the plain AutomationId - used exactly as written.
    private static string ExtractId(string cell) => (cell ?? "").Trim();

    private string CopyToTemp()
    {
        string temp = Path.Combine(Path.GetTempPath(), "ConfigRandom_" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.Copy(_excelPath, temp, true);       // works even while the workbook is open in Excel
        return temp;
    }
}
