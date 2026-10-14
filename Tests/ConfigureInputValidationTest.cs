using System.Diagnostics;
using System.Threading.Tasks;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using ClosedXML.Excel;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;

namespace PowerGUIAutomation.Tests;

// INPUT VALIDATION of the Configure tab: types values that must be rejected (and a few that must be accepted) into the text fields of the
// parameters and checks how the GUI reacts. Parameters are found like in ConfigureChangeAndRestoreAllParametersTest (tree page scan).
//
// HOW THE GUI REACTS (found on the live GUI): pressing Enter in a field applies the value to the unit at once (no Save needed). If the value
// is refused, a red X appears to the right of the field together with a small chip that shows the last value the unit accepted. The X is not
// a UIA element, but the chip is: a visible text to the right of the field. That chip is what this test looks for. Save to Unit Flash is NOT
// used as a signal (it can be enabled for other reasons) and this test NEVER presses Save: nothing is written to flash.
//
// Per parameter (text fields only; ComboBoxes have nothing to type):
//   number with Min/Max : Max+1 and Min-1 must be rejected, "abc" must be rejected, empty must be rejected, Max and Min must be accepted
//                         (when the constraint depends on another parameter: "must be lower than" -> only Min; "must be higher than" -> Max as is, and Min
//                         while the other parameter is temporarily set to its own Min and put back afterwards)
//   number, no range    : "abc" and empty must be rejected
//   text                : Hebrew letters must be rejected, empty must be rejected (Min >= 1), Max+1 characters must be rejected or cut to
//                         Max, a normal text with English letters, digits, space and symbols must be accepted
// After every input the original value is typed back, so the unit ends exactly as it was (RAM and flash).
//
// Scope: every System and channel parameter of the unit (all rows, or the rows / ranges / failures file given at the start, like test 2). Never touched: baud rates, addresses,
// Logic Control / Logic Expression, Group Control.
//
// Like test 2, the run asks at the start: stop at the first failing input (default) or continue and list every failure at the end (log + a
// workbook in the Failures folder next to the logs), and which Excel rows to run (Enter = all, or "54-70, 72, 100-").
public class ConfigureInputValidationTest
{
    private const int RejectWaitMs = 1800;        // how long to wait for the red-X chip before deciding that the value was accepted
    private const string HebrewSample = "שלום";
    private static readonly string TextSample = Environment.GetEnvironmentVariable("VALIDATION_TEXT") is { Length: > 0 } t ? t : "Ab 1-_";      // env var: developer switch

    private static readonly HashSet<string> NeverChange = new(StringComparer.OrdinalIgnoreCase)
    {
        "CAN Bus Baud Rate", "Serial Baud Rate", "Destination Address", "Unit ID Address",
    };

    private readonly string _excelPath;
    private readonly Action<string> _log;
    private readonly UnitProfile _unit;

    private sealed record Row(string Sheet, int ExcelRow, int Channel, string Name, string NameId, string ValueId, string Default)
    {
        public int Scope => ConfigureNav.ScopeOf(Channel, Name);
        public string Label => (Channel > 0 ? "Channel " + Channel + " | " : "System | ") + Name + " (" + Sheet + " row " + ExcelRow + ")";
    }

    private sealed record Finding(Row Row, string Input, string Expected, string Actual, string GuiLog, string Problem, string Original = "", string Typed = "", string Shown = "")
    {
        public int Round { get; init; }                      // the loop round (0 = not a loop)
        public DateTime At { get; init; } = DateTime.Now;    // when the failure happened (the "Run time" column)
    }

    // LOOP (like test 2): the whole run is repeated `_rounds` times, or round after round until the stop key (_unlimited), with an optional time limit. A failure that
    // shows up only after many writes to the unit means the quantity is the problem and not the parameter.
    private int _rounds = 1, _round;
    private bool _unlimited;
    private TimeSpan? _timeLimit;
    private bool Looping => _unlimited || _rounds > 1;
    private string LoopText => !Looping ? "no loop (one pass)" : (_unlimited ? "round after round until stopped" : _rounds + " rounds") + (_timeLimit.HasValue ? ", time limit " + RunReport.Format(_timeLimit.Value) : "")
                                                                  + "; rounds started: " + Math.Max(_round, 1);
    private string RoundText => _unlimited ? "round " + _round : "round " + _round + " of " + _rounds;
    private int _writes;                                     // values typed into the unit so far (every typed input and every put-back, Enter applies it)
    private int _firstFailureWrites = -1;                    // writes done when the first failure came
    private string _firstFailureWhere = "";

    private bool _stopOnFail = true;
    private bool _stopRequested;
    private string _guiVersion = "";
    private const int MaxFailuresInARow = 10;      // continue mode: this many failing inputs in a row = something is wrong with the unit/GUI, stop

    // What the test is doing right now (heartbeat / "stuck" report); the run is declared stuck after NoProgressMinutes without a new step.
    private volatile string _where = "";
    private long _progressAt = Environment.TickCount64;
    private const int NoProgressMinutes = 3;

    private void At(string what)
    {
        _where = what;
        _progressAt = Environment.TickCount64;
    }

    private bool RunWatched(Action action)
    {
        At("starting");
        var task = Task.Run(action);
        long lastBeat = Environment.TickCount64;
        while (!task.Wait(1000))
        {
            long now = Environment.TickCount64;
            if (now - _progressAt > NoProgressMinutes * 60_000L)
            {
                return false;
            }

            if (now - lastBeat >= 30_000 && now - _progressAt >= 30_000)
            {
                lastBeat = now;
                _log("INFO  | still working - " + ((now - _progressAt) / 1000) + " s on: " + _where);
            }
        }

        task.GetAwaiter().GetResult();
        return true;
    }

    private readonly List<Finding> _findings = new();

    // The failures workbook of THIS run (fixed name, rewritten after every failing input and at the end - see RunReport) and what it needs to say
    // about a run that did not reach its end: the rows of the run, how many are finished, and the one being checked right now.
    private const string FilePrefix = "InputValidation";
    private RunReport? _report;
    private List<Row> _todo = new();
    private List<Row> _allRows = new();                 // every row of the sheet (the other parameter of a "must be higher than ..." constraint may not be in the run)
    private volatile string _tempChange = "";           // a parameter that is temporarily changed right now (the Min check of a "higher than" parameter); empty = none
    private int _finished;
    private Row? _current;
    private string _rowsDescription = "all";
    private string? _endStatus;

    // Recorded under the report's lock: ValidateOne runs on a worker thread and the Ctrl+C / console-close handler writes the workbook from another one.
    private void AddFinding(Finding f)
    {
        f = f with { Round = Looping ? _round : 0 };
        if (Looping)
        {
            f = f with { Problem = RoundText[0].ToString().ToUpperInvariant() + RoundText.Substring(1) + ", after " + _writes + " writes to the unit: " + f.Problem };
        }

        if (_firstFailureWrites < 0)
        {
            _firstFailureWrites = _writes;
            _firstFailureWhere = f.Row.Label + (Looping ? ", " + RoundText : "");
        }

        if (_report == null)
        {
            _findings.Add(f);
            return;
        }

        lock (_report.Sync)
        {
            _findings.Add(f);
        }

        WriteFailuresFile();      // the file is current after every failure, not only at the end of a long run
    }

    // Ctrl+C or closing the console: the run ends here, so the file must say so. The row being checked is listed - a value typed there is applied by
    // Enter (nothing is saved to flash) and may still be on the unit.
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
                _findings.Add(new Finding(_current, "(interrupted)", "-", "the run was stopped while this parameter was being checked", "",
                                          "The test was stopped by user while modifying this parameter, please verify that the value is correct."
                                          + (_tempChange.Length > 0 ? " It was also changing " + _tempChange + " temporarily - verify that one too." : "")) { Round = Looping ? _round : 0 });
            }

            report.Finish(status);
        }

        _log("STOPPED | " + status + (_current != null ? " | interrupted while checking: " + _current.Label + " - check this parameter on the unit" : ""));
        WriteFailuresFile();
        _log("RESULT| run time        : " + RunReport.Format(report.Elapsed));
        _log("RESULT| failures file   : " + report.Path);
    }

    private int _failInARow;
    private int _inputs, _passed, _skippedParams, _params;
    private readonly Dictionary<int, double> _lastPct = new();
    private Dictionary<int, List<int>> _scopeRows = new();           // per scope: the Excel rows in list order
    private AutomationElement? _logElement;
    private Window? _window;

    public ConfigureInputValidationTest(string excelPath, Action<string> log, UnitProfile unit)
    {
        _excelPath = excelPath;
        _log = log;
        _unit = unit;
    }

    // The run is wrapped so that Ctrl+C, closing the console, a crash and a setup failure all leave the failures file with the right status.
    public bool Run()
    {
        ConsoleCancelEventHandler onCancel = (_, _) => Interrupted("Stopped by the user (Ctrl+C)");
        EventHandler onExit = (_, _) => Interrupted("Stopped - the console was closed or the program was ended");
        Console.CancelKeyPress += onCancel;
        AppDomain.CurrentDomain.ProcessExit += onExit;
        try
        {
            bool ok = RunCore();
            if (_report != null && !_report.Finished)
            {
                // left before the summary (the unit is not the right one, the Configure tab did not open, ...): nothing was checked
                _report.Finish("Test setup failed - nothing was checked");
                WriteFailuresFile();
                _log("RESULT| status          : " + _report.Status);
                _log("RESULT| total run time  : " + RunReport.Format(_report.Elapsed));
                _log("RESULT| failures file   : " + _report.Path);
            }

            return ok;
        }
        catch (Exception ex)
        {
            if (_report != null && !_report.Finished)
            {
                // an exception that nothing caught: "test failed" in the middle - the file says so and names the row that was being checked
                if (_current != null)
                {
                    lock (_report.Sync)
                    {
                        _findings.Add(new Finding(_current, "(crash)", "-", ex.GetType().Name + ": " + ex.Message, "", "The test crashed while checking this parameter, please verify that the value is correct.") { Round = Looping ? _round : 0 });
                    }
                }

                _report.Finish("Test failed - crashed: " + ex.GetType().Name + ": " + ex.Message);
                WriteFailuresFile();
                _log("FAIL  | the test crashed: " + ex.GetType().Name + ": " + ex.Message);
                _log("RESULT| total run time  : " + RunReport.Format(_report.Elapsed));
                _log("RESULT| failures file   : " + _report.Path);
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
        string onlyName = Environment.GetEnvironmentVariable("VALIDATION_ONLY") ?? "";      // developer switch: run one parameter only (name contains)

        List<Row> rows;
        try
        {
            rows = ReadRows();
            _allRows = rows;
        }
        catch (Exception ex)
        {
            _log("FAIL  | could not read the sheet '" + _unit.ParameterSheet + "': " + ex.Message);
            return false;
        }

        _scopeRows = rows.GroupBy(r => r.Scope).ToDictionary(g => g.Key, g => g.OrderBy(r => r.ExcelRow).Select(r => r.ExcelRow).ToList());
        _stopOnFail = StartQuestions.AskStopOnFail("input");

        var todo = rows.Where(r => r.Channel <= _unit.Channels
                                   && !NeverChange.Contains(r.Name) && !IsLogicParam(r.Name) && r.Scope != ConfigureNav.GroupScope
                                   && (onlyName.Length == 0 || r.Name.Contains(onlyName, StringComparison.OrdinalIgnoreCase))).ToList();

        var selection = RowSelection.Ask(FilePrefix, _unit.PartNumber, out _rowsDescription, 2);
        StartQuestions.AskLoop(out _rounds, out _unlimited, out _timeLimit);
        if (selection != null)
        {
            todo = todo.Where(r => selection.Contains(r.ExcelRow)).ToList();
        }

        _guiVersion = ReadGuiVersion();
        if (!GuiVersionIsAtLeast(_guiVersion, 2, 1, 2))
        {
            _log("WARN  | GUI '" + _guiVersion + "' is older than 2.1.2: this test was developed on 2.1.2 (a refused value leaves a red X + the stored value next to the field and Enter applies the value). Its results on this GUI are NOT validated.");
        }

        _log("START | Configure input validation | unit=" + _unit.Name + " | scope: System + all " + _unit.Channels + " channels"
             + " | parameters=" + todo.Count + (onlyName.Length > 0 ? " | only names containing '" + onlyName + "'" : "") + " | on failure: " + (_stopOnFail ? "stop" : "continue")
             + " | rows: " + _rowsDescription + " | NOTHING is saved to flash"
             + (Looping ? " | LOOP: " + (_unlimited ? "round after round until you stop it" : _rounds + " rounds") + " of the same rows" : ""));
        if (selection?.SourceFile != null)
        {
            _log("INFO  | RE-RUN of the failed rows listed in: " + selection.SourceFile);
            foreach (var mismatch in selection.NameMismatches(todo.Select(r => (r.ExcelRow, r.Name))))
            {
                _log("WARN  | " + mismatch + " - was the master workbook edited since that run?");
            }
        }

        if (todo.Count == 0)
        {
            _log("FAIL  | no parameter to check");
            return false;
        }

        _todo = todo;
        _report = new RunReport(FilePrefix, _unit.PartNumber);
        RunClock.Start();      // the menu's "took" time starts here: the questions above are not counted
        _log("INFO  | this run's failures file (rewritten after every failure and at the end, also after Ctrl+C or a crash): " + _report.Path);
        WriteFailuresFile();      // exists from the first minute: status "Running" until the run ends (deleted again if nothing fails)

        using var automation = new UIA3Automation();
        var window = AppConnection.Attach(automation);
        _window = window;
        string? unitProblem = UnitCheck.Verify(window, _unit, _log);
        if (unitProblem != null)
        {
            _log("FAIL  | (test setup) | STEP: check the connected unit | WHY: " + unitProblem + " | CHECK: connect '" + _unit.Name + "' or choose the right unit (U in the main menu)");
            return false;
        }

        AppConnection.GoToTab(window, "ConfigureTab");
        if (!ConfigureNav.WaitFetchDone(window, 300000, _log))
        {
            _log("FAIL  | (test setup) | the GUI kept 'Fetching parameters...' for more than 5 minutes");
            return false;
        }

        _logElement = window.FindFirstDescendant(cf => cf.ByAutomationId("ConfigureLog"));
        var watch = Stopwatch.StartNew();
        int index = 0;
        for (int round = 1; _unlimited || round <= _rounds; round++)
        {
            _round = round;
            index = 0;
            if (Looping)
            {
                _log("ROUND | LOOP ROUND " + round + (_unlimited ? "" : " OF " + _rounds) + " STARTED | " + todo.Count + " parameters | failed parameters so far: " + _findings.Select(f => f.Row.ExcelRow).Distinct().Count()
                     + " | writes to the unit so far: " + _writes + " | running for " + RunReport.Format(watch.Elapsed) + (_timeLimit.HasValue ? " (time limit " + RunReport.Format(_timeLimit.Value) + ")" : ""));
            }

            foreach (var scope in todo.GroupBy(r => r.Scope))
            {
                string? crumb = ConfigureNav.SelectScope(window, scope.Key);
                if (crumb == null)
                {
                    _log("FAIL  | the tree page of scope " + scope.Key + " could not be selected");
                    return false;
                }

                foreach (var row in scope)
                {
                    if (StopControl.Requested)
                    {
                        _endStatus = "Stopped by the user (stop key) after finishing the last parameter";
                        _log("STOPPED | the stop key was pressed - the run stops here, the last parameter was finished and its original value is back");
                        _log("RESUME| to continue later answer 'Excel rows to run' with:  " + row.ExcelRow + "-  (from this parameter to the end)");
                        goto Summary;
                    }

                    if (_timeLimit.HasValue && watch.Elapsed >= _timeLimit.Value)
                    {
                        // the loop's time limit: the previous parameter is finished and its original value is back, nothing is half done
                        _endStatus = "Completed - the time limit of " + RunReport.Format(_timeLimit.Value) + " was reached";
                        _log("STOPPED | the time limit of " + RunReport.Format(_timeLimit.Value) + " was reached - the run stops here, the last parameter was finished and its original value is back");
                        goto Summary;
                    }

                    index++;
                    _current = row;
                    _log("DRAW  | " + (Looping ? RoundText + " | " : "") + index + "/" + todo.Count + " | " + row.Label);
                    bool finished;
                    try
                    {
                        finished = RunWatched(() => ValidateOne(window, row));
                    }
                    catch (Exception ex)
                    {
                        _log("WARN  | " + row.Label + " | interrupted (" + ex.GetType().Name + ": " + ex.Message + ") - the value is put back by a Fetch");
                        FetchDiscard(window);
                        finished = true;
                    }

                    if (!finished)
                    {
                        AddFinding(new Finding(row, "(stuck)", "no hang", "no progress for " + NoProgressMinutes + " min - last action: " + _where, "", "The test got stuck (no progress for " + NoProgressMinutes + " minutes, last action: " + _where + "), please verify that the value is correct."));
                        _log("FAIL  | " + row.Label + " | STEP: " + _where + " | WHY: no progress for " + NoProgressMinutes + " minutes | CHECK: the GUI may be frozen or a UI call is not returning; check the unit and the GUI");
                        _stopRequested = true;
                    }

                    _current = null;
                    _finished = index;
                    if (_stopRequested)
                    {
                        _endStatus = !finished ? "Stopped - stuck (no progress for " + NoProgressMinutes + " minutes)"
                                   : _stopOnFail ? "Stopped after the first failing input"
                                   : "Stopped - " + _failInARow + " failing inputs in a row (unit or GUI problem?)";
                        _log("STOPPED | " + (finished ? "halted after the first failing input" : "the run is stuck and cannot continue safely") + " - remaining parameters were not checked");
                        _log("RESUME| to continue later answer 'Excel rows to run' with:  " + row.ExcelRow + "-  (this parameter again)   or give the failures file (only its failed rows)");
                        goto Summary;
                    }
                }
            }

            if (Looping)
            {
                _log("ROUND | LOOP ROUND " + round + (_unlimited ? "" : " OF " + _rounds) + " FINISHED | parameters checked=" + _params + " inputs checked=" + _inputs + " passed=" + _passed + " failures=" + _findings.Count
                     + " | writes to the unit so far: " + _writes + " | " + RunReport.Format(watch.Elapsed));
            }
        }

        Summary:

        _report!.Finish(_endStatus ?? (_findings.Count == 0 ? "Completed - no failures" : "Completed - with failures"));
        _log("DONE  | parameters=" + _params + " skipped=" + _skippedParams + " | inputs checked=" + _inputs + " passed=" + _passed + " findings=" + _findings.Count + " | " + RunReport.Format(watch.Elapsed) + " | scroll positions tried: " + ConfigureNav.ScrollSteps);
        _log("RESULT| " + new string('=', 90));
        _log("RESULT| " + (_findings.Count == 0 ? "PASSED" : "FAILED") + " | unit=" + _unit.Name + " | inputs checked=" + _inputs + " | failures=" + _findings.Count + " | mode=" + (_stopOnFail ? "stop on fail" : "continue"));
        _log("RESULT| status          : " + _report.Status);
        if (Looping)
        {
            _log("RESULT| loop            : " + _round + (_unlimited ? "" : " of " + _rounds) + " rounds started, " + _writes + " writes to the unit in total; "
                 + (_firstFailureWrites >= 0 ? "the FIRST failure came after " + _firstFailureWrites + " writes (" + _firstFailureWhere + ")" : "no failure"));
        }

        _log("RESULT| total run time  : " + RunReport.Format(_report.Elapsed) + "  (" + _report.Started.ToString("HH:mm:ss") + " - " + (_report.Ended ?? DateTime.Now).ToString("HH:mm:ss") + ")");
        foreach (var g in _findings.GroupBy(x => x.Row.ExcelRow))
        {
            _log("RESULT|   row " + g.Key + " | " + g.First().Row.Label.Split('(')[0].Trim() + " -> " + string.Join("; ", g.Select(x => x.Problem)));
        }

        string? file = WriteFailuresFile();
        if (file != null)
        {
            _log("RESULT| failures file : " + file + (_findings.Count > 0 ? "   (to re-run only these rows give this file at 'Excel rows to run')" : ""));
        }

        _log("RESULT| " + new string('=', 90));
        return _findings.Count == 0;
    }

    // ------------------------------------------------------------------ one parameter

    private enum Expect { Reject, Accept, RejectOrCut }

    private void ValidateOne(Window window, Row row)
    {
        var tb = FindValueElement(window, row);
        if (tb == null)
        {
            _skippedParams++;
            _log("SKIP  | " + row.Label + " | not found on the page");
            return;
        }

        if (tb.ControlType != ControlType.Edit)
        {
            _skippedParams++;
            _log("SKIP  | " + row.Label + " | " + tb.ControlType + " - nothing to type");
            return;
        }

        string original = DisplayedValue(window, row, tb);
        var tip = ReadConstraints(tb);
        _log("READ  | " + row.Label + " | current='" + original + "' | constraints: Min=" + Fmt(tip.Min) + " Max=" + Fmt(tip.Max) + " unit/type=" + tip.TypeToken
             + (tip.Direction.Length > 0 ? " | must be " + tip.Direction + " than another parameter (" + (tip.Direction == "lower" ? "only the Min limit is typed as a valid value" : "Max is typed as is; Min is typed while the other parameter is temporarily at its minimum") + ")" : ""));
        if (original.Length == 0)
        {
            _skippedParams++;
            _log("SKIP  | " + row.Label + " | the current value could not be read");
            return;
        }

        // The unit must be at its defaults before the run (reset): a different value means the unit was not reset, a previous run left a value behind, or the Excel
        // default is wrong. The parameter is still checked (and put back to the value it had), but this is reported - like in the change-and-restore test.
        if (row.Default.Length > 0 && !SameValue(original, row.Default))
        {
            _log("WARN  | " + row.Label + " | before the test the value is '" + original + "' but the Excel default is '" + row.Default + "' (the unit is not at its default)");
            AddFinding(new Finding(row, "(before the test)", "the unit is at its default", "unit shows '" + original + "'", "",
                                   "Before the test the unit shows “" + original + "”, but the Excel default is “" + row.Default + "”: the unit is not at its default, or the Excel default is wrong.",
                                   original, row.Default, original));
            _stopRequested |= _stopOnFail;
        }

        _params++;
        int inputs0 = _inputs, passed0 = _passed, findings0 = _findings.Count;
        bool isString = tip.TypeToken.Equals("String", StringComparison.OrdinalIgnoreCase) || !double.TryParse(original, NumberStyles.Float, CultureInfo.InvariantCulture, out _);
        var inputs = new List<(string text, Expect expect, string what)>();
        if (isString)
        {
            inputs.Add((HebrewSample, Expect.Reject, "Hebrew letters"));
            if (tip.Min is >= 1)
            {
                inputs.Add(("", Expect.Reject, "empty"));
            }

            if (tip.Max is > 0 and <= 128)
            {
                inputs.Add((new string('A', (int)tip.Max.Value + 1), Expect.RejectOrCut, "Max+1 characters (" + ((int)tip.Max.Value + 1) + ")"));
            }

            if (!tip.Max.HasValue || tip.Max.Value >= TextSample.Length)
            {
                inputs.Add((TextSample, Expect.Accept, "English letters, digit, space, symbols"));
            }
        }
        else
        {
            inputs.Add(("abc", Expect.Reject, "letters"));
            inputs.Add(("5abc", Expect.Reject, "digit mixed with letters"));
            inputs.Add(("", Expect.Reject, "empty"));
            if (tip.Min.HasValue && tip.Max.HasValue)
            {
                inputs.Add(((tip.Max.Value + 1).ToString("0.###", CultureInfo.InvariantCulture), Expect.Reject, "Max+1"));
                inputs.Add(((tip.Min.Value - 1).ToString("0.###", CultureInfo.InvariantCulture), Expect.Reject, "Min-1"));
                // The two limits themselves must be ACCEPTED. When the constraint depends on another parameter ("must be lower/higher than ..."), only the
                // limit that is valid next to the other parameters' current values is used: lower -> Min, higher -> Max.
                if (tip.Direction != "lower")
                {
                    inputs.Add((tip.Max.Value.ToString("0.###", CultureInfo.InvariantCulture), Expect.Accept, "Max (the limit itself)"));
                }

                if (tip.Direction != "higher")
                {
                    inputs.Add((tip.Min.Value.ToString("0.###", CultureInfo.InvariantCulture), Expect.Accept, "Min (the limit itself)"));
                }
            }
        }

        foreach (var (text, expect, what) in inputs)
        {
            // a value that must be ACCEPTED is tried a second time before it counts as a failure (a red X left over from the previous input can look like a rejection)
            if (expect != Expect.Accept || !CheckInput(window, row, original, text, expect, what, tip, record: false))
            {
                if (expect == Expect.Accept)
                {
                    _log("NOTE  | " + row.Label + " | " + what + " '" + text + "' was refused - trying once more");
                }

                CheckInput(window, row, original, text, expect, what, tip);
            }

            if (_stopRequested || StopControl.Requested)
            {
                break;       // the original value has been put back by CheckInput; the run stops after this parameter
            }
        }

        if (tip.Direction == "higher" && tip.Min.HasValue && !_stopRequested && !StopControl.Requested)
        {
            CheckMinWithOtherAtItsMinimum(window, row, original, tip);
        }

        // the original is typed back after every input; make sure the unit really shows it again
        ClearRejection(window, row);
        string now = "";
        for (int i = 0; i < 10 && !SameValue(now, original); i++)
        {
            Thread.Sleep(i == 0 ? 0 : 200);
            var again = FindValueElement(window, row);
            now = again == null ? "" : DisplayedValue(window, row, again);
        }
        if (!SameValue(now, original))
        {
            _log("WARN  | " + row.Label + " | after the checks the field shows '" + now + "' instead of '" + original + "' - Fetch to read the unit again");
            FetchDiscard(window);
        }

        // the end of this parameter's block: the next parameter starts after the separator line
        _log("INFO  | " + row.Label + " | all validations of this parameter done: " + (_inputs - inputs0) + " inputs, " + (_passed - passed0) + " passed, " + (_findings.Count - findings0) + " failed");
    }

    // record=false: a failing result is NOT counted/logged (used for the first try of a value that must be ACCEPTED, which is repeated once); returns whether the result was as expected.
    private bool CheckInput(Window window, Row row, string original, string text, Expect expect, string what, Tip tip, bool record = true)
    {
        // a red-X chip left over from the previous input would look like a rejection of this one: start from a clean field
        ClearRejection(window, row);

        string before = ReadAppLog();
        var tb = FindValueElement(window, row)!;
        At("typing '" + (text.Length > 20 ? text.Substring(0, 20) + "..." : text) + "'");
        TypeAndEnter(tb, text);
        bool chip = WaitForReject(window, row, RejectWaitMs, () => NewGuiErrors(before).Length > 0, () => SameText(DisplayedValue(window, row, FindValueElement(window, row) ?? tb), text),
                                  out bool seenX, out bool seenV);
        tb = FindValueElement(window, row) ?? tb;
        string shown = DisplayedValue(window, row, tb);
        if (!chip && SameText(shown, text))
        {
            // The box shows what was typed - that proves nothing: a refused or an ignored input (an empty one) stays in the box, and the green V can be left over
            // from the previous input. Only the UNIT tells whether the value was taken: read the row from the unit again.
            string fromUnit = ReadFromUnit(window, row);
            if (fromUnit != null)
            {
                shown = fromUnit;
                tb = FindValueElement(window, row) ?? tb;
            }
        }

        string guiLog = NewGuiErrors(before);
        bool accepted = !chip && SameText(shown, text);
        string actual = chip ? "REJECTED (red X / last accepted value shown)" : accepted ? "ACCEPTED (value taken)" : "no red X, but the field shows '" + shown + "'";
        bool ok = expect switch
        {
            Expect.Reject => chip || !SameText(shown, text),              // refused by the red X, or silently not taken
            Expect.Accept => accepted,
            _ => chip || !SameText(shown, text) || shown.Length <= (int)(tip.Max ?? int.MaxValue),     // refused, or cut to the maximum length
        };
        string expected = expect switch { Expect.Reject => "REJECTED", Expect.Accept => "ACCEPTED", _ => "REJECTED or cut to Max" };
        string icons = " | icons: red X=" + (seenX ? "YES" : "no") + ", green V=" + (seenV ? "yes" : "no");

        if (!ok && !record)
        {
            tb = FindValueElement(window, row) ?? tb;
            if (chip)
            {
                ClearRejection(window, row);
            }
            else
            {
                TypeAndEnter(tb, original);
                WaitForRejectChipGone(window, row, 3000);
                ClearRejection(window, row);
            }

            return false;
        }

        _inputs++;
        string shownText = text.Length > 20 ? text.Substring(0, 20) + "...(" + text.Length + ")" : text;
        if (ok)
        {
            _passed++;
            _log("PASS  | " + row.Label + " | " + what + " '" + shownText + "' | expected " + expected + " | " + actual + icons + (guiLog.Length > 0 ? " | GUI log: " + guiLog : ""));
        }
        else
        {
            _log("FAIL  | " + row.Label + " | " + what + " '" + shownText + "' | expected " + expected + " | " + actual + icons + (guiLog.Length > 0 ? " | GUI log: " + guiLog : "")
                 + (expect == Expect.Reject && accepted ? " | THE UNIT HELD THIS INVALID VALUE until it was put back" : ""));
            AddFinding(new Finding(row, what + " '" + shownText + "'", expected, actual, guiLog, Problem(what, text, expect, expected, actual), original, shownText, chip ? RejectChipText(tb) : shown));
            _failInARow++;
            if (_stopOnFail || _failInARow >= MaxFailuresInARow)
            {
                _stopRequested = true;
                if (!_stopOnFail)
                {
                    _log("INFO  | " + _failInARow + " failing inputs in a row - that looks like a problem with the unit or the GUI, stopping");
                }
            }
        }
        if (ok)
        {
            _failInARow = 0;
        }

        At("putting the original value back");
        // Back to the original: after a rejection the unit still holds the original, so the row's own refresh button (reads the unit again,
        // clears the red X, ~0.5 s) is enough; after an accepted value the unit holds the typed value, so the original is typed back.
        if (chip)
        {
            string unitValueChip = RejectChipText(FindValueElement(window, row) ?? tb);      // what the unit holds, as the GUI shows it next to the red X
            ClearRejection(window, row);      // the row refresh reads the value from the UNIT
            if (expect != Expect.Accept)
            {
                // the unit must still hold the original after the refusal - it must never keep the invalid value
                string held = "";
                for (int i = 0; i < 8; i++)
                {
                    var again = FindValueElement(window, row);
                    held = again == null ? "" : DisplayedValue(window, row, again);
                    if (SameValue(held, original))
                    {
                        break;
                    }

                    Thread.Sleep(200);
                }

                if (!SameValue(held, original) && held.Length > 0)
                {
                    if (held.All(c => c == '?') && text.Any(c => c > 127))
                    {
                        // Refused non-ASCII text (Hebrew): the refresh shows "????" - that is how this GUI/unit shows it and it is correct, not a failure. The question marks are
                        // deleted (Ctrl+A) and the original value is typed again.
                        _log("NOTE  | " + row.Label + " | after refusing '" + shownText + "' the refresh shows '" + held + "' - this is how the GUI/unit shows refused non-ASCII text (not a failure); the original value is typed back");
                        TypeAndEnter(FindValueElement(window, row) ?? tb, original);
                        WaitForRejectChipGone(window, row, 3000);
                        held = original;
                    }
                    else
                    {
                        // The field can still show the REFUSED text (this GUI does not revert it): that is not the unit. The unit's own value is the chip next to the red X, and a
                        // Fetch reads every value from the unit again and resets the field - only a value that is still wrong after that counts.
                        _log("NOTE  | " + row.Label + " | the field still shows '" + held + "' after the refusal" + (unitValueChip.Length > 0 ? " (next to the red X the unit reports '" + unitValueChip + "')" : "")
                             + " - reading the unit again (Fetch)");
                        FetchDiscard(window);
                        for (int i = 0; i < 8; i++)
                        {
                            var again = FindValueElement(window, row);
                            held = again == null ? "" : DisplayedValue(window, row, again);
                            if (SameValue(held, original))
                            {
                                break;
                            }

                            Thread.Sleep(200);
                        }
                    }
                }

                if (!SameValue(held, original) && held.Length > 0)
                {
                    string shownHeld = held.Length > 20 ? held.Substring(0, 20) + "..." : held;
                    _log("FAIL  | " + row.Label + " | " + what + " '" + shownText + "' | after the refusal the UNIT holds '" + shownHeld + "' instead of the original '" + original + "'");
                    AddFinding(new Finding(row, what + " '" + shownText + "'", "the unit keeps '" + original + "'", "unit holds '" + shownHeld + "'", guiLog, "After the value was refused, the unit holds “" + shownHeld + "” (the original value was “" + original + "”).", original, shownText, shownHeld));
                    _stopRequested |= _stopOnFail;
                    TypeAndEnter(FindValueElement(window, row) ?? tb, original);      // put the original back
                    WaitForRejectChipGone(window, row, 3000);
                }
            }
        }
        else
        {
            tb = FindValueElement(window, row) ?? tb;
            TypeAndEnter(tb, original);
            WaitForRejectChipGone(window, row, 3000);
            ClearRejection(window, row);
        }

        return ok;
    }

    // "must be higher than <other>" (Vin1 High Threshold Px vs Vin1 Low Threshold): the Min limit of this parameter can only be accepted while <other> is below it, and
    // <other> can go no lower than ITS Min - so <other> is set to its Min for this one input, the Min is typed (must be ACCEPTED), and <other> is put back (also on an error).
    // If the unit refuses Min here (e.g. both Min are 8 and "higher" is strict) it is reported as a failure of the Min input.
    private void CheckMinWithOtherAtItsMinimum(Window window, Row row, string original, Tip tip)
    {
        var other = _allRows.Where(r => r.Sheet == row.Sheet && r.Channel == row.Channel && tip.Other.StartsWith(r.Name, StringComparison.OrdinalIgnoreCase))
                            .OrderByDescending(r => r.Name.Length).FirstOrDefault();
        if (other == null || other.Scope != row.Scope)
        {
            _log("NOTE  | " + row.Label + " | the Min limit was not checked: the other parameter ('" + tip.Other + "') is not on this page of the sheet");
            return;
        }

        var otherTb = FindValueElement(window, other);
        string otherOriginal = otherTb == null ? "" : DisplayedValue(window, other, otherTb);
        var otherTip = otherTb == null ? new Tip() : ReadConstraints(otherTb);
        if (otherTb == null || otherOriginal.Length == 0 || !otherTip.Min.HasValue)
        {
            _log("NOTE  | " + row.Label + " | the Min limit was not checked: the value or the Min of '" + other.Name + "' could not be read");
            return;
        }

        string otherMin = Fmt(otherTip.Min);
        _tempChange = other.Label + " (original '" + otherOriginal + "')";
        try
        {
            At("setting " + other.Name + " to its minimum " + otherMin);
            TypeAndEnter(otherTb, otherMin);
            Thread.Sleep(600);
            string got = ReadFromUnit(window, other) ?? DisplayedValue(window, other, FindValueElement(window, other) ?? otherTb);
            if (!SameValue(got, otherMin))
            {
                _log("WARN  | " + row.Label + " | the Min limit was not checked: " + other.Name + " could not be set to its minimum " + otherMin + " (the unit holds '" + got + "')");
                return;
            }

            _log("INFO  | " + row.Label + " | " + other.Name + " temporarily set to its minimum " + otherMin + " (original '" + otherOriginal + "') to check the Min limit");
            string what = "Min (the limit itself) [" + other.Name + " temporarily set to its minimum " + otherMin + "]";
            string text = Fmt(tip.Min);
            if (!CheckInput(window, row, original, text, Expect.Accept, what, tip, record: false))
            {
                _log("NOTE  | " + row.Label + " | " + what + " '" + text + "' was refused - trying once more");
                CheckInput(window, row, original, text, Expect.Accept, what, tip);
            }
        }
        finally
        {
            At("putting " + other.Name + " back");
            var tb = FindValueElement(window, other);
            if (tb != null)
            {
                ClearRejection(window, other);
                TypeAndEnter(FindValueElement(window, other) ?? tb, otherOriginal);
                WaitForRejectChipGone(window, other, 3000);
                ClearRejection(window, other);
            }

            string back = ReadFromUnit(window, other) ?? "";
            if (!SameValue(back, otherOriginal))
            {
                FetchDiscard(window);
                var again = FindValueElement(window, other);
                back = again == null ? "" : DisplayedValue(window, other, again);
            }

            if (!SameValue(back, otherOriginal))
            {
                _log("FAIL  | " + other.Label + " | after the Min check of " + row.Name + " it shows '" + back + "' instead of the original '" + otherOriginal + "' - fix it on the unit");
                AddFinding(new Finding(other, "(restore after the Min check of " + row.Name + ")", "the original value '" + otherOriginal + "'", "unit holds '" + back + "'", "",
                                       "This parameter was temporarily set to its minimum to check " + row.Name + " and could not be put back: it holds “" + back + "” (original “" + otherOriginal + "”).",
                                       otherOriginal, otherMin, back));
                _stopRequested = true;
            }

            _tempChange = "";
        }
    }

    // The red X stays until the row is refreshed: press the row's "Refresh parameter value from unit" button (id = name id + "Refresh").
    private void ClearRejection(Window window, Row row)
    {
        var tb = FindValueElement(window, row);
        if (tb == null || !HasRejectChip(tb))
        {
            return;
        }

        var list = window.FindFirstDescendant(cf => cf.ByAutomationId("ParametersList"));
        var refresh = list?.FindFirstDescendant(cf => cf.ByAutomationId(row.NameId + "Refresh"));
        if (refresh != null && refresh.IsEnabled)
        {
            refresh.AsButton().Invoke();
            WaitForRejectChipGone(window, row, 3000);
            tb = FindValueElement(window, row);
        }

        if (tb != null && HasRejectChip(tb))
        {
            _log("NOTE  | " + row.Label + " | the red X is still shown after the row refresh - reading all parameters again (Fetch)");
            FetchDiscard(window);
        }
    }

    // ------------------------------------------------------------------ GUI helpers

    // The failure in plain words, as it appears in the Failures workbook ("does not accept Min (0)", "accepts a value above Max (26)", ...).
    private static string Problem(string what, string text, Expect expect, string expected, string actual)
    {
        string shown = text.Length > 20 ? text.Substring(0, 20) + "..." : text;
        if (what.StartsWith("Max (", StringComparison.Ordinal)) return "The field does not accept the maximum value (" + shown + ").";
        if (what.StartsWith("Min (", StringComparison.Ordinal))
        {
            int ctx = what.IndexOf('[');      // "[Vin1 Low Threshold temporarily set to its minimum 8]" - the condition the value was typed under
            return "The field does not accept the minimum value (" + shown + ")" + (ctx >= 0 ? " " + what.Substring(ctx) : "") + ".";
        }
        if (what == "Max+1") return "The field accepts a value above the maximum (" + shown + ").";
        if (what == "Min-1") return "The field accepts a value below the minimum (" + shown + ").";
        if (what == "letters") return "The field accepts letters (“" + shown + "”).";
        if (what == "digit mixed with letters") return "The field accepts a digit mixed with letters (“" + shown + "”).";
        if (what == "empty") return "The field accepts an empty value.";
        if (what == "Hebrew letters") return "The field accepts Hebrew letters.";
        if (what.StartsWith("Max+1 characters", StringComparison.Ordinal)) return "The field accepts text longer than the maximum (" + text.Length + " characters).";
        if (what.StartsWith("English letters", StringComparison.Ordinal)) return "The field does not accept valid text (English letters, digit, space, symbols).";
        return "Unexpected result for " + what + ": expected " + expected + ", the GUI shows " + actual + ".";
    }

    private static bool SameText(string a, string b)
    {
        if (double.TryParse(a, NumberStyles.Float, CultureInfo.InvariantCulture, out double x) && double.TryParse(b, NumberStyles.Float, CultureInfo.InvariantCulture, out double y))
        {
            return Math.Abs(x - y) < 1e-6;
        }

        return string.Equals(a.Trim(), b.Trim(), StringComparison.Ordinal);
    }

    private static bool SameValue(string a, string b) => SameText(a, b);

    private void TypeAndEnter(AutomationElement tb, string text)
    {
        // the keystrokes go to the FRONT window - it must be the GUI
        if (_window != null && !AppConnection.EnsureGuiInFront(_window))
        {
            throw new InvalidOperationException("Power Rider Studio is not the front window (the keystrokes would go to another window) and could not be brought forward");
        }

        try
        {
            tb.Focus();
        }
        catch (Exception ex) when (ex is InvalidOperationException or System.Runtime.InteropServices.COMException)
        {
            try
            {
                tb.Patterns.ScrollItem.PatternOrDefault?.ScrollIntoView();
                Thread.Sleep(150);
                tb.Focus();
            }
            catch
            {
                // typed into whatever has the focus - the checks that follow notice a wrong result
            }
        }

        Thread.Sleep(120);
        Interlocked.Increment(ref _writes);      // Enter below applies the value to the unit
        Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_A);
        if (text.Length == 0)
        {
            Keyboard.Type(VirtualKeyShort.DELETE);
        }
        else
        {
            TextCopy.ClipboardService.SetText(text);
            Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_V);
        }

        Thread.Sleep(150);
        Keyboard.Type(VirtualKeyShort.ENTER);
    }

    // The ListBoxItem of the row = two levels above the value field.
    private static AutomationElement? RowItem(AutomationElement valueEl)
    {
        var e = valueEl;
        for (int i = 0; i < 2 && e != null; i++)
        {
            e = e.Parent;
        }

        return e;
    }

    // Is the red X shown next to the field? (an icon, found by its colour - see RowIcons). The GUI shows the unit's own value in a chip next to the field after a
    // set, also when the value was accepted (GUI 2.1.2-beta.2), so the chip text alone does not tell; the red X does.
    private static bool HasRejectChip(AutomationElement valueEl) => RowIcons.Read(valueEl).RedX;

    // The text of the chip (what the unit holds, e.g. "-1" or the last accepted value), "" when there is none.
    private static string RejectChipText(AutomationElement valueEl)
    {
        try
        {
            var item = RowItem(valueEl);
            if (item == null)
            {
                return "";
            }

            double right = valueEl.BoundingRectangle.Right;
            return Walk(item, 0).Where(t => t.ControlType == ControlType.Text && string.IsNullOrEmpty(t.AutomationId) && !string.IsNullOrWhiteSpace(t.Name)
                                            && !t.Properties.IsOffscreen.ValueOrDefault && t.BoundingRectangle.Width > 0 && t.BoundingRectangle.X >= right - 2)
                                .Select(t => t.Name.Trim()).FirstOrDefault() ?? "";
        }
        catch
        {
            return "";
        }
    }

    // The row's own refresh button reads the value from the UNIT again; returns what the field shows afterwards (null = the button is not available).
    private string? ReadFromUnit(Window window, Row row)
    {
        var list = window.FindFirstDescendant(cf => cf.ByAutomationId("ParametersList"));
        var refresh = list?.FindFirstDescendant(cf => cf.ByAutomationId(row.NameId + "Refresh"));
        if (refresh == null || !refresh.IsEnabled)
        {
            return null;
        }

        refresh.AsButton().Invoke();
        Thread.Sleep(900);
        var tb = FindValueElement(window, row);
        return tb == null ? null : DisplayedValue(window, row, tb);
    }

    // The verdict on a typed value from the icons the GUI shows: a green V = accepted, a red X on its own = refused. A red X that was left over from the PREVIOUS
    // input is seen at once, before the GUI has answered the new one, so a red X without the GUI's own error counts only after 1 s; and when both icons or none
    // are shown at the end, the red X counts only if the GUI logged an error or the field does not hold the typed value.
    private bool WaitForReject(Window window, Row row, int timeoutMs, Func<bool> guiError, Func<bool> fieldHasTyped, out bool seenX, out bool seenV)
    {
        var sw = Stopwatch.StartNew();
        Thread.Sleep(250);
        seenX = false;
        seenV = false;
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            var tb = FindValueElement(window, row);
            if (tb != null)
            {
                (seenX, seenV) = RowIcons.Read(tb);
            }

            if (seenV && !seenX)
            {
                return false;
            }

            if (seenX && !seenV && (sw.ElapsedMilliseconds >= 1000 || guiError()))
            {
                return true;
            }

            Thread.Sleep(120);
        }

        return seenX && (guiError() || !fieldHasTyped());
    }

    private bool WaitForRejectChip(Window window, Row row, int timeoutMs)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            var tb = FindValueElement(window, row);
            if (tb != null && HasRejectChip(tb))
            {
                return true;
            }

            Thread.Sleep(120);
        }

        return false;
    }

    private void WaitForRejectChipGone(Window window, Row row, int timeoutMs)
    {
        var sw = Stopwatch.StartNew();
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            var tb = FindValueElement(window, row);
            if (tb == null || !HasRejectChip(tb))
            {
                return;
            }

            Thread.Sleep(150);
        }
    }

    private void FetchDiscard(Window window)
    {
        var fetch = window.FindFirstDescendant(cf => cf.ByHelpText("Refresh parameters from unit"));
        if (fetch != null && fetch.IsEnabled)
        {
            fetch.AsButton().Invoke();
            Thread.Sleep(2500);
            ConfigureNav.WaitFetchDone(window, 120000, _log);
        }
    }

    // ---- locate: select the tree page, scroll to the row (the previous row's position is the starting point)

    private AutomationElement? FindValueElement(Window window, Row row)
    {
        var list = window.FindFirstDescendant(cf => cf.ByAutomationId("ParametersList"));
        var direct = list?.FindFirstDescendant(cf => cf.ByAutomationId(row.ValueId));
        if (direct != null)
        {
            return direct;
        }

        if (list == null)
        {
            return null;
        }

        var scroll = list.Patterns.Scroll.PatternOrDefault;
        bool scrollable = scroll != null && scroll.VerticallyScrollable.ValueOrDefault;
        double view = scrollable ? Math.Max(5, scroll!.VerticalViewSize.ValueOrDefault) : 100;
        double estimate = _scopeRows.TryGetValue(row.Scope, out var rowsOfScope)
            ? ConfigureNav.EstimateScrollPercent(rowsOfScope.Count, rowsOfScope.IndexOf(row.ExcelRow), view) : -1;
        double lastPct = _lastPct.TryGetValue(row.Scope, out double lp) ? lp : -1;
        foreach (double pct in ConfigureNav.ScrollTries(estimate, lastPct, view))
        {
            if (scrollable)
            {
                ConfigureNav.ScrollSteps++;
                scroll!.SetScrollPercent(-1, pct);
                Thread.Sleep(150);
            }

            var el = list.FindFirstDescendant(cf => cf.ByAutomationId(row.ValueId));
            if (el != null)
            {
                _lastPct[row.Scope] = pct;
                return el;
            }

            if (!scrollable)
            {
                break;
            }
        }

        return null;
    }

    // ---- reading

    private sealed class Tip
    {
        public double? Min, Max;
        public string TypeToken = "";
        public string Direction = "";      // "lower" / "higher" when the constraint says "must be lower/higher than <another parameter>"
        public string Other = "";          // the text after "must be higher than" (starts with the other parameter's name, e.g. "Vin1 Low Threshold. ...")
    }

    private static readonly Regex ConstraintsRegex = new(@"Min:\s*(-?[\d.,]+)[^\w-]+Max:\s*(-?[\d.,]+)", RegexOptions.Compiled);
    private static readonly HashSet<string> DataTypeTokens = new(StringComparer.OrdinalIgnoreCase) { "String", "Integer", "Float", "Number", "Boolean" };

    private static Tip ReadConstraints(AutomationElement valueEl)
    {
        var tip = new Tip();
        try
        {
            var row = valueEl.Parent;
            for (int up = 0; up < 2 && row != null; up++, row = row.Parent)      // the row and its list item - never the whole list
            {
                var texts = Walk(row, 0).Where(x => x.ControlType == ControlType.Text).ToList();
                var c = texts.Select(t => (t.Name ?? "").Trim()).FirstOrDefault(t => t.Contains("Constraints:", StringComparison.Ordinal));
                if (c == null)
                {
                    continue;
                }

                var m = ConstraintsRegex.Match(c);
                if (m.Success)
                {
                    if (double.TryParse(m.Groups[1].Value.Replace(",", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out double mn))
                    {
                        tip.Min = mn;
                    }

                    if (double.TryParse(m.Groups[2].Value.Replace(",", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out double mx))
                    {
                        tip.Max = mx;
                    }
                }

                var dep = Regex.Match(c, @"must be\s+(lower|less|smaller|higher|greater|bigger)\s+(?:than|then)?\s*([^\r\n]*)", RegexOptions.IgnoreCase);
                if (dep.Success)
                {
                    tip.Other = dep.Groups[2].Value.Trim();
                    tip.Direction = dep.Groups[1].Value.ToLowerInvariant() is "lower" or "less" or "smaller" ? "lower" : "higher";
                }

                // text parameters can state their length instead of Min/Max, e.g. "Default: Single Ch. n MAX 20 bytes"
                var bytes = Regex.Match(c, @"MAX\s+(\d+)\s*bytes", RegexOptions.IgnoreCase);
                if (!tip.Max.HasValue && bytes.Success)
                {
                    tip.Max = int.Parse(bytes.Groups[1].Value);
                }

                tip.TypeToken = texts.Where(x => !x.Properties.IsOffscreen.ValueOrDefault).Select(x => (x.Name ?? "").Trim())
                    .FirstOrDefault(x => Regex.IsMatch(x, @"^\[.+\]$") || x.Equals("String", StringComparison.OrdinalIgnoreCase)) ?? "";
                break;
            }
        }
        catch
        {
            // best effort
        }

        return tip;
    }

    // The value the row shows: the visible text over the field (the editor itself exposes an empty value to UIA).
    private static string DisplayedValue(Window window, Row row, AutomationElement valueEl)
    {
        try
        {
            string name = (window.FindFirstDescendant(cf => cf.ByAutomationId(row.NameId))?.Name ?? row.Name).Trim();
            double left = valueEl.BoundingRectangle.Left, right = valueEl.BoundingRectangle.Right;
            var item = RowItem(valueEl);
            if (item == null)
            {
                return "";
            }

            return Walk(item, 0)
                .Where(t => t.ControlType == ControlType.Text && string.IsNullOrEmpty(t.AutomationId) && !t.Properties.IsOffscreen.ValueOrDefault
                            && t.BoundingRectangle.Width > 0 && t.BoundingRectangle.X >= left - 2 && t.BoundingRectangle.X < right - 2)
                .Select(t => (t.Name ?? "").Trim())
                .FirstOrDefault(t => t.Length > 0 && t != name && !DataTypeTokens.Contains(t) && !t.StartsWith("[")) ?? "";
        }
        catch
        {
            return "";
        }
    }

    private static IEnumerable<AutomationElement> Walk(AutomationElement e, int depth)
    {
        var walker = e.Automation.TreeWalkerFactory.GetRawViewWalker();
        for (var c = walker.GetFirstChild(e); c != null; c = walker.GetNextSibling(c))
        {
            yield return c;
            if (depth < 6)
            {
                foreach (var x in Walk(c, depth + 1))
                {
                    yield return x;
                }
            }
        }
    }

    private static string Fmt(double? d) => d.HasValue ? d.Value.ToString("0.###", CultureInfo.InvariantCulture) : "?";

    private static bool IsLogicParam(string name) => name.Contains("Logic Control", StringComparison.OrdinalIgnoreCase) || name.Equals("Logic Expression", StringComparison.OrdinalIgnoreCase);

    // ---- the GUI's own log ("Failed to set <parameter>", "Parameter error: ...")

    private string ReadAppLog()
    {
        try
        {
            return _logElement == null ? "" : string.Join("\n", _logElement.FindAllDescendants(cf => cf.ByControlType(ControlType.Text)).Select(t => t.Name ?? ""));
        }
        catch
        {
            return "";
        }
    }

    private string NewGuiErrors(string before)
    {
        var old = before.Split('\n').ToHashSet();
        var added = ReadAppLog().Split('\n').Where(l => l.Length > 0 && !old.Contains(l) && Regex.IsMatch(l, @"Failed|error", RegexOptions.IgnoreCase));
        return string.Join(" ## ", added);
    }

    // ---- Excel

    private List<Row> ReadRows()
    {
        string temp = Path.Combine(Path.GetTempPath(), "ConfigValidation_" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.Copy(_excelPath, temp, true);
        try
        {
            using var wb = new XLWorkbook(temp);
            var result = new List<Row>();
            if (!wb.Worksheets.Contains(_unit.ParameterSheet))
            {
                return result;
            }

            foreach (var r in wb.Worksheet(_unit.ParameterSheet).RowsUsed().Skip(1))
            {
                string name = r.Cell(3).GetString().Trim(), nameId = r.Cell(2).GetString().Trim(), valueId = r.Cell(4).GetString().Trim();
                if (name.Length == 0 || nameId.Length == 0 || valueId.Length == 0)
                {
                    continue;
                }

                int.TryParse(r.Cell(1).GetString().Trim(), out int channel);
                result.Add(new Row(_unit.ParameterSheet, r.RowNumber(), channel, name, nameId, valueId, r.Cell(5).GetString().Trim()));
            }

            return result;
        }
        finally
        {
            try { File.Delete(temp); } catch { /* temp copy */ }
        }
    }

    private static string ReadGuiVersion()
    {
        try
        {
            var info = Process.GetProcessesByName("PowerRiderStudio").FirstOrDefault()?.MainModule?.FileVersionInfo;
            return info == null ? "" : (info.ProductVersion ?? info.FileVersion ?? "");
        }
        catch
        {
            return "";
        }
    }

    // One line per failing INPUT (Excel row first - give this file at 'Excel rows to run' to run only these parameters again), each with WHY it failed in
    // plain words ("does not accept Min (0)", "accepts a value above Max (26)", ...), so every range problem of a parameter is listed, not only the first.
    // Failures\InputValidation_<PN>_<run start>.xlsx next to the log files. Sheets: Failed, Run (status / time), Not run (rows never reached).
    // Called after every failing input and at the end, whatever the end is. Written when something failed or the run did not complete normally.
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
                if (report.Finished && _findings.Count == 0 && report.Status.StartsWith("Completed", StringComparison.Ordinal))
                {
                    try { File.Delete(report.Path); } catch { /* the "running" version written at the start - nothing failed, nothing to keep */ }
                    return null;
                }

                using var wb = new XLWorkbook();
                var ws = wb.AddWorksheet(RunReport.SheetFailed);
                string[] head = { "Excel row", "Channel", "Parameter", "Failure type", "Previous value", "Value we tried to write", "Value read from the unit", "Unit", "GUI version", "Run time" };
                if (Looping)
                {
                    head = head.Append("Loop round").ToArray();
                }

                for (int i = 0; i < head.Length; i++)
                {
                    ws.Cell(1, i + 1).Value = head[i];
                    ws.Cell(1, i + 1).Style.Font.Bold = true;
                }

                int r = 2;
                foreach (var f in _findings)
                {
                    ws.Cell(r, 1).Value = f.Row.ExcelRow;
                    ws.Cell(r, 2).Value = f.Row.Channel > 0 ? f.Row.Channel.ToString() : "System";
                    ws.Cell(r, 3).Value = f.Row.Name;
                    ws.Cell(r, 4).Value = f.Problem;
                    ws.Cell(r, 5).Value = f.Original;
                    ws.Cell(r, 6).Value = f.Typed;
                    ws.Cell(r, 7).Value = f.Shown;
                    ws.Cell(r, 8).Value = _unit.Name;
                    ws.Cell(r, 9).Value = _guiVersion;
                    ws.Cell(r, 10).Value = f.At.ToString("yyyy-MM-dd HH:mm");
                    if (Looping)
                    {
                        ws.Cell(r, 11).Value = f.Round > 0 ? "Round " + f.Round + (_unlimited ? "" : " of " + _rounds) : "";
                    }

                    r++;
                }

                ws.Columns().AdjustToContents();

                report.AddRunSheet(wb, new[]
                {
                    ("Unit", _unit.Name),
                    ("Rows of this run", _rowsDescription),
                    ("Loop", LoopText),
                });
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

    private static bool GuiVersionIsAtLeast(string version, int major, int minor, int patch)
    {
        var m = Regex.Match(version ?? "", @"^(\d+)\.(\d+)\.(\d+)");
        if (!m.Success)
        {
            return true;       // unknown format: do not block
        }

        int[] v = { int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value), int.Parse(m.Groups[3].Value) };
        int[] need = { major, minor, patch };
        for (int i = 0; i < 3; i++)
        {
            if (v[i] != need[i])
            {
                return v[i] > need[i];
            }
        }

        return true;
    }
}
