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
// confirmation question here; the only prompt is the "unit was reset to default?" check of the main menu (Program.cs).
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

    private AutomationElement? _logElement;

    public ConfigureChangeAndRestoreAllParametersTest(string excelPath, Action<string> log, UnitProfile unit)
    {
        _excelPath = excelPath;
        _log = log;
        _unit = unit;
    }

    private string[] Sheets => new[] { _unit.ParameterSheet };

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
        if (_failStep.Length == 0)
        {
            _failParam = who;
            _failStep = step;
            _failWhy = why;
            _failCheck = check;
        }
    }

    private void LogResult(bool passed, int done, int passedCount, bool restored, int skipped)
    {
        _log("RESULT| " + new string('=', 90));
        _log("RESULT| " + (passed ? "PASSED" : "FAILED") + " | unit=" + _unit.Name + " | parameters changed+restored=" + passedCount + " of " + done + " checked | skipped=" + skipped);
        if (!passed && _failStep.Length > 0)
        {
            _log("RESULT| failed parameter : " + _failParam);
            _log("RESULT| failed at step   : " + _failStep);
            _log("RESULT| reason           : " + _failWhy);
            _log("RESULT| what to check    : " + _failCheck);
            _log("RESULT| unit state       : " + (_failParam == "(test setup)" ? "nothing was changed on the unit"
                                                       : restored ? "parameter is back at its Excel default (verified)" : "NOT verified - check this parameter on the unit manually"));
        }

        _log("RESULT| " + new string('=', 90));
    }

    // ------------------------------------------------------------------ entry point

    public bool Run()
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

        _log("START | Configure change-and-restore test | unit=" + _unit.Name + " (sheet '" + _unit.ParameterSheet + "') | GUI version=" + _guiVersion
             + " | every parameter of the sheet");

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
        var excluded = all.Where(NotChanged).ToList();
        var todo = all.Where(p => !NotChanged(p)).ToList();
        _log("INFO  | parameters in Excel=" + all.Count + " | never changed (baud rates, addresses, logic, group control)=" + excluded.Count + " | to check in this run=" + todo.Count);
        if (todo.Count == 0)
        {
            _log("FAIL  | nothing to check - the sheet has no changeable parameter");
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
        if (!WaitFetchDone(window, 300000, "Configure tab just opened"))
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

        int done = 0, passed = 0, skipped = 0, index = 0;
        var watch = Stopwatch.StartNew();
        foreach (var p in todo)
        {
            index++;
            _log("DRAW  | " + index + "/" + todo.Count + " | " + p.Label);
            var outcome = TestOne(window, search, p);
            if (outcome == Outcome.Skipped)
            {
                skipped++;
                continue;
            }

            done++;
            if (outcome == Outcome.Failed)
            {
                _log("STOPPED | halted after first failure - remaining parameters were not checked");
                ClearSearch(window);
                LogResult(false, done, passed, _lastRestored, skipped);
                return false;
            }

            passed++;
        }

        ClearSearch(window);
        _log("DONE  | checked=" + done + " passed=" + passed + " skipped=" + skipped + " | " + watch.Elapsed.TotalMinutes.ToString("0.0") + " min");
        LogResult(true, done, passed, true, skipped);
        return true;
    }

    // ------------------------------------------------------------------ one parameter

    private enum Outcome { Passed, Failed, Skipped }

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

    private (AutomationElement? name, AutomationElement? value) LocateOnPage(Window window, Param p)
    {
        try
        {
            string? crumb = ConfigureNav.SelectScope(window, p.Scope);      // re-selecting rebuilds the list (also after Save/Fetch)
            if (crumb == null)
            {
                return (null, null);
            }

            var list = window.FindFirstDescendant(cf => cf.ByAutomationId("ParametersList"));
            if (list == null)
            {
                return (null, null);
            }

            double view = 100;
            var scroll = list.Patterns.Scroll.PatternOrDefault;
            bool scrollable = scroll != null && scroll.VerticallyScrollable.ValueOrDefault;
            if (scrollable)
            {
                view = Math.Max(5, scroll!.VerticalViewSize.ValueOrDefault);
            }

            double start = _lastPct.TryGetValue(p.Scope, out double last) ? Math.Max(0, last - view / 2) : 0;
            foreach (double from in new[] { start, 0 })
            {
                for (double pct = from; ; pct += view / 2)
                {
                    list = window.FindFirstDescendant(cf => cf.ByAutomationId("ParametersList")) ?? list;
                    scroll = list.Patterns.Scroll.PatternOrDefault;
                    if (scrollable && scroll != null)
                    {
                        scroll.SetScrollPercent(-1, Math.Min(pct, 100));
                        Thread.Sleep(150);
                    }

                    var nameEl = list.FindFirstDescendant(cf => cf.ByAutomationId(p.NameId));
                    var valueEl = list.FindFirstDescendant(cf => cf.ByAutomationId(p.ValueId));
                    if (nameEl != null && valueEl != null && (nameEl.Name ?? "").Trim() == p.Name)
                    {
                        _lastPct[p.Scope] = Math.Min(pct, 100);
                        return (nameEl, valueEl);
                    }

                    if (!scrollable || pct >= 100)
                    {
                        break;
                    }
                }

                if (start == 0)
                {
                    break;
                }
            }
        }
        catch (Exception ex) when (ex is FlaUI.Core.Exceptions.PropertyNotSupportedException or FlaUI.Core.Exceptions.ElementNotAvailableException or System.Runtime.InteropServices.COMException or InvalidOperationException)
        {
            _log("NOTE  | " + p.Label + " | page scan interrupted (" + ex.GetType().Name + ")");
        }

        return (null, null);
    }

    private Outcome TestOne(Window window, TextBox search, Param p)
    {
        _lastRestored = true;

        _log("STEP  | 1 find the parameter (tree page scan)");
        var (nameEl, valueEl) = LocateRow(window, search, p, refresh: false);
        if (nameEl == null || valueEl == null)
        {
            Fail(p, "1 locate the parameter", "the parameter row was not found on screen (ids " + p.NameId + " / " + p.ValueId + ")",
                 "is the Configure tab open? does the page '" + (p.Channel > 0 ? "Channel " + p.Channel : "System") + "' list the parameter?");
            return Outcome.Failed;
        }

        string type = valueEl.ControlType.ToString();

        _log("STEP  | 2 read the current value and the constraints (range)");
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
        var tip = ReadConstraintsFromRow(valueEl);
        if (!tip.Found)
        {
            tip = ReadTooltip(window, nameEl);
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

                ok = ChangeAndVerify(window, search, p, valueEl, before, plan[i].value, plan[i].why);
                before = plan[i].value;
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
        bool restored = RestoreDefault(window, search, p);
        _lastRestored = restored;
        return ok && restored ? Outcome.Passed : Outcome.Failed;
    }

    // One change: set `newValue`, Save, Fetch and verify that the unit returns it. False = failed (reason already logged with Fail).
    private bool ChangeAndVerify(Window window, TextBox search, Param p, AutomationElement valueEl, string before, string newValue, string why)
    {
        SetValue(window, valueEl, newValue);
        string shownNow = ReadParam(window, p, out _);
        _log("SET   | " + p.Label + " | '" + before + "' -> '" + newValue + "' | the GUI now shows '" + shownNow + "' | " + why
             + " | Save to Unit Flash is " + (FindByHelp(window, "Save to Unit Flash")?.IsEnabled == true ? "enabled" : "DISABLED"));
        if (!SameValue(shownNow, newValue))
        {
            Fail(p, "3 set the new value", "after setting '" + newValue + "' the control shows '" + shownNow + "'",
                 "the value was rejected/not committed (range, focus) - look at the Configure screen");
            return false;
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
            Fail(p, "3 set the new value", "the value shows '" + shownNow + "' but 'Save to Unit Flash' stayed DISABLED - the GUI did not register a pending change",
                 "the edit was probably not committed (needs Enter/focus change?) or equals the stored value");
            return false;
        }

        _log("STEP  | 4 Save to Unit Flash" + (DoReset ? ", Reset Unit" : "") + ", Fetch Parameters");
        if (!SaveResetFetch(window, search, p))
        {
            return false;
        }

        _log("STEP  | 5 verify the value after Save / Fetch");
        var (_, v2) = LocateRow(window, search, p, refresh: true);
        string after = v2 == null ? "<NOT FOUND>" : ReadParam(window, p, out _);
        bool changed = SameValue(after, newValue);
        _log((changed ? "PASS  | " : "FAIL  | ") + p.Label + " | after Save/Fetch: expected='" + newValue + "' actual='" + after + "'");
        if (!changed)
        {
            Fail(p, "5 verify after Save/Fetch", "expected '" + newValue + "' but the unit returned '" + after + "' after fetching",
                 "the unit did not keep the value (rejected by the unit, or Save did not complete)");
        }

        return changed;
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

        int decimals = cleaned.Contains('.') ? cleaned.Length - cleaned.IndexOf('.') - 1 : 0;
        double step = decimals == 0 ? 1 : Math.Pow(10, -decimals);
        double value = cur0;
        for (int i = 0; i < 50 && Math.Abs(value - cur0) < step / 2; i++)
        {
            double raw = lo + _random.NextDouble() * (hi - lo);
            value = decimals == 0 ? Math.Round(raw) : Math.Round(raw, decimals);
            value = Math.Max(lo, Math.Min(hi, value));
        }

        if (Math.Abs(value - cur0) < step / 2)
        {
            why = "could not find a different value inside " + Fmt(lo) + ".." + Fmt(hi);
            return null;
        }

        why = "random value inside " + Fmt(lo) + ".." + Fmt(hi) + direction;
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
    private bool WaitFetchDone(Window window, int timeoutMs, string reason)
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
                // the status text flickers while the GUI works - it must stay away for 1 s in a row
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
            announced = true;
            if (progress != lastProgress)
            {
                lastProgress = progress;
                progressSince = sw.ElapsedMilliseconds;
            }
            else if (sw.ElapsedMilliseconds - progressSince >= 8000 && FreshSearch(window) is { IsEnabled: true })
            {
                // same counter for 8 s while the search box is usable = stale status text, not a real load (a real load disables the search)
                _log("WARN  | the status bar is stuck on 'Fetching parameters... " + progress + "' for 8 s but the search box is enabled - treating the load as finished (" + reason + ")");
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
            }

            Thread.Sleep(250);
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

        // text box: paste the whole value, then commit with Tab
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
            for (int up = 0; up < 3 && row != null && tip.Text.Length == 0; up++, row = row.Parent)
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
        if (!Press(window, "Refresh parameters from unit", "Fetch", FetchTimeoutMs, p)) return false;
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
                        _log("NOTE  | " + p.Label + " | the GUI log reports a problem after '" + label + "' (ignored, the read-back after Fetch decides)");
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
