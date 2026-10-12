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
//                         (the two accepted ones only when the constraint does not depend on another parameter)
//   number, no range    : "abc" and empty must be rejected
//   text                : Hebrew letters must be rejected, empty must be rejected (Min >= 1), Max+1 characters must be rejected or cut to
//                         Max, a normal text with English letters, digits, space and symbols must be accepted
// After every input the original value is typed back, so the unit ends exactly as it was (RAM and flash).
//
// Scope: System parameters and channel 1 (a parameter's validation is the same in every channel). Never touched: baud rates, addresses,
// Logic Control / Logic Expression, Group Control.
//
// Like test 3, the run asks at the start: stop at the first failing input (default) or continue and list every failure at the end (log + a
// workbook in the Failures folder next to the logs), and which Excel rows to run (Enter = all, or "54-70, 72, 100-").
public class ConfigureInputValidationTest
{
    private static readonly int[] ValidationChannels = { 1 };
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

    private sealed record Finding(Row Row, string Input, string Expected, string Actual, string GuiLog);

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
    private int _failInARow;
    private int _inputs, _passed, _skippedParams, _params;
    private readonly Dictionary<int, double> _lastPct = new();
    private AutomationElement? _logElement;

    public ConfigureInputValidationTest(string excelPath, Action<string> log, UnitProfile unit)
    {
        _excelPath = excelPath;
        _log = log;
        _unit = unit;
    }

    public bool Run()
    {
        string onlyName = Environment.GetEnvironmentVariable("VALIDATION_ONLY") ?? "";      // developer switch: run one parameter only (name contains)

        List<Row> rows;
        try
        {
            rows = ReadRows();
        }
        catch (Exception ex)
        {
            _log("FAIL  | could not read the sheet '" + _unit.ParameterSheet + "': " + ex.Message);
            return false;
        }

        var todo = rows.Where(r => (r.Channel == 0 || ValidationChannels.Contains(r.Channel)) && r.Channel <= _unit.Channels
                                   && !NeverChange.Contains(r.Name) && !IsLogicParam(r.Name) && r.Scope != ConfigureNav.GroupScope
                                   && (onlyName.Length == 0 || r.Name.Contains(onlyName, StringComparison.OrdinalIgnoreCase))).ToList();
        Console.WriteLine();
        Console.WriteLine("When an input fails:");
        Console.WriteLine("  1. stop at the first failure  [default]");
        Console.WriteLine("  2. continue; every failure is listed at the end (and in the Failures folder)");
        Console.Write("Choice [1]: ");
        _stopOnFail = (Console.ReadLine() ?? "").Trim() != "2";

        RowSelection? selection;
        while (true)
        {
            Console.Write("Excel rows to run (Enter = all;  e.g.  54-70, 72, 100-  ): ");
            if (RowSelection.TryParse(Console.ReadLine(), out selection, out string selectionError))
            {
                break;
            }

            Console.WriteLine(selectionError);
        }

        Console.WriteLine();
        if (selection != null)
        {
            todo = todo.Where(r => selection.Contains(r.ExcelRow)).ToList();
        }

        _guiVersion = ReadGuiVersion();
        _log("START | Configure input validation | unit=" + _unit.Name + " | scope: System + channel(s) " + string.Join(",", ValidationChannels)
             + " | parameters=" + todo.Count + (onlyName.Length > 0 ? " | only names containing '" + onlyName + "'" : "") + " | on failure: " + (_stopOnFail ? "stop" : "continue")
             + " | rows: " + (selection == null ? "all" : selection.ToString()) + " | NOTHING is saved to flash");
        if (todo.Count == 0)
        {
            _log("FAIL  | no parameter to check");
            return false;
        }

        using var automation = new UIA3Automation();
        var window = AppConnection.Attach(automation);
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
                index++;
                _log("DRAW  | " + index + "/" + todo.Count + " | " + row.Label);
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
                    _findings.Add(new Finding(row, "(stuck)", "no hang", "no progress for " + NoProgressMinutes + " min - last action: " + _where, ""));
                    _log("FAIL  | " + row.Label + " | STEP: " + _where + " | WHY: no progress for " + NoProgressMinutes + " minutes | CHECK: the GUI may be frozen or a UI call is not returning; check the unit and the GUI");
                    _stopRequested = true;
                }

                if (_stopRequested)
                {
                    _log("STOPPED | " + (finished ? "halted after the first failing input" : "the run is stuck and cannot continue safely") + " - remaining parameters were not checked");
                    _log("RESUME| to continue later answer 'Excel rows to run' with:  " + row.ExcelRow + "-  (this parameter again)");
                    goto Summary;
                }
            }
        }

        Summary:

        _log("DONE  | parameters=" + _params + " skipped=" + _skippedParams + " | inputs checked=" + _inputs + " passed=" + _passed + " findings=" + _findings.Count + " | " + watch.Elapsed.TotalMinutes.ToString("0.0") + " min");
        _log("RESULT| " + new string('=', 90));
        _log("RESULT| " + (_findings.Count == 0 ? "PASSED" : "FAILED") + " | unit=" + _unit.Name + " | inputs checked=" + _inputs + " | failures=" + _findings.Count + " | mode=" + (_stopOnFail ? "stop on fail" : "continue"));
        foreach (var f in _findings)
        {
            _log("RESULT|   row " + f.Row.ExcelRow + " | " + f.Row.Label.Split('(')[0].Trim() + " | input " + f.Input + " | expected " + f.Expected + " | GUI: " + f.Actual + (f.GuiLog.Length > 0 ? " | GUI log: " + f.GuiLog : ""));
        }

        string? file = WriteFailuresFile();
        if (file != null)
        {
            _log("RESULT| failures file : " + file);
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
             + (tip.Direction.Length > 0 ? " | must be " + tip.Direction + " than another parameter (only the " + (tip.Direction == "lower" ? "Min" : "Max") + " limit is typed as a valid value)" : ""));
        if (original.Length == 0)
        {
            _skippedParams++;
            _log("SKIP  | " + row.Label + " | the current value could not be read");
            return;
        }

        _params++;
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
            CheckInput(window, row, original, text, expect, what, tip);
            if (_stopRequested)
            {
                break;       // the original value has been put back by CheckInput; the run stops after this parameter
            }
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
    }

    private void CheckInput(Window window, Row row, string original, string text, Expect expect, string what, Tip tip)
    {
        // a red-X chip left over from the previous input would look like a rejection of this one: start from a clean field
        ClearRejection(window, row);

        string before = ReadAppLog();
        var tb = FindValueElement(window, row)!;
        At("typing '" + (text.Length > 20 ? text.Substring(0, 20) + "..." : text) + "'");
        TypeAndEnter(tb, text);
        bool chip = WaitForRejectChip(window, row, RejectWaitMs);
        tb = FindValueElement(window, row) ?? tb;
        string shown = DisplayedValue(window, row, tb);
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

        _inputs++;
        string shownText = text.Length > 20 ? text.Substring(0, 20) + "...(" + text.Length + ")" : text;
        if (ok)
        {
            _passed++;
            _log("PASS  | " + row.Label + " | " + what + " '" + shownText + "' | expected " + expected + " | " + actual + (guiLog.Length > 0 ? " | GUI log: " + guiLog : ""));
        }
        else
        {
            _log("FAIL  | " + row.Label + " | " + what + " '" + shownText + "' | expected " + expected + " | " + actual + (guiLog.Length > 0 ? " | GUI log: " + guiLog : "")
                 + (expect == Expect.Reject && accepted ? " | THE UNIT HELD THIS INVALID VALUE until it was put back" : ""));
            _findings.Add(new Finding(row, what + " '" + shownText + "'", expected, actual, guiLog));
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
            ClearRejection(window, row);
        }
        else
        {
            tb = FindValueElement(window, row) ?? tb;
            TypeAndEnter(tb, original);
            WaitForRejectChipGone(window, row, 3000);
            ClearRejection(window, row);
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

    // The "chip": a visible text to the RIGHT of the field that is not a labelled element - it shows the last value the unit accepted next to the red X.
    private static bool HasRejectChip(AutomationElement valueEl)
    {
        try
        {
            var item = RowItem(valueEl);
            if (item == null)
            {
                return false;
            }

            double right = valueEl.BoundingRectangle.Right;
            return Walk(item, 0).Any(t => t.ControlType == ControlType.Text && string.IsNullOrEmpty(t.AutomationId) && !string.IsNullOrWhiteSpace(t.Name)
                                          && !t.Properties.IsOffscreen.ValueOrDefault && t.BoundingRectangle.Width > 0 && t.BoundingRectangle.X >= right - 2);
        }
        catch
        {
            return false;
        }
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
        double start = _lastPct.TryGetValue(row.Scope, out double last) ? Math.Max(0, last - view / 2) : 0;
        foreach (double from in new[] { start, 0 })
        {
            for (double pct = from; ; pct += view / 2)
            {
                if (scrollable)
                {
                    scroll!.SetScrollPercent(-1, Math.Min(pct, 100));
                    Thread.Sleep(150);
                }

                var el = list.FindFirstDescendant(cf => cf.ByAutomationId(row.ValueId));
                if (el != null)
                {
                    _lastPct[row.Scope] = Math.Min(pct, 100);
                    return el;
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

        return null;
    }

    // ---- reading

    private sealed class Tip
    {
        public double? Min, Max;
        public string TypeToken = "";
        public string Direction = "";      // "lower" / "higher" when the constraint says "must be lower/higher than <another parameter>"
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

                var dep = Regex.Match(c, @"must be\s+(lower|less|smaller|higher|greater|bigger)", RegexOptions.IgnoreCase);
                if (dep.Success)
                {
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

    // One row per failing input (Excel row first): Failures\InputValidation_<PN>_<time>.xlsx next to the log files.
    private string? WriteFailuresFile()
    {
        if (_findings.Count == 0)
        {
            return null;
        }

        try
        {
            string dir = Path.Combine(AppContext.BaseDirectory, "Failures");
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, "InputValidation_" + _unit.PartNumber + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Failed");
            string[] headers = { "Excel row", "Channel", "Parameter", "Input", "Expected", "GUI result", "GUI log", "Unit", "GUI version", "Run time" };
            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cell(1, i + 1).Value = headers[i];
                ws.Cell(1, i + 1).Style.Font.Bold = true;
            }

            int r = 2;
            foreach (var f in _findings)
            {
                ws.Cell(r, 1).Value = f.Row.ExcelRow;
                ws.Cell(r, 2).Value = f.Row.Channel > 0 ? f.Row.Channel.ToString() : "System";
                ws.Cell(r, 3).Value = f.Row.Name;
                ws.Cell(r, 4).Value = f.Input;
                ws.Cell(r, 5).Value = f.Expected;
                ws.Cell(r, 6).Value = f.Actual;
                ws.Cell(r, 7).Value = f.GuiLog.Length > 200 ? f.GuiLog.Substring(0, 200) + "..." : f.GuiLog;
                ws.Cell(r, 8).Value = _unit.Name;
                ws.Cell(r, 9).Value = _guiVersion;
                ws.Cell(r, 10).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                r++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
            return path;
        }
        catch (Exception ex)
        {
            _log("WARN  | could not write the failures file: " + ex.Message);
            return null;
        }
    }
}
