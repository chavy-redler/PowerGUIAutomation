using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Threading;
using ClosedXML.Excel;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;

namespace PowerGUIAutomation.Tests;

// Compares every parameter of the unit's Excel sheet ("<PN> Parameters": name, value id, default) with what the Configure tab shows.
//
// FAST PAGE SCAN (default path): the Configure tree has one node per scope - "System Parameters" and "Channel N Parameters". Selecting
// the node shows that scope's whole parameter list (a virtualized list, so it is scrolled in half-page steps). Every visible row whose
// AutomationId is wanted is compared on the spot. Which channel is compared is therefore certain: it is the tree node that was
// selected, and the GUI's own breadcrumb ("Channel Parameters > Channel 12 Parameters") is logged with every result.
//
// SEARCH FALLBACK: a row that was not seen by the page scan is looked up with the Configure search box ("N of M matches", "Next Result")
// until the breadcrumb shows the wanted channel. If the wanted channel cannot be reached the test FAILS - it never compares another
// channel's value.
//
// Stops on the first mismatch, same as every other test in this project.
public class ConfigureCompareParametersToExcelTest
{
    private const int MaxFindRetries = 12;
    private const int RetryDelayMs = 250;

    private readonly string _excelPath;
    private readonly Action<string> _log;
    private readonly string[] _sheetNames;
    private readonly UnitProfile? _unit;
    private readonly Func<int, string, bool>? _rowFilter;     // debug runs: (channel, parameter name) -> run this row?
    private static readonly Regex BracketedRegex = new(@"^\[\s*([^,\]]+?)\s*[,\]]");
    private static readonly Regex FetchingRegex = new(@"Fetching parameters\.\.\.\s*(\d+)\s*/\s*(\d+)", RegexOptions.Compiled);

    // the row also shows the parameter's data type as a plain text ("String") - that is a label, not the value
    private static readonly HashSet<string> DataTypeTokens = new(StringComparer.OrdinalIgnoreCase) { "String", "Integer", "Float", "Number", "Boolean" };

    private int _pass, _fail, _skip;

    // "Channel N Group Control" rows are listed on their own tree page "Group Channel Control" (all channels together), not under the channel
    private static readonly Regex GroupRowRegex = new(@"^Channel \d+ Group Control$", RegexOptions.Compiled);
    private const int GroupScope = -1;
    private static bool IsGroupRow(Row r) => r.Channel > 0 && GroupRowRegex.IsMatch(r.ExpectedName);

    private sealed record Row(string ExcelRow, string ParameterPath, string ExpectedName, string ValuePath, string ExpectedValue, int Channel);

    public ConfigureCompareParametersToExcelTest(string excelPath, Action<string> log, string[]? sheetNames = null, UnitProfile? unit = null, Func<int, string, bool>? rowFilter = null)
    {
        _rowFilter = rowFilter;
        _excelPath = excelPath;
        _log = log;
        _unit = unit;
        _sheetNames = sheetNames ?? new[] { "System Parameters", "Channel Parameters" };
    }

    public bool Run()
    {
        using var automation = new UIA3Automation();
        var mainWindow = AppConnection.Attach(automation);

        if (_unit != null)
        {
            string? unitProblem = UnitCheck.Verify(mainWindow, _unit, _log);
            if (unitProblem != null)
            {
                _log("FAIL  | (test setup) | STEP: check the connected unit | WHY: " + unitProblem + " | CHECK: connect the unit '" + _unit.Name + "' or choose the right unit (U in the main menu)");
                return false;
            }
        }

        AppConnection.GoToTab(mainWindow, "ConfigureTab");

        List<Row> dataRows;
        try
        {
            dataRows = ReadTestData();
        }
        catch (Exception ex)
        {
            _log("FAIL | Could not read test data from Excel: " + ex.Message);
            return false;
        }

        if (_rowFilter != null)
        {
            dataRows = dataRows.Where(r => _rowFilter(r.Channel, r.ExpectedName)).ToList();
            _log("INFO  | debug filter active: " + dataRows.Count + " rows will be checked");
            if (dataRows.Count == 0)
            {
                _log("FAIL  | the filter matched no rows");
                return false;
            }
        }

        if (!WaitFetchDone(mainWindow, 300000))
        {
            _log("FAIL  | (test setup) | STEP: open the Configure tab | WHY: the GUI kept 'Fetching parameters...' for more than 5 minutes | CHECK: is the unit connected and responding?");
            return false;
        }

        // rows without an id cannot be looked up
        foreach (var r in dataRows.Where(r => string.IsNullOrWhiteSpace(r.ParameterPath)))
        {
            _log("SKIP  | Row " + r.ExcelRow + " | " + r.ExpectedName + " | no RxPath defined");
            _skip++;
        }

        var runnable = dataRows.Where(r => !string.IsNullOrWhiteSpace(r.ParameterPath)).ToList();
        var sw = Stopwatch.StartNew();

        foreach (var scope in runnable.GroupBy(r => IsGroupRow(r) ? GroupScope : r.Channel))
        {
            string scopeName = scope.Key == GroupScope ? "Group Channel Control" : scope.Key > 0 ? "Channel " + scope.Key : "System";
            var pending = scope.ToList();
            _log("INFO  | scope " + scopeName + ": " + pending.Count + " parameters to check (page scan)");

            string? crumb = SelectScope(mainWindow, scope.Key);
            if (crumb == null)
            {
                _log("FAIL  | " + scopeName + " | the tree node '" + (scope.Key == GroupScope ? "Group Channel Control" : scope.Key > 0 ? "Channel " + scope.Key + " Parameters" : "System Parameters") + "' could not be selected - breadcrumb now: '" + ReadBreadcrumb(mainWindow) + "'");
                _log("STOPPED | halted after first mismatch - remaining parameters were not checked");
                return false;
            }

            if (!ScanPages(mainWindow, scopeName, crumb, pending))
            {
                _log("STOPPED | halted after first mismatch - remaining parameters were not checked");
                return false;
            }

            // anything the page scan did not see: look it up with the search box (channel verified by the breadcrumb)
            foreach (var row in pending.ToList())
            {
                _log("INFO  | Row " + row.ExcelRow + " | " + row.ExpectedName + " | not seen in the page scan - using the search box");
                if (!CheckBySearch(mainWindow, row))
                {
                    _log("STOPPED | halted after first mismatch - remaining parameters were not checked");
                    return false;
                }
            }
        }

        ClearSearch(mainWindow);
        _log("DONE | pass=" + _pass + " fail=" + _fail + " skip=" + _skip + " | " + sw.Elapsed.TotalSeconds.ToString("0") + " s");

        return _fail == 0;
    }

    // ------------------------------------------------------------------ page scan

    private static AutomationElement[] TreeItems(Window w) => w.FindAllDescendants(cf => cf.ByControlType(ControlType.TreeItem));

    // Selects "System Parameters" / "Channel N Parameters" in the Configure tree and waits until the GUI's breadcrumb says so.
    // Returns the breadcrumb text, or null when it did not work.
    private string? SelectScope(Window w, int channel)
    {
        string nodeName = channel == GroupScope ? "Group Channel Control" : channel > 0 ? "Channel " + channel + " Parameters" : "System Parameters";
        var node = TreeItems(w).FirstOrDefault(x => x.Name == nodeName);
        if (node == null && channel > 0)   // (GroupScope is a top-level node, nothing to expand)
        {
            TreeItems(w).FirstOrDefault(x => x.Name == "Channel Parameters")?.Patterns.ExpandCollapse.PatternOrDefault?.Expand();
            Thread.Sleep(600);
            node = TreeItems(w).FirstOrDefault(x => x.Name == nodeName);
        }

        if (node == null)
        {
            // the tree is virtualized too: scroll it until the node is realized
            var tree = w.FindFirstDescendant(cf => cf.ByAutomationId("Parameters_TreeView"));
            var scroll = tree?.Patterns.Scroll.PatternOrDefault;
            if (scroll != null)
            {
                for (double pct = 0; pct <= 100 && node == null; pct += 10)
                {
                    scroll.SetScrollPercent(-1, pct);
                    Thread.Sleep(250);
                    node = TreeItems(w).FirstOrDefault(x => x.Name == nodeName);
                }
            }
        }

        if (node == null)
        {
            return null;
        }

        node.Patterns.SelectionItem.PatternOrDefault?.Select();
        var swWait = Stopwatch.StartNew();
        while (swWait.ElapsedMilliseconds < 6000)
        {
            string crumb = ReadBreadcrumb(w);
            bool ok = channel == GroupScope
                ? crumb.Contains("Group Channel Control", StringComparison.Ordinal)
                : channel > 0
                    ? crumb.Contains("Channel " + channel + " Parameters", StringComparison.Ordinal)
                    : crumb.Contains("System Parameters", StringComparison.Ordinal) && !crumb.Contains("Channel", StringComparison.Ordinal);
            if (ok)
            {
                Thread.Sleep(300);       // let the list rebuild after the selection
                return crumb;
            }

            Thread.Sleep(150);
        }

        return null;
    }

    // Scrolls the scope's parameter list in half-page steps (virtualized: only the visible rows exist) and compares every wanted row
    // as soon as it is visible. Rows left in `pending` were not seen.
    private bool ScanPages(Window w, string scopeName, string crumb, List<Row> pending)
    {
        // The GUI may rebuild the list after a tree selection, so the list element is looked up again on every page (a stale one shows nothing).
        var list = w.FindFirstDescendant(cf => cf.ByAutomationId("ParametersList"));
        var scroll = list?.Patterns.Scroll.PatternOrDefault;
        if (list == null)
        {
            _log("WARN  | " + scopeName + " | the parameter list was not found - falling back to the search box");
            return true;
        }

        double view = scroll != null && scroll.VerticallyScrollable.ValueOrDefault ? Math.Max(5, scroll.VerticalViewSize.ValueOrDefault) : 100;
        string where = " | " + scopeName + " | breadcrumb='" + crumb + "' | page scan";
        for (double pct = 0; pending.Count > 0; pct += view / 2)
        {
            list = w.FindFirstDescendant(cf => cf.ByAutomationId("ParametersList")) ?? list;
            scroll = list.Patterns.Scroll.PatternOrDefault;
            if (scroll != null && scroll.VerticallyScrollable.ValueOrDefault)
            {
                scroll.SetScrollPercent(-1, Math.Min(pct, 100));
                Thread.Sleep(150);
            }

            var visible = list.FindAllDescendants().Select(e => e.AutomationId ?? "").Where(id => id.Length > 0).ToHashSet();
            if (!visible.Any(id => id.StartsWith("Param")))
            {
                // the list is still being (re)built - give it time once and look again
                Thread.Sleep(1200);
                list = w.FindFirstDescendant(cf => cf.ByAutomationId("ParametersList")) ?? list;
                visible = list.FindAllDescendants().Select(e => e.AutomationId ?? "").Where(id => id.Length > 0).ToHashSet();
            }

            foreach (var row in pending.ToList())
            {
                string nameId = ExtractAutomationId(row.ParameterPath)!;
                string? valueId = ExtractAutomationId(row.ValuePath);
                if (!visible.Contains(nameId) || (valueId != null && !visible.Contains(valueId)))
                {
                    continue;
                }

                var nameElement = list.FindFirstDescendant(cf => cf.ByAutomationId(nameId));
                if (nameElement == null)
                {
                    continue;       // scrolled away between the two calls - seen again on the next page
                }

                pending.Remove(row);
                try
                {
                    if (!CompareRow(w, row, nameElement, where))
                    {
                        return false;
                    }
                }
                catch (Exception ex) when (ex is FlaUI.Core.Exceptions.PropertyNotSupportedException or FlaUI.Core.Exceptions.ElementNotAvailableException or System.Runtime.InteropServices.COMException)
                {
                    // the virtualized list recycled the row while it was being read - read it again on a later page (or by search)
                    _log("NOTE  | Row " + row.ExcelRow + " | the row changed while it was read (" + ex.GetType().Name + ") - will read it again");
                    pending.Add(row);
                }
            }

            if (scroll == null || pct >= 100)
            {
                break;
            }
        }

        return true;
    }

    // ------------------------------------------------------------------ search fallback

    private static void ClearSearch(Window w)
    {
        try
        {
            var x = w.FindFirstDescendant(cf => cf.ByHelpText("Clear Search"));
            if (x != null && x.IsEnabled)
            {
                x.AsButton().Invoke();
                Thread.Sleep(300);
            }
        }
        catch
        {
            // best effort
        }
    }

    private bool CheckBySearch(Window w, Row row)
    {
        string? nameAutomationId = ExtractAutomationId(row.ParameterPath);
        if (nameAutomationId == null)
        {
            _log("FAIL  | Row " + row.ExcelRow + " | could not extract automationid from ParameterPath: " + row.ParameterPath);
            return false;
        }

        var searchBox = w.FindFirstDescendant(cf => cf.ByAutomationId("searchTextBox"))?.AsTextBox();
        if (searchBox == null)
        {
            _log("FAIL  | could not find searchTextBox - is the Configure tab open?");
            return false;
        }

        // Bring the row into view via the app's own search box - paste (Ctrl+V) the whole string in one shot rather than simulating
        // keystrokes, so it isn't corrupted by whatever OS keyboard layout happens to be active.
        ClearSearch(w);
        searchBox.Focus();
        searchBox.Text = "";
        TextCopy.ClipboardService.SetText(row.ExpectedName);
        Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_V);

        var nameElement = FindWithRetry(w, nameAutomationId);
        string where = "";       // what the GUI itself says is displayed (breadcrumb + result counter) - proof of WHICH channel was compared
        if (nameElement != null && IsGroupRow(row))
        {
            string crumb = ReadBreadcrumb(w);
            if (!crumb.Contains("Group Channel Control", StringComparison.Ordinal))
            {
                _log("FAIL  | Row " + row.ExcelRow + " | Expected='" + row.ExpectedName + "' | the GUI shows '" + crumb + "' instead of the 'Group Channel Control' page");
                return false;
            }

            where = " | Group Channel Control | breadcrumb='" + crumb + "' " + ReadCounter(w);
        }
        else if (nameElement != null && row.Channel > 0)
        {
            var (channelElement, crumb, counter) = GoToChannel(w, nameAutomationId, row.Channel);
            if (channelElement == null)
            {
                // never fall back to the first result (another channel): that would compare the wrong channel's value
                _log("FAIL  | Row " + row.ExcelRow + " | Expected='" + row.ExpectedName + "' | could not reach 'Channel " + row.Channel + " Parameters' - breadcrumb of the last result: '" + crumb + "' " + counter);
                return false;
            }

            nameElement = channelElement;
            where = " | Ch" + row.Channel + " | breadcrumb='" + crumb + "' " + counter;
        }
        else if (nameElement != null)
        {
            string crumb = ReadBreadcrumb(w);
            if (Regex.IsMatch(crumb, @"Channel \d+ Parameters"))
            {
                _log("FAIL  | Row " + row.ExcelRow + " | Expected='" + row.ExpectedName + "' | a SYSTEM parameter was expected but the GUI shows a channel result: '" + crumb + "'");
                return false;
            }

            where = " | System | breadcrumb='" + crumb + "' " + ReadCounter(w);
        }

        if (nameElement == null)
        {
            _log("FAIL  | Row " + row.ExcelRow + " | Expected='" + row.ExpectedName + "' | name element not found after " + MaxFindRetries + " attempts (automationid=" + nameAutomationId + ")");
            return false;
        }

        return CompareRow(w, row, nameElement, where + " | search");
    }

    // ------------------------------------------------------------------ compare one row

    private bool CompareRow(Window mainWindow, Row row, AutomationElement nameElement, string where)
    {
        string excelRow = row.ExcelRow, expectedName = row.ExpectedName, expectedValue = row.ExpectedValue;
        _log("START | Row " + excelRow + " | Expected='" + expectedName + "'");

        string actualName = (nameElement.Name ?? "").Trim();
        bool namePassed = string.Equals(actualName, expectedName, StringComparison.Ordinal);
        _log((namePassed ? "PASS  | " : "FAIL  | ") + "Row " + excelRow + " | Expected='" + expectedName + "' | Actual='" + actualName + "'" + where);
        if (namePassed)
        {
            _pass++;
        }
        else
        {
            _fail++;
            return false;
        }

        string? valueAutomationId = ExtractAutomationId(row.ValuePath);
        if (valueAutomationId == null)
        {
            _log("SKIP  | Row " + excelRow + " | no Value automationid defined");
            _skip++;
            return true;
        }

        string actualValue = "<NOT FOUND>";
        var valueElement = FindWithRetry(mainWindow, valueAutomationId);
        if (valueElement != null)
        {
            // Finding the element doesn't mean its ValuePattern has caught up yet - right after Focus() it can still read back empty
            // for a beat. Retry the read itself, not just the find, until we get something non-empty or run out of attempts.
            for (int attempt = 1; attempt <= MaxFindRetries; attempt++)
            {
                string candidate = "";
                if (valueElement.Patterns.Value.TryGetPattern(out var valuePattern))
                {
                    candidate = (valuePattern.Value.Value ?? "").Trim();
                }
                else if (valueElement.ControlType == ControlType.ComboBox)
                {
                    var comboBox = valueElement.AsComboBox();
                    candidate = (comboBox.SelectedItem?.Text ?? "").Trim();
                }

                // A ComboBox can answer ValuePattern with "" although an item IS selected (the selection is then only in
                // SelectedItem or in the text shown inside the control) - try those before giving up.
                if (candidate.Length == 0 && valueElement.ControlType == ControlType.ComboBox)
                {
                    var selected = valueElement.AsComboBox().SelectedItem;
                    candidate = (selected?.Name ?? "").Trim();
                    if (candidate.Length == 0)
                    {
                        candidate = (selected?.Text ?? "").Trim();
                    }

                    if (candidate.Length == 0)
                    {
                        candidate = RawTexts(valueElement).FirstOrDefault(t => t.Length > 0) ?? "";
                    }
                }

                if (candidate.Length > 0)
                {
                    actualValue = candidate;
                    break;
                }

                // An editor often reads empty without focus although the row SHOWS the value as a separate text (that text is what the user
                // sees): use it right away instead of waiting for retries.
                if (valueElement.ControlType != ControlType.ComboBox)
                {
                    string shown = ReadDisplayedText(nameElement, valueElement);
                    if (shown.Length > 0)
                    {
                        actualValue = shown;
                        _log("NOTE  | Row " + excelRow + " | the value control reads empty - using the text DISPLAYED in the row: '" + shown + "'");
                        break;
                    }

                    // otherwise a text box needs Focus() before its value can be read (a ComboBox is read as it is: a click into it can blank its text)
                    try
                    {
                        valueElement.Focus();
                    }
                    catch (Exception ex) when (ex is InvalidOperationException or System.Runtime.InteropServices.COMException)
                    {
                        // not focusable right now (partly visible row / inactive window) - the next attempt reads again
                    }
                }

                if (attempt == 1)
                {
                    actualValue = ""; // remember we did get a (empty) read, in case every attempt stays empty
                }

                Thread.Sleep(RetryDelayMs);
            }
        }

        // Logic-expression parameters ("Digital Output 1 Logic Control", ...) have an EDITOR text box that is empty by design
        // (clicking into it shows an empty field) and show the current expression in a separate Text of the same row. The user
        // sees that Text, so it is the value to compare.
        if (valueElement != null && actualValue.Length == 0)
        {
            string displayed = ReadDisplayedText(nameElement, valueElement);
            if (displayed.Length > 0)
            {
                actualValue = displayed;
                _log("NOTE  | Row " + excelRow + " | the value control is an empty editor - using the text DISPLAYED in the row: '" + displayed + "'");
            }
        }

        if (valueElement != null && actualValue.Length == 0)
        {
            _log("NOTE  | Row " + excelRow + " | the value control reads EMPTY: control=" + valueElement.ControlType + " class=" + valueElement.ClassName
                 + " | ValuePattern=" + (valueElement.Patterns.Value.IsSupported ? "supported (empty)" : "not supported")
                 + " | the GUI shows no value for '" + expectedName + "' (Excel expects '" + expectedValue + "')");
        }

        // Some ComboBox-backed parameters expose their raw bound object instead of the displayed text - "[250, 3]" or "[Disable, 0]"
        // rather than plain "250"/"Disable". Take the first element inside the brackets (numeric or text) for comparison.
        string comparisonValue = actualValue;
        var bracketedMatch = BracketedRegex.Match(actualValue);
        if (bracketedMatch.Success)
        {
            comparisonValue = bracketedMatch.Groups[1].Value;
            _log("NOTE  | Row " + excelRow + " | Actual value came back bracketed ('" + actualValue + "') - comparing against first element '" + comparisonValue + "'");
        }

        bool valuePassed;
        if (double.TryParse(expectedValue, out double expectedNumber) && double.TryParse(comparisonValue, out double actualNumber))
        {
            valuePassed = expectedNumber == actualNumber;
        }
        else
        {
            valuePassed = string.Equals(comparisonValue, expectedValue, StringComparison.Ordinal);
        }

        _log((valuePassed ? "PASS  | " : "FAIL  | ") + "Row " + excelRow + " | Value Expected='" + expectedValue + "' | Actual='" + actualValue + "'" + where);
        if (valuePassed)
        {
            _pass++;
            return true;
        }

        _fail++;
        return false;
    }

    // ------------------------------------------------------------------ status bar

    // After opening the Configure tab the GUI keeps loading the unit's parameters ("Fetching parameters... 5/688"): values are blank until done.
    // The status text can also stay stuck on a number although the GUI is idle: same counter for 8 s with a usable search box = finished.
    private bool WaitFetchDone(Window window, int timeoutMs)
    {
        var sw = Stopwatch.StartNew();
        long? absentSince = null;
        string? last = null;
        long since = 0;
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            string? progress = null;
            try
            {
                foreach (var t in window.FindAllDescendants(cf => cf.ByControlType(ControlType.Text)))
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
                absentSince ??= sw.ElapsedMilliseconds;
                if (sw.ElapsedMilliseconds - absentSince >= 1000)
                {
                    return true;
                }

                Thread.Sleep(300);
                continue;
            }

            absentSince = null;
            if (progress != last)
            {
                last = progress;
                since = sw.ElapsedMilliseconds;
            }
            else if (sw.ElapsedMilliseconds - since >= 8000 && window.FindFirstDescendant(cf => cf.ByAutomationId("searchTextBox")) is { IsEnabled: true })
            {
                _log("WARN  | the status bar is stuck on 'Fetching parameters... " + progress + "' but the search box is enabled - treating the load as finished");
                return true;
            }

            Thread.Sleep(1000);
        }

        return false;
    }

    // ------------------------------------------------------------------ helpers

    // The visible Text of the parameter's row that is neither the parameter name nor a unit/description: what the user reads as the value.
    private static string ReadDisplayedText(AutomationElement nameElement, AutomationElement valueElement)
    {
        try
        {
            string paramName = (nameElement.Name ?? "").Trim();
            var row = valueElement.Parent;
            if (row == null)
            {
                return "";
            }

            // only texts that are really visible: the raw tree also holds hidden template texts (unit token "String", tooltip text, a hidden copy of the value)
            return VisibleTexts(row, 0)
                .Where(t => t.Length > 0 && t.Length < 80 && t != paramName && !DataTypeTokens.Contains(t) && !t.StartsWith("[") && !Regex.IsMatch(t, @"^\d+ / \d+$"))
                .FirstOrDefault() ?? "";
        }
        catch
        {
            return "";
        }
    }

    private static IEnumerable<string> VisibleTexts(AutomationElement root, int depth)
    {
        var walker = root.Automation.TreeWalkerFactory.GetRawViewWalker();
        for (var c = walker.GetFirstChild(root); c != null; c = walker.GetNextSibling(c))
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

    // Texts inside a control, read from the raw UIA tree (the displayed value of a ComboBox is such a child Text).
    private static IEnumerable<string> RawTexts(AutomationElement root, int depth = 0)
    {
        var walker = root.Automation.TreeWalkerFactory.GetRawViewWalker();
        for (var c = walker.GetFirstChild(root); c != null; c = walker.GetNextSibling(c))
        {
            if (c.ControlType == ControlType.Text)
            {
                yield return (c.Name ?? "").Trim();
            }

            if (depth < 6)
            {
                foreach (var t in RawTexts(c, depth + 1))
                {
                    yield return t;
                }
            }
        }
    }

    private static AutomationElement? FindWithRetry(FlaUI.Core.AutomationElements.Window window, string automationId)
    {
        for (int attempt = 1; attempt <= MaxFindRetries; attempt++)
        {
            var element = window.FindFirstDescendant(cf => cf.ByAutomationId(automationId));
            if (element != null)
            {
                return element;
            }
            Thread.Sleep(RetryDelayMs);
        }
        return null;
    }

    // The id columns of the workbook hold the plain AutomationId - used exactly as written.
    private static string? ExtractAutomationId(string? cell)
        => string.IsNullOrWhiteSpace(cell) ? null : cell.Trim();

    private static string ReadBreadcrumb(FlaUI.Core.AutomationElements.Window window)
    {
        var crumb = window.FindFirstDescendant(cf => cf.ByAutomationId("BreadCrumb"));
        return crumb == null ? "" : string.Join("", RawTexts(crumb, 0));
    }

    // "3 of 16 matches" - best effort (the counter is a Text without an id); "" when it cannot be found
    private static string ReadCounter(FlaUI.Core.AutomationElements.Window window)
    {
        try
        {
            foreach (var t in window.FindAllDescendants(cf => cf.ByControlType(ControlType.Text)))
            {
                var m = Regex.Match(t.Name ?? "", @"^(\d+) of (\d+) matches");
                if (m.Success)
                {
                    return "| result " + m.Groups[1].Value + " of " + m.Groups[2].Value;
                }
            }
        }
        catch
        {
            // best effort
        }

        return "";
    }

    // Shows result N of the search for a channel parameter: presses "Next Result" until the breadcrumb says "Channel N Parameters".
    // Returns the element only when the GUI itself says that channel is displayed; otherwise null (+ the last breadcrumb for the message).
    private static (AutomationElement? element, string crumb, string counter) GoToChannel(FlaUI.Core.AutomationElements.Window window, string nameAutomationId, int channel)
    {
        string wanted = "Channel " + channel + " Parameters";
        string text = "";
        for (int step = 0; step < 20; step++)
        {
            text = ReadBreadcrumb(window);
            if (text.Contains(wanted, StringComparison.Ordinal))
            {
                return (FindWithRetry(window, nameAutomationId), text, ReadCounter(window));
            }

            var next = window.FindFirstDescendant(cf => cf.ByHelpText("Next Result"));
            if (next == null || !next.IsEnabled)
            {
                return (null, text, ReadCounter(window));
            }

            next.AsButton().Invoke();
            Thread.Sleep(450);
        }

        return (null, text, ReadCounter(window));
    }

    private List<Row> ReadTestData()
    {
        using var workbook = new XLWorkbook(_excelPath);
        var result = new List<Row>();

        foreach (var sheetName in _sheetNames)
        {
            if (!workbook.Worksheets.Contains(sheetName))
            {
                continue;
            }

            var ws = workbook.Worksheet(sheetName);
            foreach (var row in ws.RowsUsed().Skip(1)) // skip header
            {
                string scope = row.Cell(1).GetString().Trim();        // "System" or the channel number
                int.TryParse(scope, out int channelNumber);
                result.Add(new Row(
                    (channelNumber > 0 ? "Ch" + channelNumber + " r" : "r") + row.RowNumber(),
                    row.Cell(2).GetString().Trim(),
                    row.Cell(3).GetString().Trim(),
                    row.Cell(4).GetString().Trim(),
                    row.Cell(5).GetString().Trim(),
                    channelNumber
                ));
            }
        }

        return result;
    }
}
