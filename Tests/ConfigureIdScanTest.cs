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

// READ-ONLY scan of the Configure tab: for every parameter of the unit's Excel sheet it brings the row into view (search box),
// finds the row's NAME element and VALUE control as the user sees them, and reports what AutomationId each really has in the
// running GUI - compared with the id the Excel expects. Writes nothing to the unit and nothing to the master workbook; the report
// goes to a new workbook "Excel\Configure_AutomationId_Scan_<PN>_<time>.xlsx".
//
// Status per parameter: OK / NAME-ID MISSING / NAME-ID DIFFERENT / VALUE-ID MISSING / VALUE-ID DIFFERENT / VALUE-ID SHARED /
// NOT FOUND (the row never appeared in the search results).
public class ConfigureIdScanTest
{
    private const int FindRetries = 10;
    private const int RetryDelayMs = 350;
    private const int MaxResultsToStep = 25;

    private static readonly Regex MatchesRegex = new(@"^(\d+) of (\d+) matches", RegexOptions.Compiled);
    private static readonly Regex FetchingRegex = new(@"Fetching parameters\.\.\.\s*(\d+)\s*/\s*(\d+)", RegexOptions.Compiled);

    private readonly string _excelPath;
    private readonly Action<string> _log;
    private readonly UnitProfile _unit;

    private sealed class Row
    {
        public string Scope = "";        // "System" or channel number as in the sheet
        public string Name = "";
        public string ExpNameId = "";
        public string ExpValueId = "";
        public int ExcelRow;
    }

    private sealed class Result
    {
        public Row Param = null!;
        public string Status = "";
        public string Why = "";
        public string Breadcrumb = "";
        public string ActNameId = "";
        public string NameType = "";
        public int NameIdCount;
        public string ActValueId = "";   // id(s) of the value control(s) found in the row
        public string ValueType = "";
        public string ValueClass = "";
        public int ValueIdCount;
    }

    public ConfigureIdScanTest(string excelPath, Action<string> log, UnitProfile unit)
    {
        _excelPath = excelPath;
        _log = log;
        _unit = unit;
    }

    public bool Run()
    {
        using var automation = new UIA3Automation();
        var window = AppConnection.Attach(automation);

        string? unitProblem = UnitCheck.Verify(window, _unit, _log);
        if (unitProblem != null)
        {
            _log("FAIL  | (test setup) | STEP: check the connected unit | WHY: " + unitProblem + " | CHECK: connect '" + _unit.Name + "' or choose the right unit (U in the main menu)");
            return false;
        }

        List<Row> rows;
        try
        {
            rows = ReadRows();
        }
        catch (Exception ex)
        {
            _log("FAIL  | (test setup) | STEP: read the Excel sheet '" + _unit.ParameterSheet + "' | WHY: " + ex.Message + " | CHECK: does the sheet exist?");
            return false;
        }

        _log("INFO  | " + rows.Count + " unique parameter names to scan (sheet '" + _unit.ParameterSheet + "')");

        AppConnection.GoToTab(window, "ConfigureTab");
        if (!WaitFetchDone(window, 300000))
        {
            _log("FAIL  | (test setup) | STEP: open the Configure tab | WHY: the GUI kept 'Fetching parameters...' for more than 5 minutes | CHECK: is the unit connected and responding?");
            return false;
        }

        var results = new List<Result>();
        int n = 0;
        foreach (var row in rows)
        {
            n++;
            _log("DRAW  | [" + n + "/" + rows.Count + "] " + row.Name);
            Result r;
            try
            {
                r = Scan(window, row);
            }
            catch (Exception ex)
            {
                r = new Result { Param = row, Status = "ERROR", Why = ex.GetType().Name + ": " + ex.Message };
            }

            results.Add(r);
            _log((r.Status == "OK" ? "PASS  | " : "WARN  | ") + row.Name + " | " + r.Status + (r.Why.Length > 0 ? " | " + r.Why : "")
                 + " | name id='" + r.ActNameId + "' value id='" + r.ActValueId + "'");
        }

        ClearSearch(window);
        string report = WriteReport(results);

        var bad = results.Where(x => x.Status != "OK").ToList();
        _log("RESULT| ===== AutomationId scan of " + _unit.PartNumber + " =====");
        foreach (var g in results.GroupBy(x => x.Status).OrderByDescending(g => g.Count()))
        {
            _log("RESULT| " + g.Key.PadRight(22) + " " + g.Count());
        }

        _log("RESULT| report: " + report);
        _log(bad.Count == 0 ? "RESULT| PASSED - every parameter has the expected AutomationIds" : "RESULT| FAILED - " + bad.Count + " of " + results.Count + " parameters differ from the Excel (see report)");
        return bad.Count == 0;
    }

    // ------------------------------------------------------------------ scan of one parameter

    private Result Scan(Window window, Row p)
    {
        var res = new Result { Param = p };

        ClearSearch(window);
        if (!WaitFetchDone(window, 300000))
        {
            res.Status = "ERROR";
            res.Why = "the GUI kept 'Fetching parameters...' for more than 5 minutes";
            return res;
        }

        var search = FreshSearch(window);
        if (search == null || !PrepareSearch(search))
        {
            res.Status = "ERROR";
            res.Why = "the search box could not be used";
            return res;
        }

        TextCopy.ClipboardService.SetText(p.Name);
        Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_V);

        // The search shows ONE result at a time (substring match); step through the results until the exact parameter name is on screen.
        AutomationElement? nameEl = null;
        AutomationElement? next = null;
        for (int step = 0; step < MaxResultsToStep && nameEl == null; step++)
        {
            for (int attempt = 0; attempt < FindRetries && nameEl == null; attempt++)
            {
                nameEl = FindByExactName(window, p.Name);
                if (nameEl == null)
                {
                    Thread.Sleep(RetryDelayMs);
                }
            }

            if (nameEl != null)
            {
                break;
            }

            var counter = Counter(window);
            if (counter == null || counter.Value.index >= counter.Value.total)
            {
                break;
            }

            next ??= window.FindFirstDescendant(cf => cf.ByHelpText("Next Result"));
            if (next == null || !next.IsEnabled)
            {
                break;
            }

            next.AsButton().Invoke();
            Thread.Sleep(700);
        }

        if (nameEl == null)
        {
            res.Status = "NOT FOUND";
            res.Why = "no element with the text '" + p.Name + "' appeared in the search results";
            return res;
        }

        res.Breadcrumb = Breadcrumb(window);

        // NAME element: what id does the text the user reads really have?
        res.ActNameId = (nameEl.AutomationId ?? "").Trim();
        res.NameType = nameEl.ControlType.ToString();
        res.NameIdCount = res.ActNameId.Length == 0 ? 0 : window.FindAllDescendants(cf => cf.ByAutomationId(res.ActNameId)).Length;

        // VALUE control: the Excel id first; otherwise the first row ancestor that holds an input control.
        var valueEl = window.FindFirstDescendant(cf => cf.ByAutomationId(p.ExpValueId));
        var rowControls = FindRowControls(nameEl);
        if (valueEl != null)
        {
            res.ActValueId = p.ExpValueId;
            res.ValueType = valueEl.ControlType.ToString();
            res.ValueClass = valueEl.ClassName ?? "";
            res.ValueIdCount = window.FindAllDescendants(cf => cf.ByAutomationId(p.ExpValueId)).Length;
        }
        else if (rowControls.Count > 0)
        {
            res.ActValueId = string.Join(" ; ", rowControls.Select(c => c.AutomationId ?? "").Distinct());
            res.ValueType = string.Join(" ; ", rowControls.Select(c => c.ControlType.ToString()).Distinct());
            res.ValueClass = string.Join(" ; ", rowControls.Select(c => c.ClassName ?? "").Distinct());
        }

        var problems = new List<string>();
        string status = "OK";
        if (res.ActNameId.Length == 0)
        {
            status = "NAME-ID MISSING";
            problems.Add("the parameter-name text has NO AutomationId (Excel expects '" + p.ExpNameId + "')");
        }
        else if (!string.Equals(res.ActNameId, p.ExpNameId, StringComparison.Ordinal))
        {
            status = "NAME-ID DIFFERENT";
            problems.Add("the name text has id '" + res.ActNameId + "' but Excel expects '" + p.ExpNameId + "'");
        }

        if (valueEl == null)
        {
            string actual = rowControls.Count == 0
                ? "no input control was found in the row"
                : "the row's control(s) have id '" + res.ActValueId + "' (" + res.ValueType + ")";
            if (status == "OK")
            {
                status = res.ActValueId.Length == 0 || res.ActValueId.Replace(";", "").Trim().Length == 0 ? "VALUE-ID MISSING" : "VALUE-ID DIFFERENT";
            }

            problems.Add("no control with the Excel value id '" + p.ExpValueId + "' - " + actual);
        }
        else if (res.ValueIdCount > 1)
        {
            if (status == "OK")
            {
                status = "VALUE-ID SHARED";
            }

            problems.Add(res.ValueIdCount + " controls share the value id '" + p.ExpValueId + "' on screen (e.g. an editor and a hidden ComboBox) - UIA cannot tell them apart");
        }

        res.Status = status;
        res.Why = string.Join(" | ", problems);
        return res;
    }

    // Input controls that sit in the same row as the name text: climb a few ancestors until one holds a control, but stop before
    // the container that holds several parameters.
    private static List<AutomationElement> FindRowControls(AutomationElement nameEl)
    {
        var types = new[] { ControlType.Edit, ControlType.ComboBox, ControlType.CheckBox, ControlType.Spinner, ControlType.Slider, ControlType.RadioButton };
        var found = new List<AutomationElement>();
        try
        {
            var walker = nameEl.Automation.TreeWalkerFactory.GetRawViewWalker();
            var node = walker.GetParent(nameEl);
            for (int up = 0; up < 8 && node != null; up++)
            {
                var controls = Walk(node, 0).Where(e => types.Contains(e.ControlType)).ToList();
                if (controls.Count > 0)
                {
                    // a container that holds more than ~4 controls is the list, not the row
                    return controls.Count <= 4 ? controls : found;
                }

                node = walker.GetParent(node);
            }
        }
        catch
        {
            // GUI re-created the row while reading
        }

        return found;
    }

    private static AutomationElement? FindByExactName(Window window, string name)
    {
        try
        {
            return window.FindFirstDescendant(cf => cf.ByName(name));
        }
        catch
        {
            return null;
        }
    }

    // ------------------------------------------------------------------ Excel in / out

    // One entry per UNIQUE parameter name (first row of the sheet that has it): the ids are the same in every channel.
    private List<Row> ReadRows()
    {
        string temp = Path.Combine(Path.GetTempPath(), "IdScan_" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.Copy(_excelPath, temp, true);       // works while the workbook is open in Excel
        try
        {
            using var wb = new XLWorkbook(temp);
            var ws = wb.Worksheet(_unit.ParameterSheet);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var rows = new List<Row>();
            foreach (var r in ws.RowsUsed().Skip(1))
            {
                string name = r.Cell(3).GetString().Trim();
                if (name.Length == 0 || !seen.Add(name))
                {
                    continue;
                }

                rows.Add(new Row
                {
                    Scope = r.Cell(1).GetString().Trim(),
                    Name = name,
                    ExpNameId = r.Cell(2).GetString().Trim(),
                    ExpValueId = r.Cell(4).GetString().Trim(),
                    ExcelRow = r.RowNumber(),
                });
            }

            return rows;
        }
        finally
        {
            try { File.Delete(temp); } catch { /* temp copy */ }
        }
    }

    private string WriteReport(List<Result> results)
    {
        string dir = Path.Combine(Path.GetDirectoryName(_excelPath)!, "Excel");
        Directory.CreateDirectory(dir);
        string path = Path.Combine(dir, "Configure_AutomationId_Scan_" + _unit.PartNumber + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx");

        using var wb = new XLWorkbook();

        var sum = wb.AddWorksheet("Summary");
        sum.Cell(1, 1).Value = "Unit"; sum.Cell(1, 2).Value = _unit.Name;
        sum.Cell(2, 1).Value = "Scanned"; sum.Cell(2, 2).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
        sum.Cell(3, 1).Value = "Sheet"; sum.Cell(3, 2).Value = _unit.ParameterSheet;
        int line = 5;
        sum.Cell(line, 1).Value = "Status"; sum.Cell(line, 2).Value = "Count"; sum.Row(line).Style.Font.Bold = true;
        foreach (var g in results.GroupBy(x => x.Status).OrderByDescending(g => g.Count()))
        {
            line++;
            sum.Cell(line, 1).Value = g.Key;
            sum.Cell(line, 2).Value = g.Count();
        }

        sum.Columns().AdjustToContents();

        var ws = wb.AddWorksheet("Scan");
        string[] head = { "Excel row", "Scope", "Parameter Name", "Status", "Why", "Expected Name ID", "Actual Name ID", "Name ctrl type",
                          "Name id count on screen", "Expected Value ID", "Actual Value ID (row)", "Value ctrl type", "Value class",
                          "Value id count on screen", "Breadcrumb" };
        for (int c = 0; c < head.Length; c++)
        {
            ws.Cell(1, c + 1).Value = head[c];
        }

        ws.Row(1).Style.Font.Bold = true;
        ws.Row(1).Style.Fill.BackgroundColor = XLColor.LightGray;
        int i = 1;
        foreach (var r in results.OrderBy(x => x.Status == "OK" ? 1 : 0).ThenBy(x => x.Param.ExcelRow))
        {
            i++;
            var v = new object[]
            {
                r.Param.ExcelRow, r.Param.Scope, r.Param.Name, r.Status, r.Why, r.Param.ExpNameId, r.ActNameId, r.NameType,
                r.NameIdCount, r.Param.ExpValueId, r.ActValueId, r.ValueType, r.ValueClass, r.ValueIdCount, r.Breadcrumb,
            };
            for (int c = 0; c < v.Length; c++)
            {
                ws.Cell(i, c + 1).Value = XLCellValue.FromObject(v[c]);
            }

            ws.Cell(i, 4).Style.Fill.BackgroundColor = r.Status == "OK" ? XLColor.LightGreen : r.Status == "NOT FOUND" || r.Status == "ERROR" ? XLColor.Orange : XLColor.LightSalmon;
        }

        ws.SheetView.FreezeRows(1);
        ws.Range(1, 1, i, head.Length).SetAutoFilter();
        ws.Columns().AdjustToContents(1, Math.Min(i, 60));
        ws.Column(5).Width = 70;
        ws.Column(5).Style.Alignment.WrapText = true;

        wb.SaveAs(path);
        return path;
    }

    // ------------------------------------------------------------------ UI helpers (same behaviour as ConfigureRandomChangeTest)

    private static TextBox? FreshSearch(Window window) => window.FindFirstDescendant(cf => cf.ByAutomationId("searchTextBox"))?.AsTextBox();

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

    private void ClearSearch(Window window)
    {
        try
        {
            var x = window.FindFirstDescendant(cf => cf.ByHelpText("Clear Search"));
            if (x != null && x.IsEnabled)
            {
                x.AsButton().Invoke();
                Thread.Sleep(800);
            }
        }
        catch (Exception ex)
        {
            _log("NOTE  | could not clear the search box: " + ex.Message);
        }
    }

    private bool WaitFetchDone(Window window, int timeoutMs)
    {
        var sw = Stopwatch.StartNew();
        long? absentSince = null;
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            bool fetching = false;
            try
            {
                fetching = Kids(window).Any(t => t.ControlType == ControlType.Text && FetchingRegex.IsMatch(t.Name ?? ""));
            }
            catch
            {
                // status bar being re-created
            }

            if (!fetching)
            {
                // the status text flickers while the GUI works - it must stay away for 3 s in a row
                absentSince ??= sw.ElapsedMilliseconds;
                if (sw.ElapsedMilliseconds - absentSince >= 3000)
                {
                    return true;
                }

                Thread.Sleep(500);
                continue;
            }

            absentSince = null;
            Thread.Sleep(1000);
        }

        return false;
    }

    private static (int index, int total)? Counter(Window window)
    {
        foreach (var t in Walk(window, 0).Where(x => x.ControlType == ControlType.Text))
        {
            var m = MatchesRegex.Match(t.Name ?? "");
            if (m.Success)
            {
                return (int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value));
            }
        }

        return null;
    }

    private static string Breadcrumb(Window window)
    {
        var crumb = window.FindFirstDescendant(cf => cf.ByAutomationId("BreadCrumb"));
        return crumb == null ? "" : string.Join("", Walk(crumb, 0).Where(t => t.ControlType == ControlType.Text).Select(t => t.Name ?? ""));
    }

    // raw-view walk (template parts included), as used for the capture
    private static IEnumerable<AutomationElement> Kids(AutomationElement e)
    {
        var walker = e.Automation.TreeWalkerFactory.GetRawViewWalker();
        for (var c = walker.GetFirstChild(e); c != null; c = walker.GetNextSibling(c))
        {
            yield return c;
        }
    }

    private static IEnumerable<AutomationElement> Walk(AutomationElement e, int depth)
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
}
