using System.Text.RegularExpressions;
using System.Threading;
using ClosedXML.Excel;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;

namespace PowerGUIAutomation.Tests;

// Ported from the standalone PowerRiderParamCheck project: reads every
// parameter row from the "System Parameters" / "Channel Parameters" sheets,
// uses the Configure tab's own search box to bring each one into view (sidesteps
// list virtualization entirely - no scrolling/keyboard navigation needed),
// and compares Name + Value. Stops on the first mismatch, same as every other
// test in this project.
public class ConfigureParametersTest
{
    private const int MaxFindRetries = 12;
    private const int RetryDelayMs = 400;

    private readonly string _excelPath;
    private readonly Action<string> _log;
    private static readonly Regex AutomationIdRegex = new(@"automationid='([^']+)'", RegexOptions.IgnoreCase);
    private static readonly Regex BracketedRegex = new(@"^\[\s*([^,\]]+?)\s*[,\]]");

    public ConfigureParametersTest(string excelPath, Action<string> log)
    {
        _excelPath = excelPath;
        _log = log;
    }

    public bool Run()
    {
        using var automation = new UIA3Automation();
        var mainWindow = AppConnection.Attach(automation);
        AppConnection.GoToTab(mainWindow, "ConfigureTab");

        var searchBoxElement = mainWindow.FindFirstDescendant(cf => cf.ByAutomationId("searchTextBox"));
        if (searchBoxElement == null)
        {
            _log("FAIL | could not find searchTextBox - is the Configure tab open?");
            return false;
        }

        var searchBox = searchBoxElement.AsTextBox();

        List<(string excelRow, string parameterPath, string expectedName, string valuePath, string expectedValue)> dataRows;
        try
        {
            dataRows = ReadTestData();
        }
        catch (Exception ex)
        {
            _log("FAIL | Could not read test data from Excel: " + ex.Message);
            return false;
        }

        int pass = 0, fail = 0, skip = 0;

        foreach (var (excelRow, parameterPath, expectedName, valuePath, expectedValue) in dataRows)
        {
            if (string.IsNullOrWhiteSpace(parameterPath))
            {
                _log("SKIP  | Row " + excelRow + " | " + expectedName + " | no RxPath defined");
                skip++;
                continue;
            }

            string? nameAutomationId = ExtractAutomationId(parameterPath);
            if (nameAutomationId == null)
            {
                _log("FAIL  | Row " + excelRow + " | could not extract automationid from ParameterPath: " + parameterPath);
                _log("STOPPED | halted after first mismatch - remaining parameters were not checked");
                return false;
            }

            _log("START | Row " + excelRow + " | Expected='" + expectedName + "'");

            // Bring the row into view via the app's own search box - paste
            // (Ctrl+V) the whole string in one shot rather than simulating
            // keystrokes, so it isn't corrupted by whatever OS keyboard layout
            // happens to be active.
            searchBox.Focus();
            searchBox.Text = "";
            TextCopy.ClipboardService.SetText(expectedName);
            Keyboard.TypeSimultaneously(VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_V);

            var nameElement = FindWithRetry(mainWindow, nameAutomationId);
            if (nameElement == null)
            {
                _log("FAIL  | Row " + excelRow + " | Expected='" + expectedName + "' | name element not found after " + MaxFindRetries + " attempts (automationid=" + nameAutomationId + ")");
                _log("STOPPED | halted after first mismatch - remaining parameters were not checked");
                return false;
            }

            string actualName = (nameElement.Name ?? "").Trim();
            bool namePassed = string.Equals(actualName, expectedName, StringComparison.Ordinal);
            _log((namePassed ? "PASS  | " : "FAIL  | ") + "Row " + excelRow + " | Expected='" + expectedName + "' | Actual='" + actualName + "'");
            if (namePassed)
            {
                pass++;
            }
            else
            {
                fail++;
                _log("STOPPED | halted after first mismatch - remaining parameters were not checked");
                return false;
            }

            string? valueAutomationId = ExtractAutomationId(valuePath);
            if (valueAutomationId == null)
            {
                _log("SKIP  | Row " + excelRow + " | no Value automationid defined");
                skip++;
                continue;
            }

            string actualValue = "<NOT FOUND>";
            var valueElement = FindWithRetry(mainWindow, valueAutomationId);
            if (valueElement != null)
            {
                // Finding the element doesn't mean its ValuePattern has caught up
                // yet - right after Focus() it can still read back empty for a
                // beat (confirmed: Serial Baud Rate, now Communication Timeout).
                // Retry the read itself, not just the find, until we get
                // something non-empty or run out of attempts.
                for (int attempt = 1; attempt <= MaxFindRetries; attempt++)
                {
                    valueElement.Focus();
                    Thread.Sleep(300);

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

                    if (candidate.Length > 0)
                    {
                        actualValue = candidate;
                        break;
                    }

                    if (attempt == 1)
                    {
                        actualValue = ""; // remember we did get a (empty) read, in case every attempt stays empty
                    }

                    Thread.Sleep(RetryDelayMs);
                }
            }

            // Some ComboBox-backed parameters expose their raw bound object
            // instead of the displayed text - "[250, 3]" or "[Disable, 0]"
            // rather than plain "250"/"Disable". Take the first element inside
            // the brackets (numeric or text) for comparison.
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

            _log((valuePassed ? "PASS  | " : "FAIL  | ") + "Row " + excelRow + " | Value Expected='" + expectedValue + "' | Actual='" + actualValue + "'");
            if (valuePassed)
            {
                pass++;
            }
            else
            {
                fail++;
                _log("STOPPED | halted after first mismatch - remaining parameters were not checked");
                return false;
            }
        }

        searchBox.Focus();
        searchBox.Text = "";
        _log("DONE | pass=" + pass + " fail=" + fail + " skip=" + skip);

        return fail == 0;
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

    private static string? ExtractAutomationId(string? rxPath)
    {
        if (string.IsNullOrWhiteSpace(rxPath)) return null;
        var matches = AutomationIdRegex.Matches(rxPath);
        return matches.Count > 0 ? matches[^1].Groups[1].Value : null;
    }

    private List<(string, string, string, string, string)> ReadTestData()
    {
        using var workbook = new XLWorkbook(_excelPath);
        var result = new List<(string, string, string, string, string)>();

        foreach (var sheetName in new[] { "System Parameters", "Channel Parameters" })
        {
            if (!workbook.Worksheets.Contains(sheetName))
            {
                continue;
            }

            var ws = workbook.Worksheet(sheetName);
            foreach (var row in ws.RowsUsed().Skip(1)) // skip header
            {
                result.Add((
                    row.Cell(1).GetString().Trim(),
                    row.Cell(2).GetString().Trim(),
                    row.Cell(3).GetString().Trim(),
                    row.Cell(4).GetString().Trim(),
                    row.Cell(5).GetString().Trim()
                ));
            }
        }

        return result;
    }
}
