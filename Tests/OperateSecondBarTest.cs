using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using ClosedXML.Excel;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;

namespace PowerGUIAutomation.Tests;

// Compares the Operate screen's top unit bar (product description, unit
// address, refresh rate, power, current, unit BIT) against the "OperateSecondBar"
// sheet in the master Excel. Every field is found by fixed position under the
// OperateUnitCard anchor (validated live via FlaUI - see the sheet's own
// "Access Method" notes on the "Operate" tab for how each index was confirmed).
public class OperateSecondBarTest
{
    private readonly string _excelPath;
    private readonly Action<string> _log;

    public OperateSecondBarTest(string excelPath, Action<string> log)
    {
        _excelPath = excelPath;
        _log = log;
    }

    public bool Run()
    {
        using var automation = new UIA3Automation();
        var mainWindow = AppConnection.Attach(automation);
        AppConnection.GoToTab(mainWindow, "OperateTab");

        var unitCard = mainWindow.FindFirstDescendant(cf => cf.ByAutomationId("OperateUnitCard"));
        if (unitCard == null)
        {
            _log("FAIL | OperateUnitCard not found - is a unit connected?");
            return false;
        }

        List<TestRow> rows;
        try
        {
            rows = ReadTestData();
        }
        catch (Exception ex)
        {
            _log("FAIL | Could not read test data from Excel: " + ex.Message);
            return false;
        }

        bool allPassed = true;

        foreach (var row in rows)
        {
            // "TD" = To Do - documented in the Excel but not actively checked yet
            // (e.g. Refresh Rate's real fluctuation range isn't nailed down). Skips
            // without touching pass/fail or the stop-on-first-mismatch behavior.
            if (row.MatchType.Equals("TD", StringComparison.OrdinalIgnoreCase))
            {
                _log($"SKIP | {row.Entity} | marked TD in Excel - not yet checked");
                continue;
            }

            var samples = new List<string>();

            for (int i = 0; i < row.SampleCount; i++)
            {
                if (i > 0)
                {
                    Thread.Sleep((int)(row.SampleIntervalSeconds * 1000));
                }

                // Re-find the Text descendants fresh on every sample - the live
                // values update in place, so we need a live read each time, not
                // a cached reference from before the wait.
                var texts = unitCard.FindAllDescendants(cf => cf.ByControlType(ControlType.Text));
                string actual = row.Index < texts.Length ? (texts[row.Index].Name ?? "") : "<INDEX OUT OF RANGE>";
                samples.Add(actual);
            }

            bool passed = row.MatchType.Equals("Range", StringComparison.OrdinalIgnoreCase)
                ? samples.All(s => TryExtractNumber(s, out double v) && v >= row.Min && v <= row.Max)
                : samples.All(s => s.Trim() == row.ExpectedValue.Trim());

            string expectedDescription = row.MatchType.Equals("Range", StringComparison.OrdinalIgnoreCase)
                ? $"[{row.Min}-{row.Max}] x{row.SampleCount} samples"
                : row.ExpectedValue;

            _log($"{(passed ? "PASS" : "FAIL")} | {row.Entity} | Expected={expectedDescription} | Actual=[{string.Join(", ", samples)}]");

            if (!passed)
            {
                allPassed = false;
                _log("STOPPED | halted after first mismatch - remaining fields were not checked");
                break;
            }
        }

        return allPassed;
    }

    // A value like "191 ms" or "0.00 W" - pull out the leading number so a
    // Range check can compare it numerically regardless of the unit suffix.
    private static bool TryExtractNumber(string text, out double value)
    {
        var match = Regex.Match(text, @"-?\d+(\.\d+)?");
        if (match.Success)
        {
            return double.TryParse(match.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        value = 0;
        return false;
    }

    // Reads straight from the "Operate" tab (not a separate test-data sheet) -
    // that tab documents every element on screen, most of which have no test
    // data at all. Columns A-F are the Screen/Sub Menu/Entity/Value/Impact/
    // Access-Method documentation; G-M are the test-execution columns we
    // added (Test Index, Expected Value, Match Type, Min, Max, Sample Count,
    // Sample Interval). A row only becomes a test row once it has something
    // in the "Test Index" column (G) - everything else on the tab is skipped.
    private List<TestRow> ReadTestData()
    {
        using var workbook = new XLWorkbook(_excelPath);
        var ws = workbook.Worksheet("Operate");
        var rows = new List<TestRow>();

        foreach (var row in ws.RowsUsed().Skip(1)) // skip header
        {
            string indexCell = row.Cell(7).GetString().Trim();
            if (indexCell.Length == 0)
            {
                continue; // no test data on this row - it's documentation only
            }

            rows.Add(new TestRow
            {
                Entity = row.Cell(3).GetString().Trim(),
                Index = int.Parse(indexCell),
                ExpectedValue = row.Cell(8).GetString().Trim(),
                MatchType = row.Cell(9).GetString().Trim(),
                Min = ParseDoubleOrZero(row.Cell(10).GetString()),
                Max = ParseDoubleOrZero(row.Cell(11).GetString()),
                SampleCount = ParseIntOrDefault(row.Cell(12).GetString(), 1),
                SampleIntervalSeconds = ParseDoubleOrZero(row.Cell(13).GetString()),
            });
        }

        return rows;
    }

    private static double ParseDoubleOrZero(string s) =>
        double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double v) ? v : 0;

    private static int ParseIntOrDefault(string s, int fallback) =>
        int.TryParse(s.Trim(), out int v) ? v : fallback;
}
