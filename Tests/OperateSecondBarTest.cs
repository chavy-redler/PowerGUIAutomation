using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using ClosedXML.Excel;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Capturing;
using FlaUI.Core.Definitions;
using FlaUI.UIA3;

namespace PowerGUIAutomation.Tests;

// Runs the "Operate" tab as an ordered sequence of steps - exactly like a
// manual tester's script. Every step is either:
//   - Click: press a button (found by AutomationId parsed out of the "Access
//     Method" column, or by two fixed text labels when no AutomationId
//     exists), or
//   - Verify: read a value and compare it, either once against OperateUnitCard
//     (Scope=Unit) or once per connected channel (Scope=PerChannel), optionally
//     also checking the ChannelToggle icon's actual on-screen color.
// Nothing about a specific scenario is hardcoded here - adding/changing a
// step means editing the Excel, not this file.
public class OperateSecondBarTest
{
    private readonly string _excelPath;
    private readonly Action<string> _log;
    private static readonly Regex AutomationIdRegex = new(@"AutomationId='([^']+)'", RegexOptions.IgnoreCase);

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

        List<Step> steps;
        try
        {
            steps = ReadSteps();
        }
        catch (Exception ex)
        {
            _log("FAIL | Could not read test data from Excel: " + ex.Message);
            return false;
        }

        foreach (var step in steps)
        {
            bool ok = step.Action == "Click"
                ? RunClickStep(unitCard, step)
                : RunVerifyStep(unitCard, step);

            if (!ok)
            {
                _log("STOPPED | halted after first mismatch - remaining steps were not checked");
                return false;
            }
        }

        _log("DONE | all steps passed");
        return true;
    }

    private bool RunClickStep(AutomationElement unitCard, Step step)
    {
        AutomationElement? button = null;

        if (step.ButtonAutomationId != null)
        {
            button = unitCard.FindFirstDescendant(cf => cf.ByAutomationId(step.ButtonAutomationId));
        }
        else if (step.ButtonLabel1 != null && step.ButtonLabel2 != null)
        {
            button = FindActionButtonByLabels(unitCard, step.ButtonLabel1, step.ButtonLabel2);
        }

        if (button == null)
        {
            _log("FAIL  | " + step.Entity + " | button not found (AutomationId=" + step.ButtonAutomationId + ", labels=" + step.ButtonLabel1 + "/" + step.ButtonLabel2 + ")");
            return false;
        }

        button.AsButton().Invoke();
        Thread.Sleep(1500); // let every affected element finish updating
        _log("CLICK | " + step.Entity);
        return true;
    }

    private bool RunVerifyStep(AutomationElement unitCard, Step step)
    {
        if (step.Scope == "PerChannel")
        {
            var channelNames = unitCard.FindAllDescendants(cf => cf.ByAutomationId("ChannelName"));
            if (channelNames.Length == 0)
            {
                _log("FAIL  | " + step.Entity + " | no channels found");
                return false;
            }

            for (int i = 0; i < channelNames.Length; i++)
            {
                var channelScope = channelNames[i].Parent;
                if (!VerifyOnce(channelScope, step, "Channel " + (i + 1)))
                {
                    return false;
                }
            }

            return true;
        }

        return VerifyOnce(unitCard, step, step.Entity);
    }

    private bool VerifyOnce(AutomationElement scope, Step step, string label)
    {
        if (step.MatchType.Equals("TD", StringComparison.OrdinalIgnoreCase))
        {
            _log("SKIP | " + label + " | " + step.Entity + " marked TD in Excel - not yet checked");
            return true;
        }

        var samples = new List<string>();
        for (int i = 0; i < step.SampleCount; i++)
        {
            if (i > 0)
            {
                Thread.Sleep((int)(step.SampleIntervalSeconds * 1000));
            }

            var texts = scope.FindAllDescendants(cf => cf.ByControlType(ControlType.Text));
            samples.Add(step.Index < texts.Length ? (texts[step.Index].Name ?? "") : "<INDEX OUT OF RANGE>");
        }

        bool valuePassed = step.MatchType.Equals("Range", StringComparison.OrdinalIgnoreCase)
            ? samples.All(s => TryExtractNumber(s, out double v) && v >= step.Min && v <= step.Max)
            : samples.All(s => s.Trim() == step.ExpectedValue.Trim());

        bool colorPassed = true;
        string colorDescription = "";
        if (!string.IsNullOrEmpty(step.ExpectedColor))
        {
            var toggle = scope.FindFirstDescendant(cf => cf.ByAutomationId("ChannelToggle"));
            if (toggle == null)
            {
                colorPassed = false;
                colorDescription = " color=<TOGGLE NOT FOUND>";
            }
            else
            {
                using var capture = FlaUI.Core.Capturing.Capture.Element(toggle);
                var bmp = capture.Bitmap;
                var pixel = bmp.GetPixel(bmp.Width / 2, bmp.Height / 2);
                bool isGreen = pixel.G > pixel.R && pixel.G >= pixel.B;
                bool expectGreen = step.ExpectedColor.Equals("Green", StringComparison.OrdinalIgnoreCase);
                colorPassed = isGreen == expectGreen;
                colorDescription = $" color=R={pixel.R} G={pixel.G} B={pixel.B} (green={isGreen}, expected={step.ExpectedColor})";
            }
        }

        bool passed = valuePassed && colorPassed;
        string expectedDescription = step.MatchType.Equals("Range", StringComparison.OrdinalIgnoreCase)
            ? $"[{step.Min}-{step.Max}] x{step.SampleCount} samples"
            : step.ExpectedValue;

        _log((passed ? "PASS  | " : "FAIL  | ") + label + " | " + step.Entity + " | Expected=" + expectedDescription + " | Actual=[" + string.Join(", ", samples) + "]" + colorDescription);

        return passed;
    }

    private static AutomationElement? FindActionButtonByLabels(AutomationElement scope, string label1, string label2)
    {
        var buttons = scope.FindAllDescendants(cf => cf.ByControlType(ControlType.Button));
        foreach (var button in buttons)
        {
            var texts = button.FindAllDescendants(cf => cf.ByControlType(ControlType.Text));
            if (texts.Length >= 2 && texts[0].Name?.Trim() == label1 && texts[1].Name?.Trim() == label2)
            {
                return button;
            }
        }
        return null;
    }

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

    private class Step
    {
        public string Entity = "";
        public string Action = "Verify"; // "Verify" or "Click"
        public string Scope = "Unit";    // "Unit" or "PerChannel"
        public int Index;
        public string ExpectedValue = "";
        public string MatchType = "Exact";
        public double Min, Max;
        public int SampleCount = 1;
        public double SampleIntervalSeconds = 0;
        public string? ExpectedColor;
        public string? ButtonAutomationId;
        public string? ButtonLabel1;
        public string? ButtonLabel2;
    }

    private List<Step> ReadSteps()
    {
        using var workbook = new XLWorkbook(_excelPath);
        var ws = workbook.Worksheet("Operate");
        var steps = new List<Step>();

        foreach (var row in ws.RowsUsed().Skip(1)) // skip header
        {
            string action = row.Cell(14).GetString().Trim();
            string indexCell = row.Cell(7).GetString().Trim();

            bool isClick = action.Equals("Click", StringComparison.OrdinalIgnoreCase);
            if (!isClick && indexCell.Length == 0)
            {
                continue; // documentation-only row, no test data
            }

            var step = new Step
            {
                Entity = row.Cell(3).GetString().Trim(),
                Action = isClick ? "Click" : "Verify",
                Scope = row.Cell(15).GetString().Trim() is { Length: > 0 } s ? s : "Unit",
            };

            if (isClick)
            {
                string accessMethod = row.Cell(6).GetString();
                var match = AutomationIdRegex.Match(accessMethod);
                step.ButtonAutomationId = match.Success ? match.Groups[1].Value : null;

                string labels = row.Cell(17).GetString().Trim();
                if (labels.Length > 0)
                {
                    var parts = labels.Split('|');
                    step.ButtonLabel1 = parts.ElementAtOrDefault(0);
                    step.ButtonLabel2 = parts.ElementAtOrDefault(1);
                }
            }
            else
            {
                step.Index = int.Parse(indexCell);
                step.ExpectedValue = row.Cell(8).GetString().Trim();
                step.MatchType = row.Cell(9).GetString().Trim() is { Length: > 0 } mt ? mt : "Exact";
                step.Min = ParseDoubleOrZero(row.Cell(10).GetString());
                step.Max = ParseDoubleOrZero(row.Cell(11).GetString());
                step.SampleCount = ParseIntOrDefault(row.Cell(12).GetString(), 1);
                step.SampleIntervalSeconds = ParseDoubleOrZero(row.Cell(13).GetString());
                string color = row.Cell(16).GetString().Trim();
                step.ExpectedColor = color.Length > 0 ? color : null;
            }

            steps.Add(step);
        }

        return steps;
    }

    private static double ParseDoubleOrZero(string s) =>
        double.TryParse(s.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double v) ? v : 0;

    private static int ParseIntOrDefault(string s, int fallback) =>
        int.TryParse(s.Trim(), out int v) ? v : fallback;
}
