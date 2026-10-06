using System.Text.RegularExpressions;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;

namespace PowerGUIAutomation;

// Safety check shared by the Configure tests: the Excel data belongs to ONE unit (part number); running it against another
// unit gives wrong "failures" (parameters that do not exist there are blank / rejected) and, in a write test, wrong writes.
public static class UnitCheck
{
    // Returns null when the connected unit is the selected one, otherwise the reason.
    public static string? Verify(Window window, UnitProfile unit, Action<string> log)
    {
        AppConnection.GoToTab(window, "OperateTab");
        var card = window.FindFirstDescendant(cf => cf.ByAutomationId("OperateUnitCard"));
        if (card == null)
        {
            return "no unit card in the Operate tab - no unit seems to be connected";
        }

        var texts = Walk(card).Where(t => t.ControlType == ControlType.Text).Select(t => (t.Name ?? "").Trim()).Where(t => t.Length > 0).ToList();
        string partText = texts.FirstOrDefault(t => Regex.IsMatch(t, @"RD\s*\d+", RegexOptions.IgnoreCase)) ?? "";
        int? shownPart = UnitProfile.ParsePartNumber(partText);
        int channels = window.FindAllDescendants(cf => cf.ByAutomationId("ChannelName")).Length;
        log("INFO  | connected unit: '" + string.Join(" | ", texts.Take(2)) + "' (part number " + (shownPart?.ToString() ?? "?") + ", " + channels + " channels shown) | selected unit: " + unit.Name);
        if (shownPart == null)
        {
            return "the part number of the connected unit could not be read from the Operate tab (texts: " + string.Join(" | ", texts.Take(3)) + ")";
        }

        if (shownPart != unit.PartNumberValue)
        {
            return "the connected unit has part number RD" + shownPart + " but the selected unit is " + unit.PartNumber
                   + " - the Excel parameters belong to " + unit.PartNumber + " and do not fit this unit";
        }

        return channels != 0 && channels != unit.Channels
            ? "the GUI shows " + channels + " channels but " + unit.PartNumber + " should have " + unit.Channels
            : null;
    }

    private static IEnumerable<AutomationElement> Walk(AutomationElement e, int depth = 0)
    {
        yield return e;
        if (depth > 8)
        {
            yield break;
        }

        var walker = e.Automation.TreeWalkerFactory.GetRawViewWalker();
        for (var c = walker.GetFirstChild(e); c != null; c = walker.GetNextSibling(c))
        {
            foreach (var x in Walk(c, depth + 1))
            {
                yield return x;
            }
        }
    }
}
