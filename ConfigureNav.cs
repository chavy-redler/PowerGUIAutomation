using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Threading;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;

namespace PowerGUIAutomation;

// Navigation of the Configure tab's parameter tree, shared by the tests that locate parameters by page (not by search):
// "System Parameters", "Channel N Parameters" (under "Channel Parameters") and "Group Channel Control" (the page that holds every
// "Channel N Group Control" parameter). Selecting a node shows that scope's whole parameter list, a virtualized list that is scrolled.
public static class ConfigureNav
{
    public const int GroupScope = -1;

    private static readonly Regex GroupRowRegex = new(@"^Channel \d+ Group Control$", RegexOptions.Compiled);

    // Scope (tree page) of an Excel row: 0 = System, 1..N = channel, GroupScope = "Group Channel Control".
    public static int ScopeOf(int channel, string parameterName) => channel > 0 && GroupRowRegex.IsMatch(parameterName) ? GroupScope : channel;

    private static AutomationElement[] TreeItems(Window w) => w.FindAllDescendants(cf => cf.ByControlType(ControlType.TreeItem));

    public static string ReadBreadcrumb(Window w)
    {
        var crumb = w.FindFirstDescendant(cf => cf.ByAutomationId("BreadCrumb"));
        if (crumb == null)
        {
            return "";
        }

        var walker = crumb.Automation.TreeWalkerFactory.GetRawViewWalker();
        return string.Join("", Texts(crumb, walker, 0));
    }

    private static IEnumerable<string> Texts(AutomationElement root, FlaUI.Core.ITreeWalker walker, int depth)
    {
        for (var c = walker.GetFirstChild(root); c != null; c = walker.GetNextSibling(c))
        {
            if (c.ControlType == ControlType.Text)
            {
                yield return (c.Name ?? "").Trim();
            }

            if (depth < 6)
            {
                foreach (var t in Texts(c, walker, depth + 1))
                {
                    yield return t;
                }
            }
        }
    }

    // Selects the scope's tree node and waits until the GUI's breadcrumb says so. Returns the breadcrumb, or null when it did not work.
    public static string? SelectScope(Window w, int scope)
    {
        string nodeName = scope == GroupScope ? "Group Channel Control" : scope > 0 ? "Channel " + scope + " Parameters" : "System Parameters";
        var node = TreeItems(w).FirstOrDefault(x => x.Name == nodeName);
        if (node == null && scope > 0)
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
            bool ok = scope == GroupScope
                ? crumb.Contains("Group Channel Control", StringComparison.Ordinal)
                : scope > 0
                    ? crumb.Contains("Channel " + scope + " Parameters", StringComparison.Ordinal)
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
}
