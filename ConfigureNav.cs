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

    private static readonly Regex FetchingRegex = new(@"Fetching parameters\.\.\.\s*(\d+)\s*/\s*(\d+)", RegexOptions.Compiled);

    // After opening the Configure tab / Fetch the GUI loads the parameters in the background ("Fetching parameters... 5/688") and values are
    // blank. Waits until that text stays away for 1 s; a counter stuck for 8 s while the search box is enabled counts as finished too.
    public static bool WaitFetchDone(Window window, int timeoutMs, Action<string> log)
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
                log("WARN  | the status bar is stuck on 'Fetching parameters... " + progress + "' but the search box is enabled - treating the load as finished");
                return true;
            }

            Thread.Sleep(1000);
        }

        return false;
    }

    // Right after the Configure view opens (or a tree page is selected) the GUI is still building the parameter list: it can show no parameters yet,
    // or show them without being scrollable yet. A locate that starts then (e.g. a run that starts from row 54, so the first parameter is far down the
    // page) would conclude "not scrollable, not on the first page" and fail. Waits until the list shows parameters and, for a list that has to scroll
    // (more rows than fit on the page), until its scroll pattern is active. Returns the list, or null after timeoutMs.
    public static AutomationElement? WaitListReady(Window w, int timeoutMs, Action<string> log, string what)
    {
        var sw = Stopwatch.StartNew();
        bool announced = false;
        while (sw.ElapsedMilliseconds < timeoutMs)
        {
            var list = w.FindFirstDescendant(cf => cf.ByAutomationId("ParametersList"));
            if (list != null)
            {
                int ids = 0;
                try
                {
                    ids = list.FindAllDescendants().Count(e => (e.AutomationId ?? "").StartsWith("Param", StringComparison.Ordinal));
                }
                catch
                {
                    // list being rebuilt
                }

                var scroll = list.Patterns.Scroll.PatternOrDefault;
                bool scrollable = scroll != null && scroll.VerticallyScrollable.ValueOrDefault;
                // a page with only a few rows may legitimately not scroll: after 3 s of showing rows without scrolling accept it
                if (ids > 0 && (scrollable || sw.ElapsedMilliseconds >= 3000))
                {
                    if (announced)
                    {
                        log("INFO  | the parameter list is ready (" + what + ") after " + (sw.ElapsedMilliseconds / 1000.0).ToString("0.0") + " s");
                    }

                    return list;
                }
            }

            if (!announced && sw.ElapsedMilliseconds > 2500)
            {
                announced = true;
                log("INFO  | waiting for the GUI to finish building the parameter list (" + what + ")");
            }

            Thread.Sleep(250);
        }

        return null;
    }
}
