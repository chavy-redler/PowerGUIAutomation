using System.Globalization;
using System.Text.RegularExpressions;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;

namespace PowerGUIAutomation;

// Which Studio DB the tests run against. The Min / Max / Default the tests check come from the parameter tooltips of the GUI, i.e. from the DB the GUI loaded - an old
// DB means the unit is checked against old limits (e.g. Temp High Max 130 in the old DB, 120 from 1.0.6.10). Two facts are read and shown at the start of tests 2 and 3
// and in the "Run" tab of the failures file:
//   - what the GUI says: the "DB x.y.z" text next to the GUI version in the title bar;
//   - the newest DB this PC has downloaded: Database_<version>.json in %LocalAppData%\Redler Technologies Ltd\PowerRiderStudio.
// When they differ the GUI may be running with an older DB than the one that is available (stale cache?) - a WARN says so.
public static class StudioDbInfo
{
    private static readonly Regex DbText = new(@"^DB\s+(\S+)$", RegexOptions.Compiled);

    // The text for the log and the Run tab; also logs it (a WARN when the GUI does not show the newest downloaded DB).
    public static string Report(Window window, Action<string> log)
    {
        string shown = ReadShown(window);
        var (newest, downloaded) = NewestDownloaded();
        string text = "GUI shows DB " + (shown.Length > 0 ? shown : "(not found in the title bar)")
                      + " | newest DB downloaded on this PC: " + (newest.Length > 0 ? newest + (downloaded.HasValue ? " (downloaded " + downloaded.Value.ToString("yyyy-MM-dd HH:mm") + ")" : "") : "(none found)");
        log("INFO  | Studio DB | " + text);
        if (shown.Length > 0 && newest.Length > 0 && !Same(shown, newest))
        {
            log("WARN  | Studio DB | the GUI shows DB " + shown + " but the newest DB on this PC is " + newest + " - the limits (Min / Max / Default) are read from the GUI, so the GUI may be running with an older DB than the one that is available; check which DB the GUI uses before trusting the limit failures");
        }

        return text;
    }

    private static string ReadShown(Window window)
    {
        try
        {
            foreach (var e in window.FindAllDescendants(cf => cf.ByControlType(ControlType.Text)))
            {
                var m = DbText.Match((e.Name ?? "").Trim());
                if (m.Success)
                {
                    return m.Groups[1].Value;
                }
            }
        }
        catch
        {
            // best effort - the tests do not depend on it
        }

        return "";
    }

    private static (string Version, DateTime? Downloaded) NewestDownloaded()
    {
        try
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Redler Technologies Ltd", "PowerRiderStudio");
            Version? best = null;
            string bestFile = "";
            foreach (var file in Directory.GetFiles(dir, "Database_*.json"))
            {
                var m = Regex.Match(Path.GetFileNameWithoutExtension(file), @"^Database_(\d+(?:_\d+){1,3})$");
                if (m.Success && Version.TryParse(m.Groups[1].Value.Replace('_', '.'), out var v) && (best == null || v > best))
                {
                    best = v;
                    bestFile = file;
                }
            }

            return best == null ? ("", null) : (best.ToString(), File.GetLastWriteTime(bestFile));
        }
        catch
        {
            return ("", null);
        }
    }

    // "1.0.5" vs "1.0.5.0" are the same version
    private static bool Same(string a, string b) => Version.TryParse(a, out var x) && Version.TryParse(b, out var y) ? Norm(x) == Norm(y) : string.Equals(a, b, StringComparison.OrdinalIgnoreCase);

    private static Version Norm(Version v) => new(v.Major, Math.Max(v.Minor, 0), Math.Max(v.Build, 0), Math.Max(v.Revision, 0));
}
