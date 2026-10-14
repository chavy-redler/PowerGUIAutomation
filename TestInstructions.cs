namespace PowerGUIAutomation;

// Everything the user is told about a test BEFORE it runs, kept in one place so that a rule applies to every test: the "TEST RUNNER" banner title and
// the "Before test N" frame. The rows every test shows (unit connected + calibrated + reset to default, the GUI in front, the closed workbook, what the test
// does, how to stop it) are built by ONE method from the catalog below. To add a test: add one TestInfo to Catalog (and the menu line in Program.cs).
//
// The frame is built as text lines: the console prints them (the two header lines cyan, the rest normal gray) and the log file gets them plain.
public static class TestInstructions
{
    private const int KeyWidth = 22;
    private const int ValueWidth = 82;

    private sealed record Item(string Key, string Value);

    // Configure = the test reads limits from the tooltips and needs the up-to-date DB. Graceful = the test can stop cleanly between parameters (see
    // StopControl); StopRisk = what may be left on the unit when it is stopped at once.
    private sealed record TestInfo(string Choice, string Title, string Banner, bool Configure, bool Graceful, string WhatItDoes, string StopRisk = "",
                                   string Workbook = "", (string Key, string Value)[]? Extra = null, string Menu = "");

    private static readonly TestInfo[] Catalog =
    {
        new("1", "Configure - Parameters comparison", "CONFIGURE - PARAMETERS COMPARISON", true, false,
            "Read-only: compares every parameter shown by the GUI (page scan of System and each channel) with its Excel default.",
            Menu: "Configure - Parameters comparison"),
        new("2", "Configure - Change and restore every parameter", "CONFIGURE - CHANGE AND RESTORE EVERY PARAMETER", true, true,
            "WRITES to the unit and saves to flash: every parameter is changed, saved, read back, put back to the value it had (its default after the reset) and checked against the Excel default. Never changed: baud rates, addresses, Logic, Group Control.",
            StopRisk: "the parameter being tested may stay on a changed value, it is listed in the failures file - check it on the unit.",
            Extra: new[] { ("Loop / overnight", "At the start you can run it in a LOOP (N rounds, or round after round until you stop it, with an optional time limit): every round is announced in the log and every failure says its round. For an overnight run: disable Windows sleep, the lock screen and the screen saver (the test keeps the PC and the display awake but cannot unlock a locked screen), keep the unit powered, choose 'continue' on failure.") },
            Menu: "Configure - Change and restore every parameter (writes to the unit)"),
        new("3", "Configure - Input validation", "CONFIGURE - INPUT VALIDATION (INVALID VALUES MUST BE REJECTED)", true, true,
            "Types invalid values (letters, empty, Min-1, Max+1, Hebrew, too-long text) into the text fields of every System and channel parameter (all rows, or the rows you choose), and checks that the GUI refuses them, and that valid ones (Min, Max, text) are accepted. NEVER presses Save; the original value is put back after every input.",
            StopRisk: "a value typed at that moment (Enter applies it, nothing is saved) may still be on the unit - check the parameter listed in the failures file.",
            Extra: new[] { ("GUI version", "Developed on Power Rider Studio 2.1.2 or newer; on an older GUI the results are not validated.") },
            Menu: "Configure - Input validation (types invalid values, never saves to flash)"),
        new("4", "Debug tests", "DEBUG TESTS", true, false,
            "1 = test 1 only on chosen channels / parameters (read-only);  2 = AutomationId scan of all parameters (read-only, report to a new workbook in the Excel folder);  3 = LOOP: test 2's change-and-restore on a group of parameters, round after round until you stop it or one hour has passed (WRITES to the unit; the stop key then stops after the parameter that is being checked).",
            Menu: "Debug tests: Configure comparison on chosen channels / parameters, AutomationId scan, loop of change-and-restore on a group of parameters"),
    };

    // The test that is running (set by the main menu): the failures workbooks name it, because the Failures folder collects the files of every test.
    public static string Current { get; set; } = "";

    // "3 - Configure - Change and restore every parameter (writes to the unit)"; the same text as the menu line.
    public static string Describe(string? choice)
    {
        var info = Catalog.FirstOrDefault(t => t.Choice == choice?.Trim());
        return info == null ? "" : info.Choice + " - " + info.Menu;
    }

    public static IEnumerable<string> MenuLines => Catalog.Select(t => t.Choice + ". " + t.Menu);

    public static bool IsTest(string? choice) => Catalog.Any(t => t.Choice == choice?.Trim());

    // The title of the "TEST RUNNER" banner (null = not a test).
    public static string? BannerTitle(string? choice) => Catalog.FirstOrDefault(t => t.Choice == choice?.Trim())?.Banner;

    // Does the test stop cleanly between parameters when the stop key is pressed (StopControl)?
    public static bool IsGraceful(string? choice) => Catalog.FirstOrDefault(t => t.Choice == choice?.Trim())?.Graceful ?? false;

    public static List<(string Text, ConsoleColor Color)> Build(string choice, UnitProfile? unit)
    {
        var info = Catalog.First(t => t.Choice == choice.Trim());
        string unitName = unit?.Name ?? "(none selected - press U in the menu first)";
        var items = new List<Item>
        {
            // the unit and its state belong together: connected, calibrated and at its defaults (the tests assume every parameter starts at its Excel default)
            new("Unit under test", unitName + ": must be connected and CALIBRATED before the test.\nReset the unit to DEFAULT before the run."),
            new("Power Rider Studio", "Open, with the unit connected. Keep its window in FRONT and do not touch the mouse or the keyboard during the run: the test types into the GUI. The red X / green V next to a value are read from a screenshot, so Windows display scaling must be 100% and the whole window must stay visible on the screen (not minimized, covered or partly off-screen). Any resolution or number of screens is fine."),
        };

        if (info.Configure)
        {
            items.Add(new("Studio DB", "The GUI must run with the UP-TO-DATE DB: Min / Max / Default are read from the parameter tooltips. Limits still only in the prerelease DB show only in debug mode - check one tooltip first."));
        }

        string workbook = info.Workbook.Length > 0 ? info.Workbook : unit == null ? "" : "Its sheet '" + unit.ParameterSheet + "' must exist and be valid.";
        items.Add(new("Master workbook", ("Close 'Power GUI AU-Tests Master.xlsx' in Excel. " + workbook).Trim()));
        foreach (var (key, value) in info.Extra ?? Array.Empty<(string, string)>())
        {
            items.Add(new(key, value));
        }

        items.Add(new("What it does", info.WhatItDoes));
        items.Add(new("Stopping", info.Graceful
            ? StopControl.KeyName + " (works while the GUI is in front) stops the test after the parameter that is being checked, and writes the failures file. Press it again to stop at once: "
              + info.StopRisk + " Ctrl+C works only when this console window is in front."
            : StopControl.KeyName + " (works while the GUI is in front) stops the test at once. Ctrl+C works only when this console window is in front."));

        return Frame("Before test " + info.Choice + " - " + info.Title, items);
    }

    // ----------------------------------------------------------------- drawing

    private static List<(string, ConsoleColor)> Frame(string title, List<Item> items)
    {
        int inner = KeyWidth + ValueWidth + 5;          // between the outer verticals: " key " (+2), the middle vertical, " value " (+2)
        string titleText = " " + title + " ";
        var lines = new List<(string, ConsoleColor)>
        {
            ("┌─" + titleText + new string('─', Math.Max(0, inner - titleText.Length - 1)) + "┐", ConsoleColor.Cyan),
            ("│" + " Requirement".PadRight(KeyWidth + 2) + "│" + " What to check".PadRight(ValueWidth + 2) + "│", ConsoleColor.Cyan),
            ("├" + new string('─', KeyWidth + 2) + "┼" + new string('─', ValueWidth + 2) + "┤", ConsoleColor.Gray),
        };

        for (int i = 0; i < items.Count; i++)
        {
            var item = items[i];
            var wrapped = Wrap(item.Value, ValueWidth);
            for (int l = 0; l < wrapped.Count; l++)
            {
                string key = l == 0 ? item.Key : "";
                lines.Add(("│ " + key.PadRight(KeyWidth) + " │ " + wrapped[l].PadRight(ValueWidth) + " │", ConsoleColor.Gray));
            }

            if (i < items.Count - 1)
            {
                lines.Add(("├" + new string('─', KeyWidth + 2) + "┼" + new string('─', ValueWidth + 2) + "┤", ConsoleColor.Gray));
            }
        }

        lines.Add(("└" + new string('─', KeyWidth + 2) + "┴" + new string('─', ValueWidth + 2) + "┘", ConsoleColor.Gray));
        return lines;
    }

    private static List<string> Wrap(string text, int width)
    {
        var all = new List<string>();
        foreach (var paragraph in text.Split('\n'))
        {
            all.AddRange(WrapParagraph(paragraph, width));
        }

        return all;
    }

    private static List<string> WrapParagraph(string text, int width)
    {
        var result = new List<string>();
        string line = "";
        foreach (var word in text.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            string w = word;
            while (w.Length > width)          // a path longer than the column
            {
                if (line.Length > 0)
                {
                    result.Add(line);
                    line = "";
                }

                result.Add(w.Substring(0, width));
                w = w.Substring(width);
            }

            if (line.Length == 0)
            {
                line = w;
            }
            else if (line.Length + 1 + w.Length <= width)
            {
                line += " " + w;
            }
            else
            {
                result.Add(line);
                line = w;
            }
        }

        if (line.Length > 0 || result.Count == 0)
        {
            result.Add(line);
        }

        return result;
    }
}
