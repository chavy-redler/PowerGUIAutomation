namespace PowerGUIAutomation;

// The "before you start" frame shown when a test is chosen in the main menu: what must be true BEFORE the test (calibrated unit, GUI in front,
// up-to-date DB, closed workbook, ...) and what the test does to the unit. Built as text lines so the same frame goes to the console (the two header lines cyan, the rest normal gray)
// and, plain, to the log file.
public static class TestInstructions
{
    private const int KeyWidth = 22;
    private const int ValueWidth = 82;

    private sealed record Item(string Key, string Value);

    public static List<(string Text, ConsoleColor Color)> Build(string choice, UnitProfile? unit, string sheetNote)
    {
        string unitName = unit?.Name ?? "(none selected - press U in the menu first)";
        var items = new List<Item>
        {
            new("Unit under test", unitName + ": must be connected and CALIBRATED before the test."),
            new("Power Rider Studio", "Open, with the unit connected. Keep its window in FRONT and do not touch the mouse or the keyboard during the run: the test types into the GUI."),
        };

        string title;
        switch (choice)
        {
            case "1":
                title = "Operate - Second Bar (unit info bar) comparison";
                items.Add(new("Master workbook", "Close 'Power GUI AU-Tests Master.xlsx' in Excel. The ids come from its 'Operate' sheet."));
                items.Add(new("What it does", "Read-only: compares the Second Bar of the Operate tab with the Excel."));
                break;
            case "2":
                title = "Configure - Parameters comparison";
                AddConfigureCommon(items, sheetNote);
                AddResetNote(items);
                items.Add(new("What it does", "Read-only: compares every parameter shown by the GUI (page scan of System and each channel) with its Excel default."));
                break;
            case "3":
                title = "Configure - Change and restore every parameter";
                AddConfigureCommon(items, sheetNote);
                AddResetNote(items);
                items.Add(new("What it does", "WRITES to the unit and saves to flash: every parameter is changed, saved, read back and restored to its Excel default. Never changed: baud rates, addresses, Logic, Group Control."));
                items.Add(new("Stopping", "Ctrl+C stops at once. The parameter being tested may stay on a changed value: it is listed in the failures file - check it on the unit."));
                break;
            case "4":
                title = "Configure - Input validation";
                AddConfigureCommon(items, sheetNote);
                items.Add(new("GUI version", "Developed on Power Rider Studio 2.1.2 or newer; on an older GUI the results are not validated."));
                items.Add(new("What it does", "Types invalid values (letters, empty, Min-1, Max+1, Hebrew, too-long text) into the text fields of System and channel 1 and checks that the GUI refuses them, and that valid ones (Min, Max, text) are accepted. NEVER presses Save; the original value is put back after every input."));
                items.Add(new("Stopping", "Ctrl+C stops at once. A value typed at that moment (Enter applies it, nothing is saved) may still be on the unit: check the parameter listed in the failures file."));
                break;
            default:
                title = "Debug tests";
                AddConfigureCommon(items, sheetNote);
                items.Add(new("What it does", "Read-only: 1 = test 2 only on chosen channels / parameters;  2 = AutomationId scan of all parameters (report to a new workbook in the Excel folder)."));
                break;
        }

        return Frame("Before test " + choice + " - " + title, items);
    }

    // Same row as the unit itself, as a second line of its text (the unit and its state belong together).
    private static void AddResetNote(List<Item> items) =>
        items[0] = items[0] with { Value = items[0].Value + "\nReset the unit to DEFAULT before the run." };

    private static void AddConfigureCommon(List<Item> items, string sheetNote)
    {
        items.Add(new("Studio DB", "The GUI must run with the UP-TO-DATE DB: Min / Max / Default are read from the parameter tooltips. Limits still only in the prerelease DB show only in debug mode - check one tooltip first."));
        items.Add(new("Master workbook", "Close 'Power GUI AU-Tests Master.xlsx' in Excel. " + sheetNote));
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
