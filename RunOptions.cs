using System.Text.RegularExpressions;

namespace PowerGUIAutomation;

// Which Excel rows a test should run: "54-70, 72, 100-" = rows 54 to 70, row 72, and row 100 to the end of the sheet. Empty = every row.
// Accepted pieces (separated by commas or semicolons):
//   72        a single row
//   54-70     a range of rows
//   100-      from row 100 to the end
//   -60       from the beginning up to row 60
public sealed class RowSelection
{
    private readonly List<(int from, int to)> _ranges = new();

    // null = no selection (every row). Returns false with a message when the text cannot be understood.
    public static bool TryParse(string? text, out RowSelection? selection, out string error)
    {
        selection = null;
        error = "";
        string t = (text ?? "").Trim();
        if (t.Length == 0)
        {
            return true;
        }

        var result = new RowSelection();
        foreach (var raw in t.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
        {
            string piece = Regex.Replace(raw.Trim(), @"^(rows?|r)\s*", "", RegexOptions.IgnoreCase);
            Match m;
            if ((m = Regex.Match(piece, @"^(\d+)$")).Success)
            {
                int n = int.Parse(m.Groups[1].Value);
                result._ranges.Add((n, n));
            }
            else if ((m = Regex.Match(piece, @"^(\d+)\s*-\s*(\d+)$")).Success)
            {
                int a = int.Parse(m.Groups[1].Value), b = int.Parse(m.Groups[2].Value);
                result._ranges.Add((Math.Min(a, b), Math.Max(a, b)));
            }
            else if ((m = Regex.Match(piece, @"^(\d+)\s*-$")).Success)
            {
                result._ranges.Add((int.Parse(m.Groups[1].Value), int.MaxValue));
            }
            else if ((m = Regex.Match(piece, @"^-\s*(\d+)$")).Success)
            {
                result._ranges.Add((0, int.Parse(m.Groups[1].Value)));
            }
            else
            {
                error = "cannot understand '" + raw.Trim() + "' - use row numbers like 54-70, 72, 100-";
                return false;
            }
        }

        selection = result;
        return true;
    }

    // Set when the rows come from the failures file of an earlier run (see Ask): the file, and the parameter name expected at each row.
    public string? SourceFile { get; private set; }
    private Dictionary<int, string> _expectedNames = new();

    public bool Contains(int row) => _ranges.Any(r => row >= r.from && row <= r.to);

    // The selection in words, for the log and the failures file (so that someone who only runs the test understands "40-").
    public string Describe() => string.Join("; ", _ranges.Select(r =>
        r.from == r.to ? "row " + r.from
        : r.to == int.MaxValue ? "rows " + r.from + "-end (from row " + r.from + " to the end of the sheet)"
        : r.from == 0 ? "rows start-" + r.to + " (from the first row up to row " + r.to + ")"
        : "rows " + r.from + "-" + r.to));

    // The question every Configure test asks: which Excel rows to run. Besides "54-70, 72, 100-" the answer can be the PATH of a failures file
    // (drag the file into the console window): then ONLY the rows listed in its "Failed" sheet run. Returns null = every row.
    // 'description' says where the rows came from (for the START line of the log).
    // number > 0: the question is drawn as "Question <number> of <total>" (see StartQuestions); 0 = the plain old layout (debug tests).
    public static RowSelection? Ask(string filePrefix, string partNumber, out string description, int number = 0, int total = StartQuestions.Total)
    {
        while (true)
        {
            if (number > 0)
            {
                StartQuestions.Header(number, total, "Which Excel rows to run");
            }
            else
            {
                Console.WriteLine("Excel rows to run:");
            }

            Console.WriteLine("   Enter             = all rows");
            Console.WriteLine("   54-70, 72, 100-   = rows 54 to 70, row 72, and row 100 to the end of the sheet");
            Console.WriteLine("   -60               = from the first row up to row 60");
            Console.WriteLine("   a failures file   = only the FAILED rows of an earlier run: drag its .xlsx into this window, or paste its path");
            Console.WriteLine("                       the failures files are in this folder:");
            Console.WriteLine("                       " + RunReport.Dir);
            Console.Write(" Your answer (just press Enter to run all rows): ");
            string? answer = Console.ReadLine();
            string text = (answer ?? "").Trim().Trim('"');
            if (text.Equals("F", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("To run only the failed rows give the failures FILE (drag the .xlsx into this window), not a letter.");
                continue;
            }

            if (text.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                var fromFile = FromFailuresFile(text, filePrefix, partNumber);
                if (fromFile != null)
                {
                    description = "ONLY the failed rows of " + Path.GetFileName(fromFile.SourceFile) + " (" + fromFile._ranges.Count + " row(s): " + fromFile + ")";
                    return fromFile;
                }

                continue;
            }

            if (TryParse(answer, out var selection, out string error))
            {
                description = selection == null ? "all rows" : selection.Describe();
                return selection;
            }

            Console.WriteLine(error);
        }
    }

    // null = the file cannot be used (the reason is printed) and the question is asked again.
    private static RowSelection? FromFailuresFile(string path, string filePrefix, string partNumber)
    {
        if (!File.Exists(path))
        {
            Console.WriteLine("File not found: " + path);
            return null;
        }

        if (!RunReport.TryRead(path, out var chosen, out string readError))
        {
            Console.WriteLine("Cannot use that file: " + readError);
            return null;
        }

        if (chosen!.Unit.Length > 0 && !chosen.Unit.StartsWith(partNumber + " ", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine("That file is from a run on unit '" + chosen.Unit + "' - the Excel rows belong to another unit than " + partNumber + ".");
            return null;
        }

        if (chosen.Failed.Count == 0)
        {
            Console.WriteLine("That file lists no failed rows.");
            return null;
        }

        if (!Path.GetFileName(path).StartsWith(filePrefix + "_", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine("Note: the file is from another test than this one - only its Excel row numbers are used.");
        }

        var selection = new RowSelection { SourceFile = chosen.Path };
        foreach (var row in chosen.Failed)
        {
            selection._ranges.Add((row.Row, row.Row));
            selection._expectedNames[row.Row] = row.Name;
        }

        return selection;
    }

    // Excel rows move when someone edits the master workbook. Compares the parameter name at each row with the name the failures file had.
    // 'found' = (excel row, name now). Returns one text per row that does not match (empty when the rows come from typing, not from a file).
    public List<string> NameMismatches(IEnumerable<(int Row, string Name)> found)
    {
        var result = new List<string>();
        var foundRows = new HashSet<int>();
        foreach (var (row, name) in found)
        {
            foundRows.Add(row);
            if (_expectedNames.TryGetValue(row, out var expected) && expected.Length > 0 && !string.Equals(expected, name, StringComparison.OrdinalIgnoreCase))
            {
                result.Add("row " + row + ": the failures file says '" + expected + "' but the master workbook now has '" + name + "'");
            }
        }

        foreach (var (row, expected) in _expectedNames.Where(e => !foundRows.Contains(e.Key)).Select(e => (e.Key, e.Value)))
        {
            result.Add("row " + row + " ('" + expected + "') from the failures file is not among the parameters this run can check");
        }

        return result;
    }

    public override string ToString() => string.Join(", ", _ranges.Select(r => r.from == r.to ? r.from.ToString()
                                                                                 : r.to == int.MaxValue ? r.from + "-"
                                                                                 : r.from == 0 ? "-" + r.to : r.from + "-" + r.to));
}
