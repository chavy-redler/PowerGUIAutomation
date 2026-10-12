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

    public bool Contains(int row) => _ranges.Any(r => row >= r.from && row <= r.to);

    public override string ToString() => string.Join(", ", _ranges.Select(r => r.from == r.to ? r.from.ToString()
                                                                                 : r.to == int.MaxValue ? r.from + "-"
                                                                                 : r.from == 0 ? "-" + r.to : r.from + "-" + r.to));
}
