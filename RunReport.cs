using System.Diagnostics;
using ClosedXML.Excel;

namespace PowerGUIAutomation;

// What the failures workbooks of tests 3 and 4 have in common (Failures\<Test>_<PN>_<run start>.xlsx next to the exe).
//
// The file name is fixed when the run starts, so the log can name it and the same file is REWRITTEN: once at the start (Status "Running"),
// after every failure, and at the end - whatever the end is (finished, stopped at a failure, stuck, Ctrl+C, closed console, crash). A run that
// is killed hard still leaves the last version. Sheets: "Failed" (written by the test; column A = Excel row, so a later run can take it),
// "Run" (status, time, mode) and "Not run" (rows that were never reached).
public sealed class RunReport
{
    public const string SheetFailed = "Failed";
    public const string SheetRun = "Run";
    public const string SheetNotRun = "Not run";
    public const string RunningStatus = "Running - not finished (the run ended without a summary: killed or power loss?)";

    private readonly Stopwatch _watch = Stopwatch.StartNew();
    public readonly object Sync = new();

    public string Path { get; }
    public DateTime Started { get; } = DateTime.Now;
    public string Status { get; private set; } = RunningStatus;
    public bool Finished { get; private set; }
    public DateTime? Ended { get; private set; }
    public TimeSpan Elapsed => _watch.Elapsed;

    // the file is named <prefix>_<PN>_<run start>.xlsx, e.g. ChangeAndRestore_RD152_20261012_153012.xlsx
    public RunReport(string prefix, string partNumber)
    {
        Directory.CreateDirectory(Dir);
        Path = System.IO.Path.Combine(Dir, prefix + "_" + partNumber + "_" + Started.ToString("yyyyMMdd_HHmmss") + ".xlsx");
    }

    public static string Dir => System.IO.Path.Combine(AppContext.BaseDirectory, "Failures");

    // "Completed" etc. Once finished the status is not changed again (a Ctrl+C after the summary must not overwrite it).
    public bool Finish(string status)
    {
        lock (Sync)
        {
            if (Finished)
            {
                return false;
            }

            Finished = true;
            Status = status;
            Ended = DateTime.Now;
            _watch.Stop();
            return true;
        }
    }

    public static string Format(TimeSpan t)
    {
        if (t.TotalHours >= 1)
        {
            return (int)t.TotalHours + " h " + t.Minutes.ToString("00") + " min " + t.Seconds.ToString("00") + " s";
        }

        return t.TotalMinutes >= 1 ? t.Minutes + " min " + t.Seconds.ToString("00") + " s" : t.Seconds + " s";
    }

    // "Run" sheet: one line per fact (the test adds its own counters in 'extra').
    public void AddRunSheet(XLWorkbook wb, IEnumerable<(string Key, string Value)> extra)
    {
        var ws = wb.AddWorksheet(SheetRun);
        var lines = new List<(string, string)>
        {
            ("Status", Status),
            ("Started", Started.ToString("yyyy-MM-dd HH:mm:ss")),
            ("Ended", Ended?.ToString("yyyy-MM-dd HH:mm:ss") ?? "(not finished)"),
            ("Run time", Format(Elapsed)),
        };
        lines.AddRange(extra.Select(e => (e.Key, e.Value)));
        for (int i = 0; i < lines.Count; i++)
        {
            ws.Cell(i + 1, 1).Value = lines[i].Item1;
            ws.Cell(i + 1, 1).Style.Font.Bold = true;
            ws.Cell(i + 1, 2).Value = lines[i].Item2;
        }

        ws.Columns().AdjustToContents();
    }

    // "Not run" sheet: parameters of the selection that the run never reached (only when there are some).
    public static void AddNotRunSheet(XLWorkbook wb, IReadOnlyList<(int Row, string Channel, string Name)> notRun)
    {
        if (notRun.Count == 0)
        {
            return;
        }

        var ws = wb.AddWorksheet(SheetNotRun);
        string[] head = { "Excel row", "Channel", "Parameter" };
        for (int i = 0; i < head.Length; i++)
        {
            ws.Cell(1, i + 1).Value = head[i];
            ws.Cell(1, i + 1).Style.Font.Bold = true;
        }

        for (int i = 0; i < notRun.Count; i++)
        {
            ws.Cell(i + 2, 1).Value = notRun[i].Row;
            ws.Cell(i + 2, 2).Value = notRun[i].Channel;
            ws.Cell(i + 2, 3).Value = notRun[i].Name;
        }

        ws.Columns().AdjustToContents();
    }

    // Saves through a temp file so an interrupted save never leaves a broken workbook behind.
    public void Save(XLWorkbook wb)
    {
        string tmp = System.IO.Path.Combine(Dir, "~" + System.IO.Path.GetFileName(Path));      // ClosedXML only saves .xlsx names; "~" keeps it out of Find()
        wb.SaveAs(tmp);
        File.Move(tmp, Path, true);
    }

    // ------------------------------------------------------------------ an earlier run's file as the list of rows to run

    public sealed record PreviousRun(string Path, string Status, string RunTime, string Unit, IReadOnlyList<(int Row, string Channel, string Name)> Failed);

    // Reads the "Failed" sheet (column A = Excel row, B = channel, C = parameter) and the "Run" sheet. Works on a copy: the file may be open in Excel.
    public static bool TryRead(string path, out PreviousRun? info, out string error)
    {
        info = null;
        error = "";
        string temp = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "failures_" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            File.Copy(path, temp, true);
            using var wb = new XLWorkbook(temp);
            if (!wb.TryGetWorksheet(SheetFailed, out var failed))
            {
                error = "the workbook has no '" + SheetFailed + "' sheet - not a failures file of this program";
                return false;
            }

            // a parameter can have several lines (one per failing input): every Excel row counts once
            var rows = new List<(int, string, string)>();
            for (int r = 2; r <= failed.LastRowUsed()?.RowNumber(); r++)
            {
                if (failed.Cell(r, 1).TryGetValue(out int excelRow) && rows.All(x => x.Item1 != excelRow))
                {
                    rows.Add((excelRow, failed.Cell(r, 2).GetString(), failed.Cell(r, 3).GetString()));
                }
            }

            string status = "", runTime = "", unit = "";
            if (wb.TryGetWorksheet(SheetRun, out var run))
            {
                for (int r = 1; r <= (run.LastRowUsed()?.RowNumber() ?? 0); r++)
                {
                    string key = run.Cell(r, 1).GetString();
                    if (key == "Status")
                    {
                        status = run.Cell(r, 2).GetString();
                    }
                    else if (key == "Run time")
                    {
                        runTime = run.Cell(r, 2).GetString();
                    }
                    else if (key == "Unit")
                    {
                        unit = run.Cell(r, 2).GetString();
                    }
                }
            }

            info = new PreviousRun(path, status, runTime, unit, rows);
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
        finally
        {
            try { File.Delete(temp); } catch { /* temp copy */ }
        }
    }
}
