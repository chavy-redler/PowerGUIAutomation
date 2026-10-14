using PowerGUIAutomation;

namespace PowerGUIAutomation.Tests;

// Debug test: does a parameter that changes fine when set by hand fail when it is changed again and again? Runs the change-and-restore of test 2
// (ConfigureChangeAndRestoreAllParametersTest) on a group of parameters, round after round until the user stops it. If the first failure comes only after many
// writes, the quantity is the problem and not the parameters. WRITES to the unit like test 2.
public class DebugLoopTest
{
    private static readonly TimeSpan TimeLimit = TimeSpan.FromHours(1);      // the loop stops by itself after this long

    private readonly string _excelPath;
    private readonly Action<string> _log;
    private readonly UnitProfile _unit;

    public DebugLoopTest(string excelPath, Action<string> log, UnitProfile unit)
    {
        _excelPath = excelPath;
        _log = log;
        _unit = unit;
    }

    public bool Run()
    {
        Console.WriteLine();
        Console.WriteLine("Loop test: the change-and-restore of test 2 on a group of parameters, round after round (WRITES to the unit).");
        Console.WriteLine("Choose a GROUP of parameters that change well when set by hand: if a failure shows up only after many rounds, the quantity is the cause.");
        Console.WriteLine();
        var selection = RowSelection.Ask("ChangeAndRestore", _unit.PartNumber, out string description);

        Console.WriteLine("The loop runs round after round until you stop it with " + StopControl.KeyName + ", until the first failure (if you choose to stop at a failure), or until " + RunReport.Format(TimeLimit) + " have passed.");
        Console.WriteLine();
        Console.WriteLine("When a parameter fails:");
        Console.WriteLine("  1. stop at the first failure  [default]  (shows exactly after how many writes it happened)");
        Console.WriteLine("  2. continue with the next parameter and round; every failure is listed at the end");
        Console.Write("Choice [1]: ");
        bool stopOnFail = (Console.ReadLine() ?? "").Trim() != "2";

        var test = new ConfigureChangeAndRestoreAllParametersTest(_excelPath, _log, _unit);
        test.UsePreset(selection, description, stopOnFail, rounds: 0, timeLimit: TimeLimit);      // rounds: 0 = no limit on the rounds
        return test.Run();
    }
}
