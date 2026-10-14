using System.Globalization;

namespace PowerGUIAutomation;

// The questions a test asks before it runs, kept in one place so that tests 2 and 3 (and every later test) ask them the same way: every question is a block
// with a cyan ruler and a title ("Question 2 of 3 - ..."), the answers listed one per line, and the prompt on its own line.
public static class StartQuestions
{
    public const int Total = 3;      // failure mode, rows, loop

    private const int RulerWidth = 100;

    // The ruler + title of question number/total.
    public static void Header(int number, int total, string title)
    {
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine(new string('─', RulerWidth));
        Console.WriteLine(" Question " + number + " of " + total + " - " + title);
        Console.ResetColor();
    }

    // Question 1: stop at the first failure (default) or continue. 'unit' = what fails ("input", "parameter").
    public static bool AskStopOnFail(string unit)
    {
        Header(1, Total, "When " + (unit == "input" ? "an input" : "a parameter") + " fails");
        Console.WriteLine("   1 = stop at the first failure  [default]");
        Console.WriteLine("   2 = continue" + (unit == "parameter" ? " with the next parameter" : "") + "; every failure is listed at the end (and in the Failures folder)");
        Console.Write(" Choice [1]: ");
        return (Console.ReadLine() ?? "").Trim() != "2";
    }

    // Question 3: the loop. rounds = 1 -> run once. unlimited = round after round until the stop key. limit = optional time limit.
    public static void AskLoop(out int rounds, out bool unlimited, out TimeSpan? limit)
    {
        rounds = 1;
        unlimited = false;
        limit = null;
        Header(3, Total, "Loop (repeat the whole run)");
        Console.WriteLine("   Enter = run once");
        Console.WriteLine("   N     = N rounds");
        Console.WriteLine("   0     = round after round until you stop it (" + StopControl.KeyName + ")");
        Console.Write(" Loop [Enter = once]: ");
        string loopText = (Console.ReadLine() ?? "").Trim();
        if (int.TryParse(loopText, out int loopRounds) && loopRounds != 1)
        {
            unlimited = loopRounds <= 0;
            rounds = Math.Max(1, loopRounds);
            Console.Write(" Stop automatically after how many hours? [Enter = no limit]: ");
            if (double.TryParse((Console.ReadLine() ?? "").Trim().Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out double hours) && hours > 0)
            {
                limit = TimeSpan.FromHours(hours);
            }
        }

        Console.WriteLine(new string('─', RulerWidth));
        Console.WriteLine();
    }
}
