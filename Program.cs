using System.IO;
using FlaUI.Core.AutomationElements;
using PowerGUIAutomation;
using PowerGUIAutomation.Tests;

// All test data lives in one Excel workbook - every test class reads its own
// sheet from this same file, so the whole project is driven by the Excel as
// the source of truth.
const string ExcelPath = @"C:\Users\User\OneDrive - Redler Technologies\QA\Power\scripts\Power GUI AU-Tests Master.xlsx";

string logPath = Path.Combine(AppContext.BaseDirectory, "GUIAutomationLog_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt");

void Log(string line)
{
    string stamped = DateTime.Now.ToString("HH:mm:ss.fff") + "  " + line;
    File.AppendAllText(logPath, stamped + Environment.NewLine);   // the file keeps the plain one-line-per-event text
    PrintColored(stamped, line);
}

// Console only: colors by event type, a frame around every failure and around the final RESULT, and a separator
// before every parameter. The tag is the text before the first '|' of the log line.
void Write(string text, ConsoleColor fore, ConsoleColor? back = null)
{
    Console.ForegroundColor = fore;
    if (back.HasValue)
    {
        Console.BackgroundColor = back.Value;
    }

    Console.WriteLine(text);
    Console.ResetColor();
}

void PrintColored(string stamped, string line)
{
    int bar = line.IndexOf('|');
    string tag = (bar > 0 ? line.Substring(0, bar) : line).Trim();
    string time = stamped.Substring(0, 12);
    string rest = bar > 0 ? line.Substring(bar + 1).Trim() : "";

    switch (tag)
    {
        case "DRAW":
            Console.WriteLine();
            Write(new string('-', 100), ConsoleColor.DarkGray);
            Write(time + "  >>> " + rest, ConsoleColor.Black, ConsoleColor.Cyan);
            return;
        case "STEP":
            Write(time + "     [step] " + rest, ConsoleColor.Cyan);
            return;
        case "FAIL" when rest.Contains(" | STEP: "):
        {
            var parts = rest.Split(" | ");
            string Part(string key) => parts.FirstOrDefault(x => x.StartsWith(key))?.Substring(key.Length).Trim() ?? "";
            string who = parts[0];
            Write(time + "  " + new string('!', 94), ConsoleColor.Red);
            Write(time + "  FAILED   " + who, ConsoleColor.White, ConsoleColor.DarkRed);
            Write(time + "    step  : " + Part("STEP:"), ConsoleColor.Yellow);
            Write(time + "    why   : " + Part("WHY:"), ConsoleColor.Red);
            Write(time + "    check : " + Part("CHECK:"), ConsoleColor.DarkYellow);
            Write(time + "  " + new string('!', 94), ConsoleColor.Red);
            return;
        }
        case "RESULT":
            if (rest.StartsWith("====="))
            {
                Write(time + "  " + rest, ConsoleColor.White);
            }
            else if (rest.StartsWith("FAILED"))
            {
                Write(time + "  " + rest + " ", ConsoleColor.White, ConsoleColor.DarkRed);
            }
            else if (rest.StartsWith("PASSED"))
            {
                Write(time + "  " + rest + " ", ConsoleColor.Black, ConsoleColor.Green);
            }
            else
            {
                Write(time + "  " + rest, ConsoleColor.Yellow);
            }

            return;
    }

    ConsoleColor color = tag switch
    {
        "PASS" or "MARK" => ConsoleColor.Green,
        "FAIL" or "STOPPED" => ConsoleColor.Red,
        "WARN" => ConsoleColor.Yellow,
        "INFO" => ConsoleColor.Gray,
        "READ" => ConsoleColor.White,
        "CHANGE" or "SET" or "RESTORE" => ConsoleColor.Yellow,
        "CLICK" => ConsoleColor.Blue,
        "APPLOG" => ConsoleColor.DarkGray,
        "SKIP" or "NOTE" => ConsoleColor.DarkYellow,
        "DRY" => ConsoleColor.DarkCyan,
        "START" or "DONE" => ConsoleColor.Magenta,
        _ => line.StartsWith("===") ? (line.Contains("PASSED") ? ConsoleColor.Green : ConsoleColor.Red) : ConsoleColor.DarkGray,
    };
    Write(stamped, color);
}

Log("SESSION START | PowerGUIAutomation launched");

// The unit under test (chosen by PART NUMBER) decides which Excel sheets the tests read (see UnitProfile.cs), how many
// channels to expect and which unit must be connected. There is no default: the user must pick one.
// Scripted runs pass it as an argument, e.g. "PowerGUIAutomation.exe unit=RD152 3".
UnitProfile? unit = null;

bool SelectUnit(string text)
{
    var wanted = UnitProfile.ParsePartNumber(text);
    var match = wanted == null ? null : UnitProfile.All.FirstOrDefault(u => u.PartNumberValue == wanted);
    if (match == null)
    {
        return false;
    }

    unit = match;
    Log("MENU | unit under test: " + unit.Name + " (sheet '" + unit.ParameterSheet + "')");
    return true;
}

void ChooseUnit()
{
    while (true)
    {
        Console.WriteLine();
        Console.WriteLine("Unit under test (by part number):");
        foreach (var u in UnitProfile.All)
        {
            Console.WriteLine("  " + u.Name);
        }

        Console.Write("Part number (e.g. RD152): ");
        string? text = Console.ReadLine();
        if (text == null)
        {
            return;
        }

        if (SelectUnit(text))
        {
            return;
        }

        Console.WriteLine("Unknown part number '" + text.Trim() + "' - choose one from the list.");
    }
}

// Passing the choice as a command-line argument (e.g. "PowerGUIAutomation.exe 1")
// runs that one test non-interactively and exits - handy for scripted runs.
// With no argument, it falls into the normal interactive menu below.
foreach (var arg in args.Where(x => x.StartsWith("unit=", StringComparison.OrdinalIgnoreCase)))
{
    if (!SelectUnit(arg.Substring(5)))
    {
        Console.WriteLine("Unknown unit '" + arg + "'");
    }
}

args = args.Where(x => !x.StartsWith("unit=", StringComparison.OrdinalIgnoreCase)).ToArray();
bool nonInteractive = args.Length > 0;
int argIndex = 0;
if (!nonInteractive)
{
    ChooseUnit();
}

while (true)
{
    string? choice;
    if (nonInteractive)
    {
        if (argIndex >= args.Length)
        {
            break;
        }

        choice = args[argIndex++];
        Console.WriteLine("Choice: " + choice);
    }
    else
    {
        Console.WriteLine();
        Console.WriteLine("=== Power GUI Automation === | Unit: " + (unit?.Name ?? "(none selected)"));
        Console.WriteLine("1. Operate - Second Bar (unit info bar) comparison");
        Console.WriteLine("2. Configure - Parameters comparison");
        Console.WriteLine("3. Configure - Random parameter change test ");
        Console.WriteLine("U. Change the unit under test");
        Console.WriteLine("0. Exit");
        Console.Write("Choice: ");
        choice = Console.ReadLine();
    }

    if (choice == "0" || choice == null)
    {
        Log("SESSION END | user chose to exit");
        break;
    }

    if (choice.Trim().Equals("U", StringComparison.OrdinalIgnoreCase))
    {
        ChooseUnit();
        continue;
    }

    Log("MENU | choice selected: " + choice);

    switch (choice)
    {
        case "1":
            try
            {
                var test = new OperateSecondBarTest(ExcelPath, Log);
                bool passed = test.Run();
                Log(passed ? "=== TEST PASSED ===" : "=== TEST FAILED ===");
            }
            catch (Exception ex)
            {
                Log("FAIL | Unhandled error: " + ex.Message);
            }
            break;
        case "2":
            try
            {
                if (unit == null) { Log("FAIL | no unit selected - choose the unit (part number) first: press U, or pass unit=RD152"); break; }
                var configureTest = new ConfigureParametersTest(ExcelPath, Log, new[] { unit.ParameterSheet }, unit);
                bool configurePassed = configureTest.Run();
                Log(configurePassed ? "=== TEST PASSED ===" : "=== TEST FAILED ===");
            }
            catch (Exception ex)
            {
                Log("FAIL | Unhandled error: " + ex.Message);
            }
            break;
        case "3":
            try
            {
                if (unit == null) { Log("FAIL | no unit selected - choose the unit (part number) first: press U, or pass unit=RD152"); break; }
                var randomChangeTest = new ConfigureRandomChangeTest(ExcelPath, Log, unit);
                bool randomChangePassed = randomChangeTest.Run();
                Log(randomChangePassed ? "=== TEST PASSED ===" : "=== TEST FAILED ===");
            }
            catch (Exception ex)
            {
                Log("FAIL | Unhandled error (" + ex.GetType().Name + "): " + ex.Message + " | at " + ex.StackTrace?.Split(Environment.NewLine).FirstOrDefault()?.Trim());
            }
            break;
        default:
            Log("MENU | unknown choice: " + choice);
            break;
    }
}

Console.WriteLine("Log written to: " + logPath);
