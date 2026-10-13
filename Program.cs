using System.IO;
using FlaUI.Core.AutomationElements;
using PowerGUIAutomation;
using PowerGUIAutomation.Tests;

// All test data lives in one Excel workbook - every test class reads its own
// sheet from this same file, so the whole project is driven by the Excel as
// the source of truth.
const string ExcelPath = @"C:\Users\User\OneDrive - Redler Technologies\QA\Power\scripts\PowerGUIAutomation\Power GUI AU-Tests Master.xlsx";

string logPath = Path.Combine(AppContext.BaseDirectory, "GUIAutomationLog_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt");

void Log(string line)
{
    string stamped = DateTime.Now.ToString("HH:mm:ss.fff") + "  " + line;
    if (line.StartsWith("DRAW", StringComparison.Ordinal))
    {
        // the log file gets the same separator between two parameters as the console
        File.AppendAllText(logPath, Environment.NewLine + new string('-', 100) + Environment.NewLine, new System.Text.UTF8Encoding(true));
    }

    File.AppendAllText(logPath, stamped + Environment.NewLine, new System.Text.UTF8Encoding(true));   // the file keeps the plain one-line-per-event text
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
            Write(time + "  >>> " + rest, ConsoleColor.Cyan);
            return;
        case "STEP":
            Write(time + "     [step] " + rest, ConsoleColor.DarkGray);
            return;
        case "FAIL" when rest.Contains(" | STEP: "):
        {
            var parts = rest.Split(" | ");
            string Part(string key) => parts.FirstOrDefault(x => x.StartsWith(key))?.Substring(key.Length).Trim() ?? "";
            string who = parts[0];
            Write(time + "  " + new string('!', 94), ConsoleColor.Red);
            Write(time + "  FAILED   " + who, ConsoleColor.Red);
            Write(time + "    step  : " + Part("STEP:"), ConsoleColor.Red);
            Write(time + "    why   : " + Part("WHY:"), ConsoleColor.Red);
            Write(time + "    check : " + Part("CHECK:"), ConsoleColor.Red);
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
                Write(time + "  " + rest, ConsoleColor.Red);
            }
            else if (rest.StartsWith("PASSED"))
            {
                Write(time + "  " + rest, ConsoleColor.Green);
            }
            else
            {
                Write(time + "  " + rest, ConsoleColor.Gray);
            }

            return;
    }

    // Few colors on purpose: green = pass, red = fail, yellow = warning/skip/note, cyan = the parameter that is being checked,
    // dark gray = the step headers and the GUI's own log lines, everything else the normal gray.
    ConsoleColor color = tag switch
    {
        "PASS" or "MARK" => ConsoleColor.Green,
        "FAIL" or "STOPPED" => ConsoleColor.Red,
        "WARN" or "SKIP" or "NOTE" => ConsoleColor.Yellow,
        "APPLOG" => ConsoleColor.DarkGray,
        "RESUME" => ConsoleColor.Cyan,
        _ => line.StartsWith("===") ? (line.Contains("PASSED") ? ConsoleColor.Green : ConsoleColor.Red) : ConsoleColor.Gray,
    };
    Write(stamped, color);
}

Log("SESSION START | PowerGUIAutomation launched");

// Boxed header printed when a test is chosen (console, cyan) and written as plain lines to the log file, so the start of every
// test is easy to find when searching the logs afterwards.
void ShowTestBanner(string? choice, UnitProfile? bannerUnit)
{
    string? title = choice?.Trim() switch
    {
        "1" => "OPERATE - SECOND BAR (UNIT INFO BAR) COMPARISON",
        "2" => "CONFIGURE - PARAMETERS COMPARISON",
        "3" => "CONFIGURE - CHANGE AND RESTORE EVERY PARAMETER",
        "4" => "CONFIGURE - INPUT VALIDATION (INVALID VALUES MUST BE REJECTED)",
        "5" => "DEBUG TESTS (READ-ONLY)",
        _ => null,
    };
    if (title == null)
    {
        return;
    }

    string unitText = "Unit: " + (bannerUnit?.Name ?? "(none selected)") + "   |   " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
    int width = Math.Max(Math.Max(title.Length, unitText.Length) + 6, 64);
    string Line(string t) => "║  " + t.PadRight(width - 4) + "║";
    string[] box =
    {
        "╔" + new string('═', width - 2) + "╗",
        Line("TEST RUNNER: " + title),
        Line(unitText),
        "╚" + new string('═', width - 2) + "╝",
    };

    Console.OutputEncoding = System.Text.Encoding.UTF8;
    Console.WriteLine();
    foreach (var l in box)
    {
        Write(l, ConsoleColor.Cyan);
    }

    Console.WriteLine();
    // the same box goes into the log file (UTF-8 with BOM so the box characters display correctly in Notepad), plus a greppable line
    File.AppendAllText(logPath, Environment.NewLine + string.Join(Environment.NewLine, box) + Environment.NewLine
                       + "TEST SELECTED: " + title + " | " + unitText + Environment.NewLine + Environment.NewLine, new System.Text.UTF8Encoding(true));
}

// The "before you start" frame (TestInstructions.cs) for the chosen test: console in color, log file as plain lines.
void ShowInstructions(string choice, UnitProfile? instructionUnit)
{
    string sheetNote = instructionUnit == null ? "" : "Its sheet '" + instructionUnit.ParameterSheet + "' must exist and be valid.";
    var lines = TestInstructions.Build(choice, instructionUnit, sheetNote);
    Console.OutputEncoding = System.Text.Encoding.UTF8;
    foreach (var (text, color) in lines)
    {
        Write(text, color);
    }

    Console.WriteLine();
    File.AppendAllText(logPath, string.Join(Environment.NewLine, lines.Select(l => l.Text)) + Environment.NewLine + Environment.NewLine, new System.Text.UTF8Encoding(true));
}

// Last line of every test: how long it took (RunClock: from the real start, the menu and the questions are not counted). Plain ruler line with no
// "TAG |" and no time stamp, so it stands out when the .txt log is scanned afterwards.
void ShowTookLine(string choice, UnitProfile? tookUnit)
{
    if (!RunClock.Running)
    {
        return;
    }

    string text = "  TEST " + choice + (tookUnit != null ? " (" + tookUnit.PartNumber + ")" : "") + " TOOK " + RunReport.Format(RunClock.Elapsed)
                  + "   (" + RunClock.StartedAt.ToString("HH:mm:ss") + " - " + DateTime.Now.ToString("HH:mm:ss") + ")  ";
    string line = "════════" + text + new string('═', 8);
    Console.WriteLine();
    Write(line, ConsoleColor.Cyan);
    File.AppendAllText(logPath, Environment.NewLine + line + Environment.NewLine, new System.Text.UTF8Encoding(true));
}

// The master workbook must exist at ExcelPath and be CLOSED in Excel while a test runs (an open workbook is locked / may be saved over).
bool MasterOk()
{
    if (!File.Exists(ExcelPath))
    {
        Log("FAIL | the master workbook was not found at: " + ExcelPath + " - fix ExcelPath in Program.cs (or move the file back)");
        return false;
    }

    try
    {
        using var fs = new FileStream(ExcelPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);   // fails while Excel has it open
        return true;
    }
    catch (IOException)
    {
        Log("FAIL | the master workbook is OPEN (in Excel?) - close it and run again: " + ExcelPath);
        return false;
    }
}

// Characters typed while a test ran can end up in the console input buffer; they must not be read as answers of the next menu question.
void FlushConsoleInput()
{
    try
    {
        while (Console.KeyAvailable)
        {
            Console.ReadKey(true);
        }
    }
    catch (InvalidOperationException)
    {
        // input is redirected (scripted run) - nothing to flush
    }
}

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
        Console.WriteLine("3. Configure - Change and restore every parameter (writes to the unit)");
        Console.WriteLine("4. Configure - Input validation (types invalid values, never saves to flash)");
        Console.WriteLine("5. Debug tests (read-only): Configure comparison on chosen channels / parameters, AutomationId scan");
        Console.WriteLine("U. Change the unit under test");
        Console.WriteLine("0. Exit");
        FlushConsoleInput();
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
    ShowTestBanner(choice, unit);
    RunClock.Reset();
    if (choice.Trim() is "1" or "2" or "3" or "4" or "5")
    {
        ShowInstructions(choice.Trim(), unit);
        if (!nonInteractive)
        {
            Console.Write("Press Enter to start the test (Q = back to the menu): ");
            if (string.Equals(Console.ReadLine()?.Trim(), "Q", StringComparison.OrdinalIgnoreCase))
            {
                Log("MENU | test not started - back to the menu");
                continue;
            }
        }
    }

    switch (choice)
    {
        case "1":
            try
            {
                if (!MasterOk()) { break; }
                RunClock.Start();
                var test = new OperateSecondBarTest(ExcelPath, Log);
                bool passed = test.Run();
                Log(passed ? "=== TEST PASSED ===" : "=== TEST FAILED ===");
            }
            catch (Exception ex)
            {
                Log("FAIL | Unhandled error (" + ex.GetType().Name + "): " + ex.Message + " | at " + ex.StackTrace?.Split(Environment.NewLine).FirstOrDefault()?.Trim());
            }
            break;
        case "2":
            try
            {
                if (unit == null) { Log("FAIL | no unit selected - choose the unit (part number) first: press U, or pass unit=RD152"); break; }
                if (!MasterOk()) { break; }
                RunClock.Start();
                var configureTest = new ConfigureCompareParametersToExcelTest(ExcelPath, Log, new[] { unit.ParameterSheet }, unit);
                bool configurePassed = configureTest.Run();
                Log(configurePassed ? "=== TEST PASSED ===" : "=== TEST FAILED ===");
            }
            catch (Exception ex)
            {
                Log("FAIL | Unhandled error (" + ex.GetType().Name + "): " + ex.Message + " | at " + ex.StackTrace?.Split(Environment.NewLine).FirstOrDefault()?.Trim());
            }
            break;
        case "3":
            try
            {
                if (unit == null) { Log("FAIL | no unit selected - choose the unit (part number) first: press U, or pass unit=RD152"); break; }
                if (!MasterOk()) { break; }
                RunClock.Start();
                var randomChangeTest = new ConfigureChangeAndRestoreAllParametersTest(ExcelPath, Log, unit);
                bool randomChangePassed = randomChangeTest.Run();
                Log(randomChangePassed ? "=== TEST PASSED ===" : "=== TEST FAILED ===");
            }
            catch (Exception ex)
            {
                Log("FAIL | Unhandled error (" + ex.GetType().Name + "): " + ex.Message + " | at " + ex.StackTrace?.Split(Environment.NewLine).FirstOrDefault()?.Trim());
            }
            break;
        case "4":
            try
            {
                if (unit == null) { Log("FAIL | no unit selected - choose the unit (part number) first: press U, or pass unit=RD152"); break; }
                if (!MasterOk()) { break; }
                RunClock.Start();
                var validationTest = new ConfigureInputValidationTest(ExcelPath, Log, unit);
                bool validationPassed = validationTest.Run();
                Log(validationPassed ? "=== TEST PASSED ===" : "=== TEST FAILED ===");
            }
            catch (Exception ex)
            {
                Log("FAIL | Unhandled error (" + ex.GetType().Name + "): " + ex.Message + " | at " + ex.StackTrace?.Split(Environment.NewLine).FirstOrDefault()?.Trim());
            }
            break;
        case "5":
            try
            {
                if (unit == null) { Log("FAIL | no unit selected - choose the unit (part number) first: press U, or pass unit=RD152"); break; }
                if (!MasterOk()) { break; }
                RunClock.Start();
                var debugTest = new DebugTests(ExcelPath, Log, unit);
                bool debugPassed = debugTest.Run();
                Log(debugPassed ? "=== TEST PASSED ===" : "=== TEST FAILED ===");
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

    if (choice.Trim() is "1" or "2" or "3" or "4" or "5")
    {
        ShowTookLine(choice.Trim(), unit);
    }
}

Console.WriteLine("Log written to: " + logPath);
