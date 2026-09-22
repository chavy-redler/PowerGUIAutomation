using System.IO;
using PowerGUIAutomation.Tests;

// All test data lives in one Excel workbook - every test class reads its own
// sheet from this same file, so the whole project is driven by the Excel as
// the source of truth.
const string ExcelPath = @"C:\Users\User\OneDrive - Redler Technologies\QA\Power\scripts\Power GUI AU-Tests Master - Operate Detailed.xlsx";

string logPath = Path.Combine(AppContext.BaseDirectory, "GUIAutomationLog_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".txt");

void Log(string line)
{
    string stamped = DateTime.Now.ToString("HH:mm:ss.fff") + "  " + line;
    File.AppendAllText(logPath, stamped + Environment.NewLine);

    if (line.StartsWith("FAIL") || line.StartsWith("STOPPED"))
    {
        Console.ForegroundColor = ConsoleColor.Red;
    }
    else if (line.StartsWith("PASS"))
    {
        Console.ForegroundColor = ConsoleColor.Green;
    }

    Console.WriteLine(stamped);
    Console.ResetColor();
}

Log("SESSION START | PowerGUIAutomation launched");

// Passing the choice as a command-line argument (e.g. "PowerGUIAutomation.exe 1")
// runs that one test non-interactively and exits - handy for scripted runs.
// With no argument, it falls into the normal interactive menu below.
bool nonInteractive = args.Length > 0;
int argIndex = 0;

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
        Console.WriteLine("=== Power GUI Automation ===");
        Console.WriteLine("1. Operate - Second Bar (unit info bar) comparison");
        Console.WriteLine("0. Exit");
        Console.Write("Choice: ");
        choice = Console.ReadLine();
    }

    if (choice == "0" || choice == null)
    {
        Log("SESSION END | user chose to exit");
        break;
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
        default:
            Log("MENU | unknown choice: " + choice);
            break;
    }
}

Console.WriteLine("Log written to: " + logPath);
