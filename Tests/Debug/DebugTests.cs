using PowerGUIAutomation;

namespace PowerGUIAutomation.Tests;

// Menu 5: the DEBUG tests (all read-only). Each one is its own class whose name starts with "Debug":
//   1. DebugCompareSubsetTest      - compare parameters to the Excel on chosen channels / parameter names only
//   2. DebugAutomationIdScanTest   - AutomationId of every parameter in the live GUI vs the Excel -> report workbook in the Excel folder
public class DebugTests
{
    private readonly string _excelPath;
    private readonly Action<string> _log;
    private readonly UnitProfile _unit;

    public DebugTests(string excelPath, Action<string> log, UnitProfile unit)
    {
        _excelPath = excelPath;
        _log = log;
        _unit = unit;
    }

    public bool Run()
    {
        Console.WriteLine();
        Console.WriteLine("Debug tests (read-only):");
        Console.WriteLine("  1. DebugCompareSubsetTest    - compare parameters to the Excel on chosen channels / parameters");
        Console.WriteLine("  2. DebugAutomationIdScanTest - AutomationId scan of all parameters (report to a new Excel file)");
        Console.Write("Choice [1]: ");
        string pick = (Console.ReadLine() ?? "").Trim();
        if (pick == "2")
        {
            _log("START | DebugAutomationIdScanTest | unit=" + _unit.Name);
            return new DebugAutomationIdScanTest(_excelPath, _log, _unit).Run();
        }

        _log("START | DebugCompareSubsetTest | unit=" + _unit.Name);
        return new DebugCompareSubsetTest(_excelPath, _log, _unit).Run();
    }
}
