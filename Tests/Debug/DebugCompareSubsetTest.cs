using PowerGUIAutomation;

namespace PowerGUIAutomation.Tests;

// DEBUG test (read-only): runs the "Configure compare parameters to Excel" test on a chosen SUBSET of the unit's Excel rows - e.g. only
// channel 15, or channels 13 and 15, or one parameter name - so a suspicion can be checked in minutes instead of a full run.
// Which rows run is decided by the questions it asks (nothing is hard-coded here).
public class DebugCompareSubsetTest
{
    private readonly string _excelPath;
    private readonly Action<string> _log;
    private readonly UnitProfile _unit;

    public DebugCompareSubsetTest(string excelPath, Action<string> log, UnitProfile unit)
    {
        _excelPath = excelPath;
        _log = log;
        _unit = unit;
    }

    public bool Run()
    {
        Console.WriteLine();
        Console.WriteLine("Debug run: Configure comparison on a subset of the rows of '" + _unit.ParameterSheet + "' (read-only).");
        Console.Write("Channels (e.g. 15  or  13,15  or  S,15 where S = System parameters; Enter = all): ");
        string channelsText = (Console.ReadLine() ?? "").Trim();
        Console.Write("Parameter name contains (Enter = any): ");
        string nameText = (Console.ReadLine() ?? "").Trim();

        HashSet<int>? channels = null;
        if (channelsText.Length > 0)
        {
            channels = new HashSet<int>();
            foreach (var token in channelsText.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (token.Equals("S", StringComparison.OrdinalIgnoreCase) || token.Equals("System", StringComparison.OrdinalIgnoreCase))
                {
                    channels.Add(0);
                }
                else if (int.TryParse(token, out int c) && c >= 1 && c <= _unit.Channels)
                {
                    channels.Add(c);
                }
                else
                {
                    _log("FAIL  | bad channel '" + token + "' (valid: S, 1.." + _unit.Channels + ")");
                    return false;
                }
            }
        }

        _log("START | Debug run | unit=" + _unit.Name + " | channels=" + (channels == null ? "all" : string.Join(",", channels.OrderBy(x => x).Select(x => x == 0 ? "System" : x.ToString())))
             + " | name contains='" + nameText + "'");

        var test = new ConfigureCompareParametersToExcelTest(_excelPath, _log, new[] { _unit.ParameterSheet }, _unit,
            (channel, name) => (channels == null || channels.Contains(channel))
                               && (nameText.Length == 0 || name.Contains(nameText, StringComparison.OrdinalIgnoreCase)));
        return test.Run();
    }
}
