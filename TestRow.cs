namespace PowerGUIAutomation;

// One row of test data read from the Excel. MatchType decides whether we
// compare against ExpectedValue exactly, or check that every sample falls
// within [Min, Max] - e.g. Refresh Rate constantly changes, so the Excel
// author marks it "Range" instead of pretending it has one fixed value.
public class TestRow
{
    public string Entity { get; set; } = "";
    public int Index { get; set; }
    public string ExpectedValue { get; set; } = "";
    public string MatchType { get; set; } = "Exact"; // "Exact" or "Range"
    public double Min { get; set; }
    public double Max { get; set; }
    public int SampleCount { get; set; } = 1;
    public double SampleIntervalSeconds { get; set; } = 0;
}
