using System.Text.RegularExpressions;

namespace PowerGUIAutomation;

// One entry per unit (identified by its PART NUMBER) that can be tested. The unit chosen in the main menu tells the
// tests which Excel sheets hold its parameters, how many channels it has, and which unit must be connected.
//
// The GUI shows the part number zero-padded ("85 RD000152"), the list below uses the short form (RD152); both are
// compared by their number.
//
// Excel sheets: the RD152 data keeps the original sheet names. Every other unit uses "<PartNumber> System Parameters" and
// "<PartNumber> Channel Parameters" - create those two sheets in the master workbook (same columns as the RD152 sheets)
// when the unit's parameter list is ready. A unit without its sheets can be selected, but the Configure tests stop with
// a clear message.
//
// To add a unit: add one line to All.
public sealed class UnitProfile
{
    public string PartNumber { get; }          // short form, e.g. "RD152"
    public int Channels { get; }
    public string Note { get; }                // free text shown in the menu, e.g. "Negative"
    public string SystemSheet { get; }
    public string ChannelSheet { get; }
    public int PartNumberValue { get; }        // 152

    public string Name => PartNumber + " - " + Channels + " CH" + (Note.Length > 0 ? " (" + Note + ")" : "");

    public UnitProfile(string partNumber, int channels, string note = "", string? systemSheet = null, string? channelSheet = null)
    {
        PartNumber = partNumber;
        Channels = channels;
        Note = note;
        SystemSheet = systemSheet ?? partNumber + " System Parameters";
        ChannelSheet = channelSheet ?? partNumber + " Channel Parameters";
        PartNumberValue = ParsePartNumber(partNumber) ?? throw new ArgumentException("bad part number " + partNumber);
    }

    // "RD000152", "RD152" and "152" all give 152.
    public static int? ParsePartNumber(string text)
    {
        var m = Regex.Match(text ?? "", @"RD\s*0*(\d+)", RegexOptions.IgnoreCase);
        if (!m.Success)
        {
            m = Regex.Match(text ?? "", @"^\s*0*(\d+)\s*$");
        }

        return m.Success && int.TryParse(m.Groups[1].Value, out int value) ? value : null;
    }

    public static readonly IReadOnlyList<UnitProfile> All = new List<UnitProfile>
    {
        new("RD152", 16, systemSheet: "System Parameters", channelSheet: "Channel Parameters"),   // the data that exists today
        new("RD323", 16),
        new("RD320", 16),
        new("RD249", 16),
        new("RD279", 12),
        new("RD336", 12, "Negative"),
        new("RD85", 1),
        new("RD250", 1),
    };

    public override string ToString() => Name;
}
