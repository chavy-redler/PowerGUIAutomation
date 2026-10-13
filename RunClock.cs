using System.Diagnostics;

namespace PowerGUIAutomation;

// The "how long did the test take" clock of the main menu. It is started when the test really starts - after the menu choice, the instructions,
// the confirmations and the questions of the test - so the time the user spent answering is NOT counted. Program.cs starts it just before the
// test is created; a test that asks its own questions (3, 4, debug) starts it again after the last answer.
public static class RunClock
{
    private static Stopwatch? _watch;
    private static DateTime _startedAt;

    public static bool Running => _watch != null;
    public static DateTime StartedAt => _startedAt;
    public static TimeSpan Elapsed => _watch?.Elapsed ?? TimeSpan.Zero;

    public static void Start()
    {
        _startedAt = DateTime.Now;
        _watch = Stopwatch.StartNew();
    }

    public static void Reset() => _watch = null;
}
