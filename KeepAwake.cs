using System.Runtime.InteropServices;

namespace PowerGUIAutomation;

// A long (overnight) run must not be interrupted by Windows going to sleep or turning the display off: the tests type into the GUI and read a screenshot.
// While a test runs the system and the display are kept awake. This cannot stop a locked screen (lock screen / screen saver with a password): disable those
// before an overnight run. Called from the main thread (the state belongs to the calling thread).
public static class KeepAwake
{
    private const uint EsContinuous = 0x80000000, EsSystemRequired = 0x00000001, EsDisplayRequired = 0x00000002;

    [DllImport("kernel32.dll")]
    private static extern uint SetThreadExecutionState(uint flags);

    public static void On() => SetThreadExecutionState(EsContinuous | EsSystemRequired | EsDisplayRequired);

    public static void Off() => SetThreadExecutionState(EsContinuous);
}
