using System.Runtime.InteropServices;

namespace PowerGUIAutomation;

// A stop key that works while Power Rider Studio is in FRONT. The tests keep the GUI in front on purpose (the keystrokes must reach it), so Ctrl+C
// - which only reaches the console when the console is the front window - cannot be used during a run.
//
// Ctrl+Shift+F12 is registered as a global hot key (the GUI does not see it). In tests that can stop cleanly (3 and 4: "graceful") the first press
// asks the test to stop after the parameter that is being checked (it is put back to its default first), a second press stops at once. In the other
// tests the key stops at once. Stopping at once ends the process: the tests' ProcessExit handler still writes the failures file.
public static class StopControl
{
    public const string KeyName = "Ctrl+Shift+F12";

    private const int HotkeyId = 0x5051;
    private const uint ModControl = 0x2, ModShift = 0x4, ModNoRepeat = 0x4000, VkF12 = 0x7B;
    private const uint WmHotkey = 0x0312, WmQuit = 0x0012;

    [StructLayout(LayoutKind.Sequential)]
    private struct Msg
    {
        public IntPtr Hwnd;
        public uint Message;
        public IntPtr WParam, LParam;
        public uint Time;
        public int PtX, PtY;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint modifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll")]
    private static extern int GetMessage(out Msg msg, IntPtr hWnd, uint min, uint max);

    [DllImport("user32.dll")]
    private static extern bool PostThreadMessage(uint threadId, uint message, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    private static Thread? _thread;
    private static uint _threadId;
    private static volatile bool _requested;
    private static bool _graceful;
    private static Action<string> _log = _ => { };

    // True after the first press in a test that stops cleanly: the test checks it between parameters.
    public static bool Requested => _requested;

    public static void Start(Action<string> log, bool graceful)
    {
        Stop();
        _requested = false;
        _graceful = graceful;
        _log = log;
        using var ready = new ManualResetEventSlim();
        bool registered = false;
        _thread = new Thread(() =>
        {
            _threadId = GetCurrentThreadId();
            registered = RegisterHotKey(IntPtr.Zero, HotkeyId, ModControl | ModShift | ModNoRepeat, VkF12);
            ready.Set();
            if (!registered)
            {
                return;
            }

            while (GetMessage(out var msg, IntPtr.Zero, 0, 0) > 0)
            {
                if (msg.Message == WmHotkey)
                {
                    OnKey();
                }
            }

            UnregisterHotKey(IntPtr.Zero, HotkeyId);
        })
        { IsBackground = true, Name = "stop key" };
        _thread.Start();
        ready.Wait(3000);
        log(registered ? "INFO  | stop key: " + KeyName + (graceful ? " = stop after the parameter that is being checked, press again to stop at once" : " = stop at once")
                       : "WARN  | the stop key " + KeyName + " could not be registered (another program uses it) - only Ctrl+C in the console window can stop the run");
    }

    public static void Stop()
    {
        if (_thread == null)
        {
            return;
        }

        PostThreadMessage(_threadId, WmQuit, IntPtr.Zero, IntPtr.Zero);
        _thread.Join(1000);
        _thread = null;
    }

    private static void OnKey()
    {
        if (_graceful && !_requested)
        {
            _requested = true;
            _log("STOPPED | the stop key was pressed - the test stops after the parameter that is being checked (press " + KeyName + " again to stop at once)");
            return;
        }

        _log("STOPPED | the stop key was pressed - stopping at once");
        Environment.Exit(1);      // the ProcessExit handler of tests 3 and 4 writes the failures file
    }
}
