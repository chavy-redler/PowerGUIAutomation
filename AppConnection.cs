using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;

namespace PowerGUIAutomation;

// Shared "attach to the running Power Rider Studio and switch tabs" logic -
// every test class needs this, so it lives in one place instead of being
// copy-pasted per test.
public static class AppConnection
{
    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    private static int _guiProcessId;

    // Keyboard input (paste, Enter) goes to the window that is in FRONT. At the start of a run that is the console, so the GUI is brought forward here and
    // re-checked before every typed input (two Win32 calls, no UIA - practically free).
    public static bool IsGuiInFront()
    {
        GetWindowThreadProcessId(GetForegroundWindow(), out uint pid);
        return _guiProcessId != 0 && pid == (uint)_guiProcessId;
    }

    public static bool EnsureGuiInFront(Window window)
    {
        if (IsGuiInFront())
        {
            return true;
        }

        try
        {
            var pattern = window.Patterns.Window.PatternOrDefault;
            if (pattern != null && pattern.WindowVisualState.ValueOrDefault == FlaUI.Core.Definitions.WindowVisualState.Minimized)
            {
                pattern.SetWindowVisualState(FlaUI.Core.Definitions.WindowVisualState.Normal);
            }

            window.SetForeground();
        }
        catch
        {
            // checked below
        }

        for (int i = 0; i < 8; i++)
        {
            Thread.Sleep(100);
            if (IsGuiInFront())
            {
                return true;
            }
        }

        return false;
    }

    // The red X / green V are read from a screenshot (RowIcons): that needs Windows display scaling of 100% (the registry value is the DPI of the primary screen).
    private static void WarnIfScaled()
    {
        try
        {
            if (Microsoft.Win32.Registry.GetValue(@"HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "AppliedDPI", 96) is int dpi && dpi != 96)
            {
                Console.WriteLine("WARNING: Windows display scaling is " + (dpi * 100 / 96) + "%, not 100% - the red X / green V icons may not be found. Set Settings > Display > Scale to 100%.");
            }
        }
        catch
        {
            // the setting cannot be read: nothing to warn about
        }
    }

    public static Window Attach(UIA3Automation automation)
    {
        WarnIfScaled();
        var processes = Process.GetProcessesByName("PowerRiderStudio");
        if (processes.Length == 0)
        {
            throw new InvalidOperationException("Power Rider Studio is not running. Start it first.");
        }

        var app = Application.Attach(processes[0]);
        _guiProcessId = processes[0].Id;
        var window = app.GetMainWindow(automation);
        EnsureGuiInFront(window);       // the console is in front when a test starts
        return window;
    }

    public static void GoToTab(Window mainWindow, string tabAutomationId)
    {
        var tab = mainWindow.FindFirstDescendant(cf => cf.ByAutomationId(tabAutomationId));
        if (tab == null)
        {
            throw new InvalidOperationException($"Could not find tab button '{tabAutomationId}'.");
        }

        tab.AsButton().Invoke();
        Thread.Sleep(1500);
    }
}
