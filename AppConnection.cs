using System.Diagnostics;
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
    public static Window Attach(UIA3Automation automation)
    {
        var processes = Process.GetProcessesByName("PowerRiderStudio");
        if (processes.Length == 0)
        {
            throw new InvalidOperationException("Power Rider Studio is not running. Start it first.");
        }

        var app = Application.Attach(processes[0]);
        return app.GetMainWindow(automation);
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
