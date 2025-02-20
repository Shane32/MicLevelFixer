using System;
using System.Windows.Forms;

namespace MicLevelFixer;

/// <summary>
/// Main entry point for the .NET Framework 4.8 tray app.
/// </summary>
internal static class Program
{
    [STAThread]
    internal static void Main()
    {
        // Standard WinForms setup for .NET 4.8
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        // Run with a custom ApplicationContext (TrayAppContext)
        var trayContext = new TrayAppContext();
        Application.Run(trayContext);
    }
}
