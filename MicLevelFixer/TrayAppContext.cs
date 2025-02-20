using System;
using System.Windows.Forms;

namespace MicLevelFixer
{
    /// <summary>
    /// The tray application context: creates the NotifyIcon, context menu, timer, etc.
    /// </summary>
    public class TrayAppContext : ApplicationContext
    {
        private NotifyIcon _trayIcon;
        private Timer _timer;
        private int _intervalSeconds;

        public TrayAppContext()
        {
            // 1. Load user settings from registry
            UserSettings.LoadFromRegistry();
            _intervalSeconds = UserSettings.CheckIntervalSeconds; // default 60 if not found

            // 2. Create the tray icon & context menu
            _trayIcon = new NotifyIcon {
                // If you have an embedded icon, use it here. Otherwise, set Visible=true and remove 'Icon' if no icon is available.
                Icon = Properties.Resources.AppIcon,
                Text = "Microphone Level Enforcer",
                Visible = true,
                ContextMenuStrip = BuildContextMenu()
            };

            // 3. Create and start a timer
            _timer = new Timer();
            _timer.Interval = _intervalSeconds * 1000; // ms
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        private ContextMenuStrip BuildContextMenu()
        {
            var menu = new ContextMenuStrip();

            var configureItem = new ToolStripMenuItem("Configure", null, OnConfigureClicked);
            var aboutItem = new ToolStripMenuItem("About", null, OnAboutClicked);
            var exitItem = new ToolStripMenuItem("Exit", null, OnExitClicked);

            menu.Items.Add(configureItem);
            menu.Items.Add(aboutItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(exitItem);

            return menu;
        }

        private void OnConfigureClicked(object sender, EventArgs e)
        {
            using (var form = new OptionsForm()) {
                // Show config dialog. If user presses OK, re-load settings in case they changed.
                if (form.ShowDialog() == DialogResult.OK) {
                    UserSettings.LoadFromRegistry();
                    _intervalSeconds = UserSettings.CheckIntervalSeconds;
                    _timer.Interval = _intervalSeconds * 1000;
                }
            }
        }

        private void OnAboutClicked(object sender, EventArgs e)
        {
            MessageBox.Show(
                "MicLevelFixer\n" +
                "Version 1.0\n\n" +
                "SACK Corporation",
                "About",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void OnExitClicked(object sender, EventArgs e)
        {
            // Clean up tray icon before exit
            _trayIcon.Visible = false;
            _trayIcon.Dispose();

            Application.Exit();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            // Periodically enforce the configured mic volumes
            VolumeEnforcer.EnforceConfiguredVolumes();
        }
    }
}
