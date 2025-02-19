using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using AudioSwitcher.AudioApi;
using AudioSwitcher.AudioApi.CoreAudio;

namespace MicLevelFixer
{
    /// <summary>
    /// Main entry point for the .NET Framework 4.8 tray app.
    /// </summary>
    static class Program
    {
        [STAThread]
        static void Main()
        {
            // Standard WinForms setup for .NET 4.8
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Run with a custom ApplicationContext (TrayAppContext)
            var trayContext = new TrayAppContext();
            Application.Run(trayContext);
        }
    }

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
            _trayIcon = new NotifyIcon
            {
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
            using (var form = new OptionsForm())
            {
                // Show config dialog. If user presses OK, re-load settings in case they changed.
                if (form.ShowDialog() == DialogResult.OK)
                {
                    UserSettings.LoadFromRegistry();
                    _intervalSeconds = UserSettings.CheckIntervalSeconds;
                    _timer.Interval = _intervalSeconds * 1000;
                }
            }
        }

        private void OnAboutClicked(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Microphone Level Enforcer\n" +
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

    /// <summary>
    /// A form to configure the microphone levels, including the check interval.
    /// </summary>
    public class OptionsForm : Form
    {
        // Controls
        private ComboBox comboDevices;
        private Label lblDevices;
        private NumericUpDown numericVolume;
        private Label lblVolume;
        private NumericUpDown numericInterval;
        private Label lblInterval;
        private Button btnAdd;
        private Button btnRemove;
        private Button btnOK;
        private Button btnCancel;
        private ListBox listMicSettings;
        private Label lblMicSettings;

        private CoreAudioController _audioController;
        private List<CoreAudioDevice> _captureDevices;

        public OptionsForm()
        {
            // Basic form setup
            this.Text = "Microphone Options";
            this.Size = new Size(500, 400);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.Font = new Font("Segoe UI", 9);

            // Initialize audio controller
            _audioController = new CoreAudioController();

            // Manually create & configure controls
            InitializeControls();

            // Load event for final data binding
            this.Load += OptionsForm_Load;
        }

        private void InitializeControls()
        {
            // Device Label
            lblDevices = new Label
            {
                Text = "Select Microphone:",
                AutoSize = true,
                Location = new Point(10, 15)
            };
            this.Controls.Add(lblDevices);

            // ComboBox for devices
            comboDevices = new ComboBox
            {
                Location = new Point(130, 10),
                Width = 300,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            this.Controls.Add(comboDevices);

            // Volume label
            lblVolume = new Label
            {
                Text = "Volume (1–100):",
                AutoSize = true,
                Location = new Point(10, 55)
            };
            this.Controls.Add(lblVolume);

            // NumericUpDown for volume
            numericVolume = new NumericUpDown
            {
                Location = new Point(130, 50),
                Width = 80,
                Minimum = 1,
                Maximum = 100,
                Value = 50
            };
            this.Controls.Add(numericVolume);

            // Add/Update button
            btnAdd = new Button
            {
                Text = "Add/Update",
                Location = new Point(220, 48),
                Width = 100
            };
            btnAdd.Click += BtnAdd_Click;
            this.Controls.Add(btnAdd);

            // Mic settings label
            lblMicSettings = new Label
            {
                Text = "Configured Microphones:",
                AutoSize = true,
                Location = new Point(10, 95)
            };
            this.Controls.Add(lblMicSettings);

            // List box for existing mic settings
            listMicSettings = new ListBox
            {
                Location = new Point(10, 115),
                Size = new Size(460, 140)
            };
            this.Controls.Add(listMicSettings);

            // Remove button
            btnRemove = new Button
            {
                Text = "Remove",
                Location = new Point(10, 265),
                Width = 80
            };
            btnRemove.Click += BtnRemove_Click;
            this.Controls.Add(btnRemove);

            // Interval label
            lblInterval = new Label
            {
                Text = "Check Interval (seconds):",
                AutoSize = true,
                Location = new Point(10, 310)
            };
            this.Controls.Add(lblInterval);

            // Numeric for interval
            numericInterval = new NumericUpDown
            {
                Location = new Point(170, 305),
                Width = 80,
                Minimum = 5,
                Maximum = 3600,
                Value = 60 // default
            };
            this.Controls.Add(numericInterval);

            // OK button
            btnOK = new Button
            {
                Text = "OK",
                Location = new Point(310, 320),
                Width = 75,
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            // Cancel button
            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(395, 320),
                Width = 75,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.Click += BtnCancel_Click;
            this.Controls.Add(btnCancel);

            // Set default accept/cancel buttons
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void OptionsForm_Load(object sender, EventArgs e)
        {
            // Load audio devices
            _captureDevices = _audioController
                .GetDevices(DeviceType.Capture, DeviceState.Active)
                .OrderBy(d => d.FullName)
                .ToList();

            comboDevices.DataSource = _captureDevices;
            comboDevices.DisplayMember = "FullName";
            comboDevices.ValueMember = "Id";

            // Load existing mic settings to list
            listMicSettings.Items.Clear();
            foreach (var ms in UserSettings.MicSettings)
            {
                listMicSettings.Items.Add(ms);
            }

            // Interval from user settings
            numericInterval.Value = UserSettings.CheckIntervalSeconds;
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (comboDevices.SelectedItem == null)
                return;

            var selectedDevice = (CoreAudioDevice)comboDevices.SelectedItem;
            int volume = (int)numericVolume.Value;

            var newMicSetting = new MicSetting
            {
                DeviceId = selectedDevice.Id.ToString(),
                DeviceName = selectedDevice.FullName,
                Volume = volume
            };

            // Check if there's an existing entry for this device
            var existing = UserSettings.MicSettings
                .FirstOrDefault(m => m.DeviceId == newMicSetting.DeviceId);

            if (existing == null)
            {
                // Add new
                UserSettings.MicSettings.Add(newMicSetting);
                listMicSettings.Items.Add(newMicSetting);
            }
            else
            {
                // Update volume
                existing.Volume = volume;
                existing.DeviceName = selectedDevice.FullName;

                // Refresh display in list
                int index = listMicSettings.Items.IndexOf(existing);
                if (index >= 0)
                {
                    listMicSettings.Items.RemoveAt(index);
                    listMicSettings.Items.Insert(index, existing);
                }
            }
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            var selected = listMicSettings.SelectedItem as MicSetting;
            if (selected == null)
                return;

            UserSettings.MicSettings.Remove(selected);
            listMicSettings.Items.Remove(selected);
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            // Save interval
            UserSettings.CheckIntervalSeconds = (int)numericInterval.Value;
            // Persist to registry
            UserSettings.SaveToRegistry();

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }

    /// <summary>
    /// Static class holding user settings and providing load/save methods to the user registry.
    /// </summary>
    public static class UserSettings
    {
        private const string RegistryKeyPath = @"Software\SACK Corporation\MicLevelTrayApp";

        public static List<MicSetting> MicSettings { get; private set; } = new List<MicSetting>();

        public static int CheckIntervalSeconds { get; set; } = 60; // default 60

        public static void LoadFromRegistry()
        {
            MicSettings.Clear();

            using (var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, false))
            {
                if (key == null)
                    return;

                // Load check interval
                var intervalStr = key.GetValue("CheckIntervalSeconds") as string;
                int parsedInterval;
                if (int.TryParse(intervalStr, out parsedInterval))
                {
                    CheckIntervalSeconds = parsedInterval;
                }
                else
                {
                    CheckIntervalSeconds = 60; // fallback
                }

                // Load mic subkeys: Mic0, Mic1, ...
                int i = 0;
                while (true)
                {
                    string subKeyName = "Mic" + i;
                    using (var micKey = key.OpenSubKey(subKeyName, false))
                    {
                        if (micKey == null)
                            break;

                        var deviceId = micKey.GetValue("DeviceId") as string;
                        var deviceName = micKey.GetValue("DeviceName") as string;
                        var volumeStr = micKey.GetValue("Volume") as string;
                        if (!string.IsNullOrEmpty(deviceId) && !string.IsNullOrEmpty(volumeStr))
                        {
                            int vol;
                            if (int.TryParse(volumeStr, out vol))
                            {
                                MicSettings.Add(new MicSetting { DeviceId = deviceId, DeviceName = deviceName, Volume = vol });
                            }
                        }
                    }
                    i++;
                }
            }
        }

        public static void SaveToRegistry()
        {
            using (var key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
            {
                // Save check interval
                key.SetValue("CheckIntervalSeconds", CheckIntervalSeconds.ToString());

                // Clean existing subkeys first
                int i = 0;
                while (true)
                {
                    string subKeyName = "Mic" + i;
                    using (var existing = key.OpenSubKey(subKeyName))
                    {
                        if (existing == null)
                            break;
                        key.DeleteSubKey(subKeyName);
                    }
                    i++;
                }

                // Write new subkeys
                for (int j = 0; j < MicSettings.Count; j++)
                {
                    string subKeyName = "Mic" + j;
                    using (var micKey = key.CreateSubKey(subKeyName))
                    {
                        micKey.SetValue("DeviceId", MicSettings[j].DeviceId);
                        micKey.SetValue("DeviceName", MicSettings[j].DeviceName);
                        micKey.SetValue("Volume", MicSettings[j].Volume.ToString());
                    }
                }
            }
        }
    }

    /// <summary>
    /// Simple data class to hold a microphone's device ID and target volume.
    /// </summary>
    public class MicSetting
    {
        public string DeviceId { get; set; }
        public string DeviceName { get; set; }
        public int Volume { get; set; }

        public override string ToString()
        {
            var nameToShow = string.IsNullOrWhiteSpace(DeviceName) ? DeviceId : DeviceName;
            return string.Format("DeviceName: {0}, Volume: {1}%", nameToShow, Volume);
        }
    }

    /// <summary>
    /// A helper class to re-apply volumes periodically.
    /// </summary>
    public static class VolumeEnforcer
    {
        private static CoreAudioController _audioController = new CoreAudioController();

        public static void EnforceConfiguredVolumes()
        {
            // For each configured microphone device, set the volume if it differs
            foreach (var ms in UserSettings.MicSettings)
            {
                int desiredVolume = ms.Volume;
                if (desiredVolume < 0) desiredVolume = 0;
                if (desiredVolume > 100) desiredVolume = 100;

                Guid deviceGuid;
                if (!Guid.TryParse(ms.DeviceId, out deviceGuid))
                    continue;

                // Attempt to get the device
                var device = _audioController.GetDevice(deviceGuid, DeviceState.Active);
                if (device == null || device.DeviceType != DeviceType.Capture)
                    continue; // device not found or inactive

                var currentVolume = device.Volume;
                if (Math.Abs(currentVolume - desiredVolume) > 0.5f)
                {
                    // Set volume
                    device.Volume = desiredVolume;
                }
            }
        }
    }
}
