using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AudioSwitcher.AudioApi;
using AudioSwitcher.AudioApi.CoreAudio;

namespace MicLevelFixer
{
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
        private ToolTip toolTip1 = new ToolTip();
        private int _lastHoveredIndex = -1;

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
            lblDevices = new Label {
                Text = "Select Microphone:",
                AutoSize = true,
                Location = new Point(10, 15)
            };
            this.Controls.Add(lblDevices);

            // ComboBox for devices
            comboDevices = new ComboBox {
                Location = new Point(130, 10),
                Width = 300,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            this.Controls.Add(comboDevices);

            // Volume label
            lblVolume = new Label {
                Text = "Volume (1–100):",
                AutoSize = true,
                Location = new Point(10, 55)
            };
            this.Controls.Add(lblVolume);

            // NumericUpDown for volume
            numericVolume = new NumericUpDown {
                Location = new Point(130, 50),
                Width = 80,
                Minimum = 1,
                Maximum = 100,
                Value = 50
            };
            this.Controls.Add(numericVolume);

            // Add/Update button
            btnAdd = new Button {
                Text = "Add/Update",
                Location = new Point(220, 48),
                Width = 100
            };
            btnAdd.Click += BtnAdd_Click;
            this.Controls.Add(btnAdd);

            // Mic settings label
            lblMicSettings = new Label {
                Text = "Configured Microphones:",
                AutoSize = true,
                Location = new Point(10, 95)
            };
            this.Controls.Add(lblMicSettings);

            // List box for existing mic settings
            listMicSettings = new ListBox {
                Location = new Point(10, 115),
                Size = new Size(460, 140)
            };
            this.Controls.Add(listMicSettings);

            // Remove button
            btnRemove = new Button {
                Text = "Remove",
                Location = new Point(10, 265),
                Width = 80
            };
            btnRemove.Click += BtnRemove_Click;
            this.Controls.Add(btnRemove);

            // Interval label
            lblInterval = new Label {
                Text = "Check Interval (seconds):",
                AutoSize = true,
                Location = new Point(10, 310)
            };
            this.Controls.Add(lblInterval);

            // Numeric for interval
            numericInterval = new NumericUpDown {
                Location = new Point(170, 305),
                Width = 80,
                Minimum = 5,
                Maximum = 3600,
                Value = 60 // default
            };
            this.Controls.Add(numericInterval);

            // OK button
            btnOK = new Button {
                Text = "OK",
                Location = new Point(310, 320),
                Width = 75,
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            // Cancel button
            btnCancel = new Button {
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
            foreach (var ms in UserSettings.MicSettings) {
                listMicSettings.Items.Add(ms);
            }

            // Interval from user settings
            numericInterval.Value = UserSettings.CheckIntervalSeconds;
            listMicSettings.MouseMove += listMicSettings_MouseMove;
        }

        private void listMicSettings_MouseMove(object sender, MouseEventArgs e)
        {
            // Figure out which item is under the mouse
            int index = listMicSettings.IndexFromPoint(e.Location);

            // Only update the tooltip if we moved to a different item
            if (index != _lastHoveredIndex) {
                _lastHoveredIndex = index;

                if (index >= 0) {
                    // We have a valid item

                    // Show the *full* device name in the tooltip
                    if (listMicSettings.Items[index] is MicSetting item && !string.IsNullOrEmpty(item.DeviceName)) {
                        toolTip1.SetToolTip(listMicSettings, item.DeviceName);
                    } else {
                        toolTip1.SetToolTip(listMicSettings, string.Empty);
                    }
                } else {
                    // Mouse is not over a valid list item
                    toolTip1.SetToolTip(listMicSettings, string.Empty);
                }
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            if (comboDevices.SelectedItem == null)
                return;

            var selectedDevice = (CoreAudioDevice)comboDevices.SelectedItem;
            int volume = (int)numericVolume.Value;

            var newMicSetting = new MicSetting {
                DeviceId = selectedDevice.Id.ToString(),
                DeviceName = selectedDevice.FullName,
                Volume = volume
            };

            // Check if there's an existing entry for this device
            var existing = UserSettings.MicSettings
                .FirstOrDefault(m => m.DeviceId == newMicSetting.DeviceId);

            if (existing == null) {
                // Add new
                UserSettings.MicSettings.Add(newMicSetting);
                listMicSettings.Items.Add(newMicSetting);
            } else {
                // Update volume
                existing.Volume = volume;
                existing.DeviceName = selectedDevice.FullName;

                // Refresh display in list
                int index = listMicSettings.Items.IndexOf(existing);
                if (index >= 0) {
                    listMicSettings.Items.RemoveAt(index);
                    listMicSettings.Items.Insert(index, existing);
                }
            }
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            if (!(listMicSettings.SelectedItem is MicSetting selected))
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
}
