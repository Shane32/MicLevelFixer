using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AudioSwitcher.AudioApi;
using AudioSwitcher.AudioApi.CoreAudio;

namespace MicLevelFixer;

/// <summary>
/// A form to configure the microphone levels, including the check interval.
/// </summary>
public class OptionsForm : Form
{
    // Controls
    private readonly ComboBox _comboDevices;
    private readonly Label _lblDevices;
    private readonly NumericUpDown _numericVolume;
    private readonly Label _lblVolume;
    private readonly NumericUpDown _numericInterval;
    private readonly Label _lblInterval;
    private readonly Button _btnAdd;
    private readonly Button _btnRemove;
    private readonly Button _btnOK;
    private readonly Button _btnCancel;
    private readonly ListBox _listMicSettings;
    private readonly Label _lblMicSettings;

    private readonly CoreAudioController _audioController;
    private readonly ToolTip _toolTip1 = new ToolTip();
    private int _lastHoveredIndex = -1;

    public OptionsForm()
    {
        // Basic form setup
        Text = "Microphone Options";
        Size = new Size(500, 400);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Font = new Font("Segoe UI", 9);

        // Initialize audio controller
        _audioController = new CoreAudioController();

        // Device Label
        _lblDevices = new Label {
            Text = "Select Microphone:",
            AutoSize = true,
            Location = new Point(10, 15)
        };
        Controls.Add(_lblDevices);

        // ComboBox for devices
        _comboDevices = new ComboBox {
            Location = new Point(130, 10),
            Width = 300,
            DropDownStyle = ComboBoxStyle.DropDownList
        };
        Controls.Add(_comboDevices);

        // Volume label
        _lblVolume = new Label {
            Text = "Volume (1–100):",
            AutoSize = true,
            Location = new Point(10, 55)
        };
        Controls.Add(_lblVolume);

        // NumericUpDown for volume
        _numericVolume = new NumericUpDown {
            Location = new Point(130, 50),
            Width = 80,
            Minimum = 1,
            Maximum = 100,
            Value = 50
        };
        Controls.Add(_numericVolume);

        // Add/Update button
        _btnAdd = new Button {
            Text = "Add/Update",
            Location = new Point(220, 48),
            Width = 100
        };
        _btnAdd.Click += BtnAdd_Click;
        Controls.Add(_btnAdd);

        // Mic settings label
        _lblMicSettings = new Label {
            Text = "Configured Microphones:",
            AutoSize = true,
            Location = new Point(10, 95)
        };
        Controls.Add(_lblMicSettings);

        // List box for existing mic settings
        _listMicSettings = new ListBox {
            Location = new Point(10, 115),
            Size = new Size(460, 140)
        };
        Controls.Add(_listMicSettings);

        // Remove button
        _btnRemove = new Button {
            Text = "Remove",
            Location = new Point(10, 265),
            Width = 80
        };
        _btnRemove.Click += BtnRemove_Click;
        Controls.Add(_btnRemove);

        // Interval label
        _lblInterval = new Label {
            Text = "Check Interval (seconds):",
            AutoSize = true,
            Location = new Point(10, 310)
        };
        Controls.Add(_lblInterval);

        // Numeric for interval
        _numericInterval = new NumericUpDown {
            Location = new Point(170, 305),
            Width = 80,
            Minimum = 5,
            Maximum = 3600,
            Value = 60 // default
        };
        Controls.Add(_numericInterval);

        // OK button
        _btnOK = new Button {
            Text = "OK",
            Location = new Point(310, 320),
            Width = 75,
            DialogResult = DialogResult.OK
        };
        _btnOK.Click += BtnOK_Click;
        Controls.Add(_btnOK);

        // Cancel button
        _btnCancel = new Button {
            Text = "Cancel",
            Location = new Point(395, 320),
            Width = 75,
            DialogResult = DialogResult.Cancel
        };
        _btnCancel.Click += BtnCancel_Click;
        Controls.Add(_btnCancel);

        // Set default accept/cancel buttons
        AcceptButton = _btnOK;
        CancelButton = _btnCancel;

        // Load event for final data binding
        Load += OptionsForm_Load;
    }

    private void OptionsForm_Load(object sender, EventArgs e)
    {
        // Load audio devices
        var captureDevices = _audioController
                .GetDevices(DeviceType.Capture, DeviceState.Active)
                .OrderBy(d => d.FullName)
                .ToList();

        _comboDevices.DataSource = captureDevices;
        _comboDevices.DisplayMember = "FullName";
        _comboDevices.ValueMember = "Id";

        // Load existing mic settings to list
        _listMicSettings.Items.Clear();
        foreach (var ms in UserSettings.MicSettings) {
            _listMicSettings.Items.Add(ms);
        }

        // Interval from user settings
        _numericInterval.Value = UserSettings.CheckIntervalSeconds;
        _listMicSettings.MouseMove += listMicSettings_MouseMove;
    }

    private void listMicSettings_MouseMove(object sender, MouseEventArgs e)
    {
        // Figure out which item is under the mouse
        int index = _listMicSettings.IndexFromPoint(e.Location);

        // Only update the tooltip if we moved to a different item
        if (index != _lastHoveredIndex) {
            _lastHoveredIndex = index;

            if (index >= 0) {
                // We have a valid item

                // Show the *full* device name in the tooltip
                if (_listMicSettings.Items[index] is MicSetting item && !string.IsNullOrEmpty(item.DeviceName)) {
                    _toolTip1.SetToolTip(_listMicSettings, item.DeviceName);
                } else {
                    _toolTip1.SetToolTip(_listMicSettings, string.Empty);
                }
            } else {
                // Mouse is not over a valid list item
                _toolTip1.SetToolTip(_listMicSettings, string.Empty);
            }
        }
    }

    private void BtnAdd_Click(object sender, EventArgs e)
    {
        if (_comboDevices.SelectedItem == null)
            return;

        var selectedDevice = (CoreAudioDevice)_comboDevices.SelectedItem;
        int volume = (int)_numericVolume.Value;

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
            _listMicSettings.Items.Add(newMicSetting);
        } else {
            // Update volume
            existing.Volume = volume;
            existing.DeviceName = selectedDevice.FullName;

            // Refresh display in list
            int index = _listMicSettings.Items.IndexOf(existing);
            if (index >= 0) {
                _listMicSettings.Items.RemoveAt(index);
                _listMicSettings.Items.Insert(index, existing);
            }
        }
    }

    private void BtnRemove_Click(object sender, EventArgs e)
    {
        if (!(_listMicSettings.SelectedItem is MicSetting selected))
            return;

        UserSettings.MicSettings.Remove(selected);
        _listMicSettings.Items.Remove(selected);
    }

    private void BtnOK_Click(object sender, EventArgs e)
    {
        // Save interval
        UserSettings.CheckIntervalSeconds = (int)_numericInterval.Value;
        // Persist to registry
        UserSettings.SaveToRegistry();

        DialogResult = DialogResult.OK;
        Close();
    }

    private void BtnCancel_Click(object sender, EventArgs e)
    {
        DialogResult = DialogResult.Cancel;
        Close();
    }
}
