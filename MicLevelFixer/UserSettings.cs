using System.Collections.Generic;
using Microsoft.Win32;

namespace MicLevelFixer
{
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

            using (var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath, false)) {
                if (key == null)
                    return;

                // Load check interval
                var intervalStr = key.GetValue("CheckIntervalSeconds") as string;
                if (int.TryParse(intervalStr, out var parsedInterval)) {
                    CheckIntervalSeconds = parsedInterval;
                } else {
                    CheckIntervalSeconds = 60; // fallback
                }

                // Load mic subkeys: Mic0, Mic1, ...
                int i = 0;
                while (true) {
                    string subKeyName = "Mic" + i;
                    using (var micKey = key.OpenSubKey(subKeyName, false)) {
                        if (micKey == null)
                            break;

                        var deviceId = micKey.GetValue("DeviceId") as string;
                        var deviceName = micKey.GetValue("DeviceName") as string;
                        var volumeStr = micKey.GetValue("Volume") as string;
                        if (!string.IsNullOrEmpty(deviceId) && !string.IsNullOrEmpty(volumeStr)) {
                            if (int.TryParse(volumeStr, out var vol)) {
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
            using (var key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath)) {
                // Save check interval
                key.SetValue("CheckIntervalSeconds", CheckIntervalSeconds.ToString());

                // Clean existing subkeys first
                int i = 0;
                while (true) {
                    string subKeyName = "Mic" + i;
                    using (var existing = key.OpenSubKey(subKeyName)) {
                        if (existing == null)
                            break;
                        key.DeleteSubKey(subKeyName);
                    }
                    i++;
                }

                // Write new subkeys
                for (int j = 0; j < MicSettings.Count; j++) {
                    string subKeyName = "Mic" + j;
                    using (var micKey = key.CreateSubKey(subKeyName)) {
                        micKey.SetValue("DeviceId", MicSettings[j].DeviceId);
                        micKey.SetValue("DeviceName", MicSettings[j].DeviceName);
                        micKey.SetValue("Volume", MicSettings[j].Volume.ToString());
                    }
                }
            }
        }
    }
}
