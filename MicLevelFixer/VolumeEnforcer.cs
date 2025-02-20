using System;
using AudioSwitcher.AudioApi;
using AudioSwitcher.AudioApi.CoreAudio;

namespace MicLevelFixer;

/// <summary>
/// A helper class to re-apply volumes periodically.
/// </summary>
public static class VolumeEnforcer
{
    private static readonly CoreAudioController _audioController = new CoreAudioController();

    public static void EnforceConfiguredVolumes()
    {
        // For each configured microphone device, set the volume if it differs
        foreach (var ms in UserSettings.MicSettings) {
            int desiredVolume = ms.Volume;
            if (desiredVolume < 0)
                desiredVolume = 0;
            if (desiredVolume > 100)
                desiredVolume = 100;

            if (!Guid.TryParse(ms.DeviceId, out var deviceGuid))
                continue;

            // Attempt to get the device
            var device = _audioController.GetDevice(deviceGuid, DeviceState.Active);
            if (device == null || device.DeviceType != DeviceType.Capture)
                continue; // device not found or inactive

            var currentVolume = device.Volume;
            if (Math.Abs(currentVolume - desiredVolume) > 0.5f) {
                // Set volume
                device.Volume = desiredVolume;
            }
        }
    }
}
