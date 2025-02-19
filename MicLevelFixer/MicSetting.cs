namespace MicLevelFixer
{
    /// <summary>
    /// Simple data class to hold a microphone's device ID and target volume.
    /// </summary>
    public class MicSetting
    {
        public string DeviceId { get; set; }
        public string DeviceName { get; set; }
        public int Volume { get; set; }
        private string TruncatedDeviceName
        {
            get {
                if (string.IsNullOrWhiteSpace(DeviceName))
                    return DeviceId; // fallback if there's no friendly name

                if (DeviceName.Length <= 20)
                    return DeviceName;
                else
                    return DeviceName.Substring(0, 20 - 3) + "...";
            }
        }

        public override string ToString()
        {
            return string.Format("DeviceName: {0}, Volume: {1}%", TruncatedDeviceName, Volume);
        }
    }
}
