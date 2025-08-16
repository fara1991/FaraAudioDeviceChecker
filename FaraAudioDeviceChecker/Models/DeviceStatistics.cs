namespace FaraAudioDeviceChecker.Models;

public class DeviceStatistics
{
    public Dictionary<string, int> ClassCount { get; set; } = new();
    public Dictionary<string, int> StatusCount { get; set; } = new();
    public Dictionary<string, int> ManufacturerCount { get; set; } = new();
}