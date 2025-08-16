namespace FaraAudioDeviceChecker.Models;

public class AudioDeviceInfo
{
    public required string Name { get; set; }
    public required string DeviceId { get; set; }
    public required string Description { get; set; }
    public required string Manufacturer { get; set; }
    public required string Status { get; set; }
    public required string Class { get; set; }
    public required string Service { get; set; }
    public string DriverVersion { get; set; } = "不明";
    public string DriverDate { get; set; } = "不明";
    public bool HasProblem { get; set; }
    public string ProblemCode { get; set; } = "正常";
    public string HardwareId { get; set; } = "不明";
}