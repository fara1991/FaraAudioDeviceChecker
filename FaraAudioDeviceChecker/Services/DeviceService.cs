namespace FaraAudioDeviceChecker.Services;

using System.Management;
using Models;
using Utilities;

public class DeviceService : IDeviceService
{
    private readonly string[] _audioClasses = ["Media", "AudioEndpoint", "SoftwareDevice"];

    public List<AudioDeviceInfo> GetAudioDevices()
    {
        var devices = new List<AudioDeviceInfo>();

        foreach (var audioClass in _audioClasses)
        {
            var query =
                $"SELECT * FROM Win32_PnPEntity WHERE PNPClass = '{audioClass}' OR Name LIKE '%audio%' OR Name LIKE '%sound%'";

            using var searcher = new ManagementObjectSearcher(query);
            var collection = searcher.Get();

            foreach (var o in collection)
            {
                var device = (ManagementObject) o;
                var deviceInfo = new AudioDeviceInfo
                {
                    Name = WmiHelper.GetPropertyValue(device, "Name"),
                    DeviceId = WmiHelper.GetPropertyValue(device, "DeviceID"),
                    Description = WmiHelper.GetPropertyValue(device, "Description"),
                    Manufacturer = WmiHelper.GetPropertyValue(device, "Manufacturer"),
                    Status = WmiHelper.GetPropertyValue(device, "Status"),
                    Class = WmiHelper.GetPropertyValue(device, "PNPClass"),
                    Service = WmiHelper.GetPropertyValue(device, "Service"),
                    HardwareId = WmiHelper.GetPropertyValue(device, "HardwareID")
                };

                SetProblemCode(device, deviceInfo);
                GetDriverInfo(deviceInfo);
                devices.Add(deviceInfo);
            }
        }

        return RemoveDuplicates(devices);
    }

    public DeviceStatistics GetDeviceStatistics(List<AudioDeviceInfo> devices)
    {
        var statistics = new DeviceStatistics();

        foreach (var device in devices)
        {
            var className = string.IsNullOrEmpty(device.Class) ? "不明" : device.Class;
            if (!statistics.ClassCount.TryAdd(className, 1))
                statistics.ClassCount[className]++;

            var status = string.IsNullOrEmpty(device.Status) ? "不明" : device.Status;
            if (!statistics.StatusCount.TryAdd(status, 1))
                statistics.StatusCount[status]++;

            var manufacturer = string.IsNullOrEmpty(device.Manufacturer) ? "不明" : device.Manufacturer;
            if (!statistics.ManufacturerCount.TryAdd(manufacturer, 1))
                statistics.ManufacturerCount[manufacturer]++;
        }

        return statistics;
    }

    public List<AudioDeviceInfo> GetProblemDevices(List<AudioDeviceInfo> devices)
    {
        return devices.FindAll(d => d.HasProblem || d.Status != "OK" || d.DriverVersion.StartsWith("取得エラー:"));
    }

    public List<AudioDeviceInfo> GetOldDriverDevices(List<AudioDeviceInfo> devices)
    {
        return devices.FindAll(IsDriverOld);
    }

    private static void SetProblemCode(ManagementObject device, AudioDeviceInfo deviceInfo)
    {
        if (device["ConfigManagerErrorCode"] == null) return;
        
        var problemCode = (uint) device["ConfigManagerErrorCode"];
        deviceInfo.HasProblem = problemCode != 0;
        deviceInfo.ProblemCode = ErrorCodeHelper.GetProblemDescription(problemCode);
    }

    private static void GetDriverInfo(AudioDeviceInfo device)
    {
        try
        {
            var escapedDeviceId = WmiHelper.EscapeWqlString(device.DeviceId);

            var query = $"SELECT * FROM Win32_PnPSignedDriver WHERE DeviceID = '{escapedDeviceId}'";
            using var searcher = new ManagementObjectSearcher(query);
            var collection = searcher.Get();

            foreach (var o in collection)
            {
                var driver = (ManagementObject) o;
                device.DriverVersion = driver["DriverVersion"]?.ToString() ?? "不明";
                device.DriverDate = driver["DriverDate"]?.ToString() ?? "不明";
                break;
            }

            if (device.DriverVersion == "不明" && !string.IsNullOrEmpty(device.Service))
            {
                GetSystemDriverInfo(device);
            }
        }
        catch (Exception ex)
        {
            device.DriverVersion = $"取得エラー: {ex.Message}";
            device.HasProblem = true;
            device.ProblemCode = "ドライバー情報取得エラー";
        }
    }

    private static void GetSystemDriverInfo(AudioDeviceInfo device)
    {
        var escapedService = WmiHelper.EscapeWqlString(device.Service);
        var query = $"SELECT * FROM Win32_SystemDriver WHERE Name = '{escapedService}'";
        using var driverSearcher = new ManagementObjectSearcher(query);
        var driverCollection = driverSearcher.Get();

        foreach (var o in driverCollection)
        {
            var driver = (ManagementObject) o;
            device.DriverVersion = driver["Version"]?.ToString() ?? "不明";

            if (driver["InstallDate"] != null)
            {
                var installDate = ManagementDateTimeConverter.ToDateTime(driver["InstallDate"].ToString());
                device.DriverDate = installDate.ToString("yyyy/MM/dd");
            }

            break;
        }
    }

    private static bool IsDriverOld(AudioDeviceInfo device)
    {
        if (string.IsNullOrEmpty(device.DriverDate) || device.DriverDate == "不明")
            return false;

        if (!DateTime.TryParse(device.DriverDate, out var driverDate))
            return false;

        var age = DateTime.Now - driverDate;
        return age.TotalDays > 365;
    }

    private static List<AudioDeviceInfo> RemoveDuplicates(List<AudioDeviceInfo> devices)
    {
        var seenDeviceIds = new HashSet<string>();
        return devices.Where(device => seenDeviceIds.Add(device.DeviceId)).ToList();
    }
}