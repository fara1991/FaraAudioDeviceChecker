namespace FaraAudioDeviceChecker.Services;

using Models;

public interface IDeviceService
{
    List<AudioDeviceInfo> GetAudioDevices();
    DeviceStatistics GetDeviceStatistics(List<AudioDeviceInfo> devices);
    List<AudioDeviceInfo> GetProblemDevices(List<AudioDeviceInfo> devices);
    List<AudioDeviceInfo> GetOldDriverDevices(List<AudioDeviceInfo> devices);
}