namespace FaraAudioDeviceChecker;

using Controllers;
using Services;
using Views;

internal static class Program
{
    private static void Main(string[] args)
    {
        // 依存性注入のセットアップ
        IDeviceService deviceService = new DeviceService();
        var controller = new AudioDeviceController(deviceService);

        // アプリケーション実行
        controller.Run();
    }
}