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
        var view = new ConsoleView();
        var controller = new AudioDeviceController(deviceService, view);

        // アプリケーション実行
        controller.Run();
    }
}