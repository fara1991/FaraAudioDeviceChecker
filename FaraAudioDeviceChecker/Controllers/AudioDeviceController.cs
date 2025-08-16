namespace FaraAudioDeviceChecker.Controllers;

using Services;
using Views;

public class AudioDeviceController(IDeviceService deviceService, ConsoleView view)
{
    public void Run()
    {
        try
        {
            ConsoleView.ShowHeader();

            var audioDevices = deviceService.GetAudioDevices();

            if (audioDevices.Count == 0)
            {
                ConsoleView.ShowNoDevicesFound();
                return;
            }

            ConsoleView.ShowDeviceCount(audioDevices.Count);

            // 各デバイスの分析を表示
            foreach (var device in audioDevices)
            {
                ConsoleView.ShowDeviceAnalysis(device);
            }

            // 問題のあるデバイスの要約
            var problemDevices = deviceService.GetProblemDevices(audioDevices);
            ConsoleView.ShowProblemSummary(problemDevices);

            // 統計情報の表示
            var statistics = deviceService.GetDeviceStatistics(audioDevices);
            ConsoleView.ShowDeviceStatistics(statistics);

            // 推奨事項の表示
            var oldDriverDevices = deviceService.GetOldDriverDevices(audioDevices);
            ConsoleView.ShowRecommendations(problemDevices, oldDriverDevices);
        }
        catch (Exception ex)
        {
            ConsoleView.ShowError(ex.Message);
        }
        finally
        {
            ConsoleView.ShowExitPrompt();
        }
    }
}