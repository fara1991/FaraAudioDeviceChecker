using System.Management;

namespace FaraDeviceChecker;

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

internal static class Program
{
    private static void Main(string[] args)
    {
        Console.WriteLine("=== Audioドライバー状態チェッカー ===\n");

        try
        {
            var audioDevices = GetAudioDevices();
            if (audioDevices.Count == 0)
            {
                Console.WriteLine("Audioデバイスが見つかりませんでした。");
                return;
            }

            Console.WriteLine($"検出されたAudioデバイス数: {audioDevices.Count}\n");

            foreach (var device in audioDevices)
            {
                AnalyzeAudioDevice(device);
                Console.WriteLine(new string('=', 80));
            }

            DisplayProblemSummary(audioDevices);
            DisplayDeviceStatistics(audioDevices);
            DisplayRecommendations(audioDevices);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"エラーが発生しました: {ex.Message}");
        }

        Console.WriteLine("\nEnterキーを押して終了してください...");
        Console.ReadLine();
    }

    private static List<AudioDeviceInfo> GetAudioDevices()
    {
        var devices = new List<AudioDeviceInfo>();
        string[] audioClasses = {"Media", "AudioEndpoint", "SoftwareDevice"};

        foreach (var audioClass in audioClasses)
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
                    Name = GetPropertyValue(device, "Name"),
                    DeviceId = GetPropertyValue(device, "DeviceID"),
                    Description = GetPropertyValue(device, "Description"),
                    Manufacturer = GetPropertyValue(device, "Manufacturer"),
                    Status = GetPropertyValue(device, "Status"),
                    Class = GetPropertyValue(device, "PNPClass"),
                    Service = GetPropertyValue(device, "Service"),
                    HardwareId = GetPropertyValue(device, "HardwareID")
                };

                if (device["ConfigManagerErrorCode"] != null)
                {
                    var problemCode = (uint) device["ConfigManagerErrorCode"];
                    deviceInfo.HasProblem = problemCode != 0;
                    deviceInfo.ProblemCode = GetProblemDescription(problemCode);
                }

                GetDriverInfo(deviceInfo);
                devices.Add(deviceInfo);
            }
        }

        return RemoveDuplicates(devices);
    }

    private static void GetDriverInfo(AudioDeviceInfo device)
    {
        try
        {
            var escapedDeviceId = EscapeWqlString(device.DeviceId);

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
                var escapedService = EscapeWqlString(device.Service);
                query = $"SELECT * FROM Win32_SystemDriver WHERE Name = '{escapedService}'";
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
        }
        catch (Exception ex)
        {
            device.DriverVersion = $"取得エラー: {ex.Message}";
            device.HasProblem = true;
            device.ProblemCode = "ドライバー情報取得エラー";
        }
    }

    private static string EscapeWqlString(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        return input
            .Replace("\\", "\\\\") // バックスラッシュ
            .Replace("'", "''") // シングルクォート
            .Replace("\"", "\\\"") // ダブルクォート
            .Replace("%", "[%]") // パーセント
            .Replace("_", "[_]") // アンダースコア
            .Replace("&", "^&"); // アンパサンド
    }

    private static void AnalyzeAudioDevice(AudioDeviceInfo device)
    {
        Console.WriteLine($"デバイス名: {device.Name}");
        Console.WriteLine($"製造元: {device.Manufacturer}");
        Console.WriteLine($"デバイスID: {device.DeviceId}");
        Console.WriteLine($"ドライバーバージョン: {device.DriverVersion}");
        Console.WriteLine($"ドライバー日付: {device.DriverDate}");
        Console.WriteLine($"ステータス: {device.Status}");
        Console.WriteLine($"クラス: {device.Class}");
        Console.WriteLine($"サービス: {device.Service}");

        var needsAttention = false;
        var issues = new List<string>();

        if (device.HasProblem)
        {
            issues.Add($"問題コード: {device.ProblemCode}");
            needsAttention = true;
        }

        if (device.Status != "OK")
        {
            issues.Add($"ステータスが正常ではありません: {device.Status}");
            needsAttention = true;
        }

        if (!string.IsNullOrEmpty(device.DriverDate) && device.DriverDate != "不明")
        {
            if (DateTime.TryParse(device.DriverDate, out var driverDate))
            {
                var age = DateTime.Now - driverDate;
                if (age.TotalDays > 365)
                {
                    issues.Add($"ドライバーが古い可能性があります（{age.TotalDays:F0}日前）");
                    needsAttention = true;
                }
            }
        }

        if (device.DriverVersion.StartsWith("取得エラー:"))
        {
            issues.Add("ドライバー情報が取得できませんでした");
            needsAttention = true;
        }

        if (needsAttention)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("\n⚠️ 注意が必要:");
            foreach (var issue in issues)
            {
                Console.WriteLine($"  - {issue}");
            }
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("\n✅ 正常に動作しています");
        }

        Console.ResetColor();
        Console.WriteLine();
    }

    private static void DisplayProblemSummary(List<AudioDeviceInfo> devices)
    {
        Console.WriteLine("\n=== 問題のあるデバイス要約 ===");

        var problemDevices =
            devices.FindAll(d => d.HasProblem || d.Status != "OK" || d.DriverVersion.StartsWith("取得エラー:"));

        if (problemDevices.Count == 0)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("すべてのAudioデバイスが正常に動作しています。");
            Console.ResetColor();
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"{problemDevices.Count}個のデバイスに問題があります:");
            Console.ResetColor();

            foreach (var device in problemDevices)
            {
                var reason = device.DriverVersion.StartsWith("取得エラー:") ? "ドライバー情報取得エラー" : device.ProblemCode;
                Console.WriteLine($"  - {device.Name}: {reason}");
            }
        }
    }

    private static void DisplayDeviceStatistics(List<AudioDeviceInfo> devices)
    {
        Console.WriteLine("\n=== デバイス統計 ===");

        var classCount = new Dictionary<string, int>();
        var statusCount = new Dictionary<string, int>();
        var manufacturerCount = new Dictionary<string, int>();

        foreach (var device in devices)
        {
            var className = string.IsNullOrEmpty(device.Class) ? "不明" : device.Class;
            if (!classCount.TryAdd(className, 1))
                classCount[className]++;

            var status = string.IsNullOrEmpty(device.Status) ? "不明" : device.Status;
            if (!statusCount.TryAdd(status, 1))
                statusCount[status]++;

            var manufacturer = string.IsNullOrEmpty(device.Manufacturer) ? "不明" : device.Manufacturer;
            if (!manufacturerCount.TryAdd(manufacturer, 1))
                manufacturerCount[manufacturer]++;
        }

        Console.WriteLine("\nデバイスクラス別統計:");
        foreach (var kvp in classCount)
        {
            Console.WriteLine($"  {kvp.Key}: {kvp.Value}個");
        }

        Console.WriteLine("\nステータス別統計:");
        foreach (var kvp in statusCount)
        {
            Console.WriteLine($"  {kvp.Key}: {kvp.Value}個");
        }

        Console.WriteLine("\n製造元別統計:");
        foreach (var kvp in manufacturerCount)
        {
            Console.WriteLine($"  {kvp.Key}: {kvp.Value}個");
        }
    }

    private static void DisplayRecommendations(List<AudioDeviceInfo> devices)
    {
        Console.WriteLine("\n=== 推奨事項 ===");

        var problemDevices =
            devices.FindAll(d => d.HasProblem || d.Status != "OK" || d.DriverVersion.StartsWith("取得エラー:"));
        var oldDriverDevices = devices.FindAll(d => IsDriverOld(d));

        if (problemDevices.Count == 0 && oldDriverDevices.Count == 0)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("現在、特に対応が必要な問題は見つかりませんでした。");
            Console.ResetColor();
            return;
        }

        if (problemDevices.Count > 0)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"🔴 緊急対応が必要: {problemDevices.Count}個のデバイス");
            Console.ResetColor();
            Console.WriteLine("以下の方法で対応してください:");
            Console.WriteLine("  1. デバイスマネージャーを開く");
            Console.WriteLine("  2. 問題のあるデバイスを右クリック");
            Console.WriteLine("  3. 「ドライバーの更新」を選択");
            Console.WriteLine("  4. 「コンピューターを参照してドライバーを検索」");
            Console.WriteLine("  5. 「コンピューター上の利用可能なドライバーの一覧から選択」");
            Console.WriteLine();
        }

        if (oldDriverDevices.Count > 0)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"🟡 更新推奨: {oldDriverDevices.Count}個のデバイスのドライバーが古い");
            Console.ResetColor();
            Console.WriteLine("以下の方法で更新してください:");
            Console.WriteLine("  1. Windows Update を実行");
            Console.WriteLine("  2. 製造元の公式サイトから最新ドライバーをダウンロード");
            Console.WriteLine("  3. デバイスマネージャーからドライバーを更新");
            Console.WriteLine();
        }

        Console.WriteLine("📋 一般的な対応手順:");
        Console.WriteLine("  • デバイスマネージャー: Windowsキー + X → デバイスマネージャー");
        Console.WriteLine("  • Windows Update: 設定 → Windows Update → 更新プログラムのチェック");
        Console.WriteLine("  • 製造元サイト: 各デバイスの製造元の公式サポートページ");
    }

    private static bool IsDriverOld(AudioDeviceInfo device)
    {
        if (string.IsNullOrEmpty(device.DriverDate) || device.DriverDate == "不明")
            return false;

        if (!DateTime.TryParse(device.DriverDate, out var driverDate)) return false;

        var age = DateTime.Now - driverDate;
        return age.TotalDays > 365;
    }

    private static string GetPropertyValue(ManagementObject obj, string propertyName)
    {
        try
        {
            return obj[propertyName]?.ToString() ?? "不明";
        }
        catch
        {
            return "取得不可";
        }
    }

    private static string GetProblemDescription(uint problemCode)
    {
        var problemCodes = new Dictionary<uint, string>
        {
            {0, "正常"},
            {1, "正しく構成されていません"},
            {3, "ドライバーが破損している可能性があります"},
            {10, "デバイスを開始できません"},
            {12, "このデバイスが使用できる空きリソースが不足しています"},
            {18, "このデバイスのドライバーを再インストールしてください"},
            {22, "このデバイスは無効です"},
            {28, "このデバイスのドライバーがインストールされていません"},
            {31, "このデバイスは正しく動作していません"},
            {37, "Windows でこのデバイス用のドライバーを読み込むことができません"},
            {39, "ドライバーが破損しているか、ドライバーがありません"},
            {43, "以前のインスタンスが実行されているため、デバイスは停止されました"},
            {45, "現在、このハードウェア デバイスはコンピューターに接続されていません"}
        };

        return problemCodes.TryGetValue(problemCode, out var description)
            ? description
            : $"不明なエラー ({problemCode})";
    }

    private static List<AudioDeviceInfo> RemoveDuplicates(List<AudioDeviceInfo> devices)
    {
        var seenDeviceIds = new HashSet<string>();

        return devices.Where(device => seenDeviceIds.Add(device.DeviceId)).ToList();
    }
}