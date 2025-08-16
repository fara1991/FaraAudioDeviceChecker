namespace FaraAudioDeviceChecker.Utilities;

using System.Management;

public static class WmiHelper
{
    public static string GetPropertyValue(ManagementObject obj, string propertyName)
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

    public static string EscapeWqlString(string input)
    {
        if (string.IsNullOrEmpty(input))
            return input;

        return input
            .Replace("\\", @"\\") // バックスラッシュ
            .Replace("'", "''") // シングルクォート
            .Replace("\"", "\\\"") // ダブルクォート
            .Replace("%", "[%]") // パーセント
            .Replace("_", "[_]") // アンダースコア
            .Replace("&", "^&"); // アンパサンド
    }
}