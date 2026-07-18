using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace AIGFH;

internal static class UpdateChecker
{
    internal const string CurrentVersion = "1.1.0";
    internal static readonly string RepositoryUrl = "https://github.com/" + ReadRepositoryPath();
    private static readonly string ApiRoot = "https://api.github.com/repos/" + ReadRepositoryPath();

    internal static void CheckAndShow()
    {
        try
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
            string json;
            using (var client = new WebClient())
            {
                client.Headers[HttpRequestHeader.UserAgent] = "AIGFH/" + CurrentVersion;
                client.Headers[HttpRequestHeader.Accept] = "application/vnd.github+json";
                try { json = client.DownloadString(ApiRoot + "/releases/latest"); }
                catch (WebException) { json = client.DownloadString(ApiRoot + "/tags?per_page=1"); }
            }
            var tag = ReadJsonValue(json, "tag_name");
            if (String.IsNullOrWhiteSpace(tag)) tag = ReadJsonValue(json, "name");
            var page = ReadJsonValue(json, "html_url");
            if (String.IsNullOrWhiteSpace(page)) page = RepositoryUrl + "/releases";
            var latest = ParseVersion(tag);
            var current = ParseVersion(CurrentVersion);
            if (latest.CompareTo(current) > 0)
            {
                if (MessageBox.Show("发现新版本 " + tag + "。\r\n\r\n当前版本：" + CurrentVersion + "\r\n是否打开版本下载页？", "检查更新", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                    Open(page);
            }
            else MessageBox.Show("当前已是最新版本。\r\n\r\n当前版本：" + CurrentVersion + "\r\n最新版本：" + tag, "检查更新", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception)
        {
            if (MessageBox.Show("暂时没有获取到版本信息。\r\n\r\n是否打开项目主页查看？", "检查更新", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                Open(RepositoryUrl + "/releases");
        }
    }

    private static string ReadJsonValue(string json, string name)
    {
        var match = Regex.Match(json ?? String.Empty, "\\\"" + Regex.Escape(name) + "\\\"\\s*:\\s*\\\"(?<value>(?:\\\\.|[^\\\"])*)\\\"", RegexOptions.IgnoreCase);
        return match.Success ? Regex.Unescape(match.Groups["value"].Value) : String.Empty;
    }

    private static string ReadRepositoryPath()
    {
        using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("AIGFH.ProjectRepository.txt"))
        using (var reader = stream == null ? null : new StreamReader(stream))
        {
            var value = reader == null ? String.Empty : reader.ReadToEnd().Trim();
            if (!Regex.IsMatch(value, @"^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$"))
                throw new InvalidDataException("项目地址配置不正确。");
            return value;
        }
    }

    private static Version ParseVersion(string value)
    {
        var match = Regex.Match(value ?? String.Empty, @"\d+(?:\.\d+){0,3}");
        Version version;
        return match.Success && Version.TryParse(match.Value, out version) ? version : new Version(0, 0, 0);
    }

    private static void Open(string url)
    {
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
        catch { Clipboard.SetText(url); }
    }
}
