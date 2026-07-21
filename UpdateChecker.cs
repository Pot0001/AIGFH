using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace AIGFH;

internal static class UpdateChecker
{
    internal const string CurrentVersion = "1.1.6";
    internal static readonly string RepositoryPath = ReadRepositoryPath();
    internal static readonly string RepositoryUrl = "https://github.com/" + RepositoryPath;

    private static readonly string ApiRoot = "https://api.github.com/repos/" + RepositoryPath;
    private static readonly string RawVersionUrl = "https://raw.githubusercontent.com/" + RepositoryPath + "/main/VERSION";

    internal static void CheckAndShow()
    {
        try
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

            string latestText;
            string downloadPage;
            if (!TryGetLatestVersion(out latestText, out downloadPage))
            {
                ShowUnavailable();
                return;
            }

            var latest = ParseVersion(latestText);
            var current = ParseVersion(CurrentVersion);
            if (latest.CompareTo(current) > 0)
            {
                var answer = MessageBox.Show(
                    "发现新版本 " + FormatVersion(latest) + "\r\n\r\n" +
                    "当前版本：" + CurrentVersion + "\r\n" +
                    "是否前往下载？",
                    "检查更新",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Information);
                if (answer == DialogResult.Yes)
                    Open(downloadPage);
                return;
            }

            MessageBox.Show(
                "当前已是最新版本。\r\n\r\n" +
                "当前版本：" + CurrentVersion,
                "检查更新",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        catch
        {
            ShowUnavailable();
        }
    }

    private static bool TryGetLatestVersion(out string latestVersion, out string downloadPage)
    {
        latestVersion = String.Empty;
        downloadPage = RepositoryUrl + "/releases";

        var best = new Version(0, 0, 0);
        var bestText = String.Empty;
        var bestPage = downloadPage;

        // GitHub Release 是首选来源；没有创建 Release 时继续读取标签。
        string text;
        if (TryDownload(ApiRoot + "/releases/latest", "application/vnd.github+json", out text))
        {
            var tag = ReadJsonValue(text, "tag_name");
            var page = ReadJsonValue(text, "html_url");
            ConsiderVersion(tag, String.IsNullOrWhiteSpace(page) ? RepositoryUrl + "/releases" : page, ref best, ref bestText, ref bestPage);
        }

        // 标签接口会返回多个标签，取其中最大的版本，避免依赖接口排序。
        if (TryDownload(ApiRoot + "/tags?per_page=30", "application/vnd.github+json", out text))
        {
            foreach (var tag in ReadJsonValues(text, "name"))
                ConsiderVersion(tag, BuildPackageUrl(tag), ref best, ref bestText, ref bestPage);
        }

        // API 受限时，直接读取仓库中的轻量版本文件。
        if (TryDownload(RawVersionUrl, "text/plain", out text))
            ConsiderVersion(text, BuildPackageUrl(text), ref best, ref bestText, ref bestPage);

        // 最后使用 GitHub 自带的标签订阅源，不依赖 API 配额。
        if (TryDownload(RepositoryUrl + "/tags.atom", "application/atom+xml", out text))
        {
            var matches = Regex.Matches(text, @"/releases/tag/(?<tag>[^""<]+)", RegexOptions.IgnoreCase);
            var tags = new List<string>();
            foreach (Match match in matches)
                tags.Add(Uri.UnescapeDataString(match.Groups["tag"].Value));

            foreach (var tag in tags)
                ConsiderVersion(tag, BuildPackageUrl(tag), ref best, ref bestText, ref bestPage);
        }

        if (String.IsNullOrWhiteSpace(bestText)) return false;
        latestVersion = bestText;
        downloadPage = bestPage;
        return true;
    }

    private static bool TryDownload(string url, string accept, out string text)
    {
        text = String.Empty;
        try
        {
            using (var client = new TimeoutWebClient())
            {
                client.Encoding = System.Text.Encoding.UTF8;
                client.Headers[HttpRequestHeader.UserAgent] = "AI-Normalizer/" + CurrentVersion;
                client.Headers[HttpRequestHeader.Accept] = accept;
                client.Headers[HttpRequestHeader.AcceptLanguage] = "zh-CN,zh;q=0.9,en;q=0.5";
                client.Proxy = WebRequest.DefaultWebProxy;
                if (client.Proxy != null)
                    client.Proxy.Credentials = CredentialCache.DefaultCredentials;
                text = client.DownloadString(url);
                return !String.IsNullOrWhiteSpace(text);
            }
        }
        catch (WebException)
        {
            return false;
        }
    }

    private static void ShowUnavailable()
    {
        var answer = MessageBox.Show(
            "暂未获取到版本信息，请检查网络后重试。\r\n\r\n" +
            "是否打开项目下载页？",
            "检查更新",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);
        if (answer == DialogResult.Yes)
            Open(RepositoryUrl + "/releases");
    }

    private static List<string> ReadJsonValues(string json, string name)
    {
        var values = new List<string>();
        var pattern = "\\\"" + Regex.Escape(name) + "\\\"\\s*:\\s*\\\"(?<value>(?:\\\\.|[^\\\"])*)\\\"";
        foreach (Match match in Regex.Matches(json ?? String.Empty, pattern, RegexOptions.IgnoreCase))
            values.Add(Regex.Unescape(match.Groups["value"].Value));
        return values;
    }

    private static string ReadJsonValue(string json, string name)
    {
        var values = ReadJsonValues(json, name);
        return values.Count == 0 ? String.Empty : values[0];
    }

    private static void ConsiderVersion(string value, string page, ref Version best, ref string bestText, ref string bestPage)
    {
        var candidate = ParseVersion(value);
        if (!IsVersion(value) || candidate.CompareTo(best) <= 0) return;
        best = candidate;
        bestText = value.Trim();
        bestPage = String.IsNullOrWhiteSpace(page) ? RepositoryUrl + "/releases" : page;
    }

    private static string BuildPackageUrl(string versionText)
    {
        var version = ParseVersion(versionText);
        if (version.CompareTo(new Version(0, 0, 0)) <= 0)
            return RepositoryUrl + "/releases";
        var fileName = "AI规范化-Pot0001-" + FormatVersion(version) + ".exe";
        return RepositoryUrl + "/raw/main/dist/" + Uri.EscapeDataString(fileName);
    }

    private static bool IsVersion(string value)
    {
        return Regex.IsMatch(value == null ? String.Empty : value.Trim(), @"^v?\d+(?:\.\d+){1,3}(?:[-+][0-9A-Za-z.-]+)?$", RegexOptions.IgnoreCase);
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

    private static string FormatVersion(Version version)
    {
        return version == null ? String.Empty : version.ToString(version.Revision >= 0 ? 4 : version.Build >= 0 ? 3 : 2);
    }

    private static void Open(string url)
    {
        try { Process.Start(new ProcessStartInfo(url) { UseShellExecute = true }); }
        catch { Clipboard.SetText(url); }
    }

    private sealed class TimeoutWebClient : WebClient
    {
        protected override WebRequest GetWebRequest(Uri address)
        {
            var request = base.GetWebRequest(address);
            request.Timeout = 10000;
            var http = request as HttpWebRequest;
            if (http != null)
            {
                http.ReadWriteTimeout = 10000;
                http.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
            }
            return request;
        }
    }
}
