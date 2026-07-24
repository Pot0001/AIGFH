using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

internal static class Program
{
    private const string Clsid = "{2B868EC7-70D7-469B-A8AA-9B30F47EB33F}";
    private const string ProgId = "AIGFH.Connect";
    private const string LegacyProgId = "OfflineOfficeAddIn.Connect";
    private const string ManagedCategory = "{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}";
    private const string ProductName = "AI规范化";
    private const string PackageVersion = "1.1.8";
    private static readonly string ProjectUrl = "https://github.com/" + ReadResourceText("Payload.ProjectRepository.txt");
    private const string UninstallKey = "Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\AIGFH";
    private const int NetFramework48Release = 528040;
    private const int MoveFileDelayUntilReboot = 0x4;

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool MoveFileEx(string existingFileName, string newFileName, int flags);

    [STAThread]
    private static int Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        var silent = Array.Exists(args, value => String.Equals(value, "/silent", StringComparison.OrdinalIgnoreCase));
        var uninstall = Array.Exists(args, value => String.Equals(value, "/uninstall", StringComparison.OrdinalIgnoreCase));
        var install = Array.Exists(args, value => String.Equals(value, "/install", StringComparison.OrdinalIgnoreCase));
        var cleanupIndex = Array.FindIndex(args, value => String.Equals(value, "/cleanup", StringComparison.OrdinalIgnoreCase));
        if (cleanupIndex >= 0)
        {
            var cleanupPath = cleanupIndex + 1 < args.Length ? args[cleanupIndex + 1] : String.Empty;
            var parsedId = -1;
            var parentId = cleanupIndex + 2 < args.Length && Int32.TryParse(args[cleanupIndex + 2], out parsedId) ? parsedId : -1;
            return RunCleanup(cleanupPath, parentId);
        }
        if (!silent && !install && !uninstall)
        {
            Application.Run(new InstallerWindow());
            return 0;
        }
        return uninstall ? RunUninstall(silent) : RunInstall(silent);
    }

    private static int RunInstall(bool silent)
    {
        try
        {
            if (!HasNetFramework48())
            {
                const string runtimeMessage = "需要先安装 .NET Framework 4.8，再重新运行安装包。";
                if (!silent) MessageBox.Show(runtimeMessage, ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                else Console.Error.WriteLine(runtimeMessage);
                return 3;
            }
            var running = IsOfficeRunning();
            var installDirectory = InstallFiles();
            RemoveLegacyRegistration();
            RegisterCom(Path.Combine(installDirectory, "AIGFH.dll"));
            RegisterHost("Software\\Microsoft\\Office\\Word\\Addins\\" + ProgId);
            RegisterWpsWhitelist();
            ClearWordDisabledItems();
            WriteUninstallerInfo(installDirectory);

            var message = ProductName + " " + PackageVersion + " 已安装。\r\n\r\n本插件免费提供，已注册到 Microsoft Word 和 WPS。";
            message += running
                ? "\r\n\r\n检测到 Word/WPS 正在运行，请关闭全部窗口后重新打开。"
                : "\r\n\r\n现在可以打开 Word 或 WPS 使用。";
            if (!silent) MessageBox.Show(message, ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return 0;
        }
        catch (Exception error)
        {
            if (!silent) MessageBox.Show("安装失败：" + error.Message, ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            else Console.Error.WriteLine(error);
            return 1;
        }
    }

    private static int RunUninstall(bool silent)
    {
        try
        {
            if (IsOfficeRunning())
            {
                if (!silent) MessageBox.Show("请先关闭全部 Word 和 WPS 窗口，再继续卸载。", ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 2;
            }
            Uninstall();
            if (!silent) MessageBox.Show(ProductName + " 已卸载。", ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return 0;
        }
        catch (Exception error)
        {
            if (!silent) MessageBox.Show("卸载失败：" + error.Message, ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            else Console.Error.WriteLine(error);
            return 1;
        }
    }

    private static bool HasNetFramework48()
    {
        if (HasNetFramework48(RegistryView.Registry32)) return true;
        return Environment.Is64BitOperatingSystem && HasNetFramework48(RegistryView.Registry64);
    }

    private static bool HasNetFramework48(RegistryView view)
    {
        try
        {
            using (var localMachine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
            using (var framework = localMachine.OpenSubKey("SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full"))
                return framework != null && Convert.ToInt32(framework.GetValue("Release", 0)) >= NetFramework48Release;
        }
        catch { return false; }
    }

    private static int RunCleanup(string path, int parentId)
    {
        try
        {
            var allowedRoot = Path.GetFullPath(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIGFH"));
            var target = Path.GetFullPath(path ?? String.Empty);
            if (!String.Equals(target, allowedRoot, StringComparison.OrdinalIgnoreCase)) return 2;
            if (parentId > 0)
            {
                try { Process.GetProcessById(parentId).WaitForExit(); } catch { }
            }
            for (var attempt = 0; attempt < 40 && Directory.Exists(target); attempt++)
            {
                try { Directory.Delete(target, true); }
                catch { Thread.Sleep(250); }
            }
            var self = Assembly.GetExecutingAssembly().Location;
            try { File.Delete(self); } catch { MoveFileEx(self, null, MoveFileDelayUntilReboot); }
            return Directory.Exists(target) ? 1 : 0;
        }
        catch { return 1; }
    }

    private sealed class InstallerWindow : Form
    {
        private readonly Label _status;
        private readonly Label _localVersion;
        private readonly Button _install;
        private readonly Button _uninstall;

        internal InstallerWindow()
        {
            Text = "AI规范化 - 安装与维护";
            Font = new System.Drawing.Font("Microsoft YaHei UI", 9F);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new System.Drawing.Size(560, 420);
            BackColor = System.Drawing.Color.White;

            Controls.Add(new Label { Text = "AI规范化", Font = new System.Drawing.Font("Microsoft YaHei UI", 23F, System.Drawing.FontStyle.Bold), ForeColor = System.Drawing.Color.FromArgb(30, 64, 175), AutoSize = true, Location = new System.Drawing.Point(34, 28) });
            Controls.Add(new Label { Text = "Word/WPS 文本、TeX 公式、表格与试卷排版工具", AutoSize = true, ForeColor = System.Drawing.Color.FromArgb(71, 85, 105), Location = new System.Drawing.Point(38, 82) });
            Controls.Add(new Label { Text = "安装包版本：" + PackageVersion, AutoSize = true, ForeColor = System.Drawing.Color.FromArgb(71, 85, 105), Location = new System.Drawing.Point(38, 112) });
            _localVersion = new Label { AutoSize = true, ForeColor = System.Drawing.Color.FromArgb(71, 85, 105), Location = new System.Drawing.Point(248, 112) };
            Controls.Add(_localVersion);

            Controls.Add(new Label { Text = "本插件免费提供，可规范 AI 文本、TeX 公式、表格和试卷版式。\r\n项目介绍、下载与问题反馈：", AutoSize = false, Width = 480, Height = 48, ForeColor = System.Drawing.Color.FromArgb(51, 65, 85), Location = new System.Drawing.Point(38, 148) });
            var projectLink = new LinkLabel { Text = ProjectUrl, AutoSize = true, LinkColor = System.Drawing.Color.FromArgb(37, 99, 235), Location = new System.Drawing.Point(38, 196) };
            projectLink.LinkClicked += (_, __) => OpenProjectUrl();
            Controls.Add(projectLink);

            _status = new Label { AutoSize = false, Width = 480, Height = 38, Location = new System.Drawing.Point(38, 244), ForeColor = System.Drawing.Color.FromArgb(51, 65, 85) };
            Controls.Add(_status);

            _install = Button("安装", 38, System.Drawing.Color.FromArgb(37, 99, 235), System.Drawing.Color.White);
            _uninstall = Button("卸载", 198, System.Drawing.Color.FromArgb(241, 245, 249), System.Drawing.Color.FromArgb(30, 41, 59));
            var close = Button("退出", 358, System.Drawing.Color.FromArgb(241, 245, 249), System.Drawing.Color.FromArgb(30, 41, 59));
            _install.Click += (_, __) => Execute(true);
            _uninstall.Click += (_, __) => Execute(false);
            close.Click += (_, __) => Close();
            Controls.Add(_install); Controls.Add(_uninstall); Controls.Add(close);
            RefreshState();
        }

        private static Button Button(string text, int left, System.Drawing.Color back, System.Drawing.Color fore)
        {
            var button = new Button { Text = text, Width = 136, Height = 42, Location = new System.Drawing.Point(left, 326), BackColor = back, ForeColor = fore, FlatStyle = FlatStyle.Flat };
            button.FlatAppearance.BorderSize = 0;
            return button;
        }

        private void Execute(bool install)
        {
            _install.Enabled = false; _uninstall.Enabled = false;
            _status.Text = install ? "正在配置 Word 和 WPS，请稍候……" : "正在移除，请稍候……";
            Refresh();
            var code = install ? RunInstall(true) : RunUninstall(true);
            RefreshState();
            if (code == 0)
            {
                if (install) _status.Text = "配置完成。重新打开 Word 或 WPS 后即可使用。";
                else
                {
                    MessageBox.Show("已从本机移除。", ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Close();
                }
            }
            else if (!install && code == 2)
                _status.Text = "请先关闭全部 Word 和 WPS 窗口，再重试。";
            else if (install && code == 3)
                _status.Text = "需要先安装 .NET Framework 4.8。";
            else
                _status.Text = install ? "配置未完成，请关闭 Word 和 WPS 后重试。" : "移除未完成，请重试。";
        }

        private void RefreshState()
        {
            var installedVersion = GetInstalledVersion();
            _localVersion.Text = "本机版本：" + (installedVersion ?? "未安装");
            _uninstall.Enabled = installedVersion != null;
            _install.Enabled = true;

            if (installedVersion == null)
            {
                _install.Text = "安装";
                _status.Text = "尚未安装。安装范围为当前 Windows 用户。";
                return;
            }

            var comparison = CompareVersions(installedVersion, PackageVersion);
            if (comparison < 0)
            {
                _install.Text = "升级";
                _status.Text = "本机已有旧版本，可升级到 " + PackageVersion + "。";
            }
            else if (comparison == 0)
            {
                if (IsInstallationHealthy())
                {
                    _install.Text = "重新安装";
                    _status.Text = "本机已安装当前版本。";
                }
                else
                {
                    _install.Text = "修复";
                    _status.Text = "检测到组件缺失，重新配置即可。";
                }
            }
            else
            {
                _install.Text = "降级安装";
                _status.Text = "本机版本更新。建议保留；如需使用安装包版本，可降级安装。";
            }
        }

        private static string GetInstalledVersion()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(UninstallKey))
            {
                if (key == null) return null;
                var value = Convert.ToString(key.GetValue("DisplayVersion"));
                return String.IsNullOrWhiteSpace(value) ? "未知" : value.Trim();
            }
        }

        private static bool IsInstallationHealthy()
        {
            try
            {
                string installDirectory;
                using (var uninstall = Registry.CurrentUser.OpenSubKey(UninstallKey))
                    installDirectory = uninstall == null ? null : Convert.ToString(uninstall.GetValue("InstallLocation"));
                if (String.IsNullOrWhiteSpace(installDirectory)) return false;

                var dllPath = Path.Combine(installDirectory, "AIGFH.dll");
                if (!File.Exists(dllPath)) return false;

                using (var word = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Office\\Word\\Addins\\" + ProgId))
                {
                    if (word == null || Convert.ToInt32(word.GetValue("LoadBehavior", 0)) != 3) return false;
                }

                using (var whitelist = Registry.CurrentUser.OpenSubKey("Software\\Kingsoft\\Office\\WPS\\AddinsWL"))
                {
                    if (whitelist == null || !Array.Exists(whitelist.GetValueNames(), name => String.Equals(name, ProgId, StringComparison.OrdinalIgnoreCase)))
                        return false;
                }

                if (!IsComRegistrationHealthy(RegistryView.Registry32, dllPath)) return false;
                if (Environment.Is64BitOperatingSystem && !IsComRegistrationHealthy(RegistryView.Registry64, dllPath)) return false;
                return true;
            }
            catch { return false; }
        }

        private static bool IsComRegistrationHealthy(RegistryView view, string expectedDllPath)
        {
            using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
            using (var server = currentUser.OpenSubKey("Software\\Classes\\CLSID\\" + Clsid + "\\InprocServer32"))
            {
                if (server == null) return false;
                var codeBase = Convert.ToString(server.GetValue("CodeBase"));
                Uri uri;
                var registeredPath = Uri.TryCreate(codeBase, UriKind.Absolute, out uri) && uri.IsFile
                    ? uri.LocalPath
                    : codeBase;
                if (String.IsNullOrWhiteSpace(registeredPath) || !File.Exists(registeredPath)) return false;
                return String.Equals(Path.GetFullPath(registeredPath), Path.GetFullPath(expectedDllPath), StringComparison.OrdinalIgnoreCase);
            }
        }

        private static int CompareVersions(string left, string right)
        {
            Version leftVersion;
            Version rightVersion;
            if (Version.TryParse((left ?? String.Empty).TrimStart('v', 'V'), out leftVersion) &&
                Version.TryParse((right ?? String.Empty).TrimStart('v', 'V'), out rightVersion))
            {
                var leftParts = new[] { leftVersion.Major, leftVersion.Minor, Math.Max(0, leftVersion.Build), Math.Max(0, leftVersion.Revision) };
                var rightParts = new[] { rightVersion.Major, rightVersion.Minor, Math.Max(0, rightVersion.Build), Math.Max(0, rightVersion.Revision) };
                for (var index = 0; index < leftParts.Length; index++)
                {
                    if (leftParts[index] < rightParts[index]) return -1;
                    if (leftParts[index] > rightParts[index]) return 1;
                }
                return 0;
            }
            return String.Equals(left, right, StringComparison.OrdinalIgnoreCase) ? 0 : -1;
        }

        private static void OpenProjectUrl()
        {
            try { Process.Start(new ProcessStartInfo(ProjectUrl) { UseShellExecute = true }); }
            catch { Clipboard.SetText(ProjectUrl); }
        }
    }

    private static string InstallFiles()
    {
        var productRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIGFH");
        var root = Path.Combine(productRoot, PackageVersion);
        var payloadHash = GetResourceHash("Payload.AIGFH.dll").Substring(0, 12);
        var directory = Path.Combine(root, payloadHash);
        Directory.CreateDirectory(directory);
        Extract("Payload.AIGFH.dll", Path.Combine(directory, "AIGFH.dll"));
        Extract("Payload.使用说明.txt", Path.Combine(directory, "使用说明.txt"));
        Extract("Payload.AI提示词.txt", Path.Combine(directory, "AI提示词.txt"));

        var maintenance = Path.Combine(root, "AI规范化-维护.exe");
        var currentExecutable = Assembly.GetExecutingAssembly().Location;
        if (!String.Equals(Path.GetFullPath(currentExecutable), Path.GetFullPath(maintenance), StringComparison.OrdinalIgnoreCase))
            File.Copy(currentExecutable, maintenance, true);
        File.WriteAllText(Path.Combine(root, "当前版本.txt"), PackageVersion + "\r\n" + directory, new UTF8Encoding(true));

        foreach (var oldDirectory in Directory.GetDirectories(root))
        {
            if (String.Equals(Path.GetFullPath(oldDirectory), Path.GetFullPath(directory), StringComparison.OrdinalIgnoreCase)) continue;
            try { Directory.Delete(oldDirectory, true); } catch { }
        }
        foreach (var oldVersion in Directory.GetDirectories(productRoot))
        {
            if (String.Equals(Path.GetFullPath(oldVersion), Path.GetFullPath(root), StringComparison.OrdinalIgnoreCase)) continue;
            try { Directory.Delete(oldVersion, true); } catch { }
        }
        var legacyRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIDocumentNormalizer", "Y01");
        try { if (Directory.Exists(legacyRoot)) Directory.Delete(legacyRoot, true); } catch { }
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\AI规范化Y01", false);
        return directory;
    }

    private static void Extract(string resourceName, string destination)
    {
        using (var source = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
        {
            if (source == null) throw new InvalidOperationException("安装包缺少资源：" + resourceName);
            using (var target = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.Read)) source.CopyTo(target);
        }
    }

    private static string ReadResourceText(string resourceName)
    {
        using (var source = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
        using (var reader = source == null ? null : new StreamReader(source, Encoding.UTF8))
        {
            if (reader == null) throw new InvalidOperationException("安装包缺少项目地址配置。");
            return reader.ReadToEnd().Trim();
        }
    }

    private static string GetResourceHash(string resourceName)
    {
        using (var source = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
        using (var sha = SHA256.Create())
        {
            if (source == null) throw new InvalidOperationException("安装包缺少核心文件。");
            return BitConverter.ToString(sha.ComputeHash(source)).Replace("-", String.Empty);
        }
    }

    private static void RegisterCom(string dllPath)
    {
        var codeBase = new Uri(Path.GetFullPath(dllPath)).AbsoluteUri;
        var payload = System.Reflection.AssemblyName.GetAssemblyName(dllPath);
        var assemblyFullName = payload.FullName;
        var assemblyVersion = payload.Version == null ? "1.0.0.0" : payload.Version.ToString();

        // CLSID is a redirected key on 64-bit Windows.  Register through the
        // explicit views instead of writing the reserved Wow6432Node path.
        // The AnyCPU managed add-in can then be activated by either x86 or x64
        // Word/WPS through the matching mscoree.dll registration.
        RegisterComView(RegistryView.Registry32, codeBase, assemblyFullName, assemblyVersion);
        if (Environment.Is64BitOperatingSystem)
            RegisterComView(RegistryView.Registry64, codeBase, assemblyFullName, assemblyVersion);
    }

    private static void RegisterComView(RegistryView view, string codeBase, string assemblyFullName, string assemblyVersion)
    {
        using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
        {
            const string classesRoot = "Software\\Classes";
            var clsidRoot = classesRoot + "\\CLSID\\" + Clsid;
            // Remove versioned InprocServer32 entries left by older builds.
            // mscoree may otherwise select a stale highest-version CodeBase.
            currentUser.DeleteSubKeyTree(clsidRoot, false);
            currentUser.DeleteSubKeyTree(classesRoot + "\\" + ProgId, false);
            using (var clsid = currentUser.CreateSubKey(clsidRoot)) clsid.SetValue(null, ProgId);
            using (var server = currentUser.CreateSubKey(clsidRoot + "\\InprocServer32"))
            {
                server.SetValue(null, "mscoree.dll");
                server.SetValue("ThreadingModel", "Both");
                server.SetValue("Class", ProgId);
                server.SetValue("Assembly", assemblyFullName);
                server.SetValue("RuntimeVersion", "v4.0.30319");
                server.SetValue("CodeBase", codeBase);
            }
            using (var version = currentUser.CreateSubKey(clsidRoot + "\\InprocServer32\\" + assemblyVersion))
            {
                version.SetValue("Class", ProgId);
                version.SetValue("Assembly", assemblyFullName);
                version.SetValue("RuntimeVersion", "v4.0.30319");
                version.SetValue("CodeBase", codeBase);
            }
            using (var progId = currentUser.CreateSubKey(clsidRoot + "\\ProgId")) progId.SetValue(null, ProgId);
            using (currentUser.CreateSubKey(clsidRoot + "\\Implemented Categories\\" + ManagedCategory)) { }
            using (var prog = currentUser.CreateSubKey(classesRoot + "\\" + ProgId)) prog.SetValue(null, ProgId);
            using (var progClsid = currentUser.CreateSubKey(classesRoot + "\\" + ProgId + "\\CLSID")) progClsid.SetValue(null, Clsid);
            currentUser.DeleteSubKeyTree(classesRoot + "\\" + LegacyProgId, false);
        }
    }

    private static void RemoveLegacyRegistration()
    {
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Microsoft\\Office\\Word\\Addins\\" + LegacyProgId, false);
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Kingsoft\\Office\\WPS\\Addins\\" + LegacyProgId, false);
        // Older installers wrote a non-standard WPS Addins subkey.  WPS COM
        // add-ins are discovered from the Word Addins key and authorized by
        // the AddinsWL value below.
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Kingsoft\\Office\\WPS\\Addins\\" + ProgId, false);
        RemoveWpsWhitelistValue(LegacyProgId);
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\AIGFH-AIDocumentNormalizer", false);
    }

    private static void RegisterHost(string path)
    {
        using (var key = Registry.CurrentUser.CreateSubKey(path))
        {
            key.SetValue("FriendlyName", ProductName + " v" + PackageVersion);
            key.SetValue("Description", "免费｜Word/WPS 文本、TeX 公式、表格与试卷排版工具｜" + ProjectUrl);
            key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
            // This is a classic IDTExtensibility2 COM add-in, not a VSTO
            // deployment.  CommandLineSafe is required by Office/WPS; startup
            // only connects the add-in and does not run document commands or
            // display modal UI, so command-line host startup is safe.
            // Do not add a Manifest value, which would invoke the VSTO loader.
            key.SetValue("CommandLineSafe", 1, RegistryValueKind.DWord);
        }
    }

    private static void RegisterWpsWhitelist()
    {
        // WPS Writer reads the normal Microsoft Word Addins registration and
        // requires the ProgID in its per-user COM add-in whitelist.  The value
        // is REG_SZ with empty data; this path is shared by x86 and x64 WPS.
        using (var key = Registry.CurrentUser.CreateSubKey("Software\\Kingsoft\\Office\\WPS\\AddinsWL"))
            key.SetValue(ProgId, String.Empty, RegistryValueKind.String);
    }

    private static void RemoveWpsWhitelistValue(string progId)
    {
        using (var key = Registry.CurrentUser.OpenSubKey("Software\\Kingsoft\\Office\\WPS\\AddinsWL", true))
            if (key != null) key.DeleteValue(progId, false);
    }

    private static void ClearWordDisabledItems()
    {
        using (var office = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Office", true))
        {
            if (office == null) return;
            foreach (var version in office.GetSubKeyNames())
            using (var disabled = office.OpenSubKey(version + "\\Word\\Resiliency\\DisabledItems", true))
            {
                if (disabled == null) continue;
                foreach (var name in disabled.GetValueNames())
                {
                    var bytes = disabled.GetValue(name) as byte[];
                    if (bytes == null) continue;
                    var unicode = Encoding.Unicode.GetString(bytes);
                    var ascii = Encoding.ASCII.GetString(bytes);
                    if (unicode.IndexOf(ProgId, StringComparison.OrdinalIgnoreCase) >= 0 || ascii.IndexOf(ProgId, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        unicode.IndexOf(LegacyProgId, StringComparison.OrdinalIgnoreCase) >= 0 || ascii.IndexOf(LegacyProgId, StringComparison.OrdinalIgnoreCase) >= 0)
                        disabled.DeleteValue(name, false);
                }
            }
        }
    }

    private static void WriteUninstallerInfo(string installDirectory)
    {
        using (var key = Registry.CurrentUser.CreateSubKey(UninstallKey))
        {
            var installer = Path.Combine(Directory.GetParent(installDirectory).FullName, "AI规范化-维护.exe");
            key.SetValue("DisplayName", ProductName);
            key.SetValue("DisplayVersion", PackageVersion);
            key.SetValue("Publisher", ProductName);
            key.SetValue("URLInfoAbout", ProjectUrl);
            key.SetValue("InstallLocation", installDirectory);
            key.SetValue("UninstallString", "\"" + installer + "\" /uninstall");
            key.SetValue("NoModify", 1, RegistryValueKind.DWord);
            key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
        }
    }

    private static void Uninstall()
    {
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Microsoft\\Office\\Word\\Addins\\" + ProgId, false);
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Kingsoft\\Office\\WPS\\Addins\\" + ProgId, false);
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Microsoft\\Office\\Word\\Addins\\" + LegacyProgId, false);
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Kingsoft\\Office\\WPS\\Addins\\" + LegacyProgId, false);
        RemoveWpsWhitelistValue(ProgId);
        RemoveWpsWhitelistValue(LegacyProgId);
        DeleteComView(RegistryView.Registry32);
        if (Environment.Is64BitOperatingSystem) DeleteComView(RegistryView.Registry64);
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\AIGFH", false);
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\AIGFH-AIDocumentNormalizer", false);
        Registry.CurrentUser.DeleteSubKeyTree("Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\AI规范化Y01", false);
        var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIGFH");
        try { if (Directory.Exists(root)) Directory.Delete(root, true); }
        catch { StartCleanupHelper(root); }
        var legacyRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AIDocumentNormalizer", "Y01");
        try { if (Directory.Exists(legacyRoot)) Directory.Delete(legacyRoot, true); } catch { }
    }

    private static void StartCleanupHelper(string root)
    {
        try
        {
            var helper = Path.Combine(Path.GetTempPath(), "AI规范化-清理-" + Guid.NewGuid().ToString("N") + ".exe");
            File.Copy(Assembly.GetExecutingAssembly().Location, helper, true);
            Process.Start(new ProcessStartInfo
            {
                FileName = helper,
                Arguments = "/cleanup \"" + root + "\" " + Process.GetCurrentProcess().Id,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            });
        }
        catch { }
    }

    private static void DeleteComView(RegistryView view)
    {
        using (var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
        {
            const string classesRoot = "Software\\Classes";
            currentUser.DeleteSubKeyTree(classesRoot + "\\CLSID\\" + Clsid, false);
            currentUser.DeleteSubKeyTree(classesRoot + "\\" + ProgId, false);
            currentUser.DeleteSubKeyTree(classesRoot + "\\" + LegacyProgId, false);
        }
    }

    private static bool IsOfficeRunning()
    {
        return Process.GetProcessesByName("WINWORD").Length > 0 || Process.GetProcessesByName("wps").Length > 0;
    }
}
