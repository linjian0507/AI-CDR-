using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace AIVectorInstaller
{
    internal sealed class CdrInstallation
    {
        public string DisplayName { get; private set; }
        public string ProgramsDirectory { get; private set; }
        public string ExecutablePath { get; private set; }
        public string AddonDirectory
        {
            get { return Path.Combine(ProgramsDirectory, "Addons", "AIVectorHelper"); }
        }

        public CdrInstallation(string displayName, string programsDirectory, string executablePath)
        {
            DisplayName = displayName;
            ProgramsDirectory = programsDirectory;
            ExecutablePath = executablePath;
        }

        public override string ToString()
        {
            return DisplayName + "  [" + ProgramsDirectory + "]";
        }
    }

    internal static class CdrDetector
    {
        private static readonly Regex VersionNumber = new Regex(@"(?<!\d)(1[4-9]|2[0-9])(?:\.(\d+))?(?!\d)", RegexOptions.Compiled);

        public static List<CdrInstallation> Detect()
        {
            var result = new Dictionary<string, CdrInstallation>(StringComparer.OrdinalIgnoreCase);
            foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                ScanRegistry(result, RegistryHive.LocalMachine, view);
                ScanRegistry(result, RegistryHive.CurrentUser, view);
            }

            ScanCommonFolders(result);
            return result.Values
                .OrderBy(x => VersionSortKey(x.DisplayName))
                .ThenBy(x => x.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private static void ScanRegistry(
            IDictionary<string, CdrInstallation> result,
            RegistryHive hive,
            RegistryView view)
        {
            try
            {
                using (var baseKey = RegistryKey.OpenBaseKey(hive, view))
                using (var root = baseKey.OpenSubKey(@"SOFTWARE\Corel\CorelDRAW"))
                {
                    if (root == null) return;
                    foreach (var versionName in root.GetSubKeyNames())
                    {
                        using (var versionKey = root.OpenSubKey(versionName))
                        {
                            if (versionKey == null) continue;
                            var programsDir = ReadString(versionKey, "ProgramsDir");
                            var configDir = ReadString(versionKey, "ConfigDir");
                            var suiteVersion = ReadString(versionKey, "SuiteVersion");
                            var displayVersion = versionName;
                            var candidates = new List<string>();

                            AddProgramCandidate(candidates, programsDir);
                            if (!string.IsNullOrWhiteSpace(configDir))
                            {
                                var installRoot = Directory.GetParent(configDir.TrimEnd('\\', '/'));
                                if (installRoot != null)
                                {
                                    AddProgramCandidate(candidates, Path.Combine(installRoot.FullName, "Programs64"));
                                    AddProgramCandidate(candidates, Path.Combine(installRoot.FullName, "Programs"));
                                    AddProgramCandidate(candidates, Path.Combine(installRoot.FullName, "program"));
                                }
                            }

                            foreach (var candidate in candidates.Distinct(StringComparer.OrdinalIgnoreCase))
                            {
                                AddExecutable(result, candidate, displayVersion, suiteVersion);
                            }
                        }
                    }
                }
            }
            catch
            {
                // 某些旧版注册表项可能没有读取权限，继续扫描其他来源。
            }
        }

        private static void ScanCommonFolders(IDictionary<string, CdrInstallation> result)
        {
            var roots = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), "Corel")
            };

            foreach (var root in roots.Where(Directory.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    foreach (var directory in Directory.EnumerateDirectories(root, "*", SearchOption.TopDirectoryOnly))
                    {
                        var name = Path.GetFileName(directory) ?? "";
                        if (name.IndexOf("Corel", StringComparison.OrdinalIgnoreCase) < 0) continue;
                        foreach (var programName in new[] { "Programs64", "Programs", "program" })
                        {
                            var candidate = Path.Combine(directory, programName);
                            AddExecutable(result, candidate, "", "");
                        }
                    }
                }
                catch
                {
                    // 忽略不可访问目录。
                }
            }
        }

        private static void AddProgramCandidate(ICollection<string> candidates, string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            var trimmed = value.Trim().TrimEnd('\\', '/');
            if (trimmed.Length > 0) candidates.Add(trimmed);
        }

        private static void AddExecutable(
            IDictionary<string, CdrInstallation> result,
            string programsDirectory,
            string registryVersion,
            string suiteVersion)
        {
            if (string.IsNullOrWhiteSpace(programsDirectory)) return;
            var executable = Path.Combine(programsDirectory, "CorelDRW.exe");
            if (!File.Exists(executable)) return;

            var fullPrograms = Path.GetFullPath(programsDirectory).TrimEnd('\\');
            var display = BuildDisplayName(fullPrograms, registryVersion, suiteVersion);
            if (!result.ContainsKey(executable))
            {
                result.Add(executable, new CdrInstallation(display, fullPrograms, executable));
            }
        }

        private static string BuildDisplayName(string programsDirectory, string registryVersion, string suiteVersion)
        {
            var major = 0;
            var minor = 0;
            if (!TryParseCdrVersion(registryVersion, out major, out minor))
            {
                TryParseCdrVersion(programsDirectory, out major, out minor);
            }

            var product = ProductName(major, minor);
            var is64 = programsDirectory.EndsWith("Programs64", StringComparison.OrdinalIgnoreCase)
                || programsDirectory.IndexOf("x64", StringComparison.OrdinalIgnoreCase) >= 0;
            var architecture = is64 ? "64位" : "32位";
            if (string.IsNullOrWhiteSpace(product))
            {
                product = string.IsNullOrWhiteSpace(suiteVersion) ? "CorelDRAW" : "CorelDRAW " + suiteVersion;
            }
            return product + " " + architecture;
        }

        private static bool TryParseCdrVersion(string text, out int major, out int minor)
        {
            major = 0;
            minor = 0;
            var match = VersionNumber.Match(text ?? "");
            if (!match.Success) return false;
            int.TryParse(match.Groups[1].Value, out major);
            if (match.Groups[2].Success) int.TryParse(match.Groups[2].Value, out minor);
            return major > 0;
        }

        private static string ProductName(int major, int minor)
        {
            if (major >= 14 && major <= 18) return "X" + (major - 10);
            if (major == 19) return "2017";
            if (major == 24) return minor >= 3 ? "2023" : "2022";
            if (major >= 25 && major <= 30) return (1999 + major).ToString();
            if (major >= 20 && major <= 23) return (1998 + major).ToString();
            return "";
        }

        private static int VersionSortKey(string displayName)
        {
            var match = Regex.Match(displayName ?? "", @"(X[4-8]|20\d{2})");
            if (!match.Success) return int.MaxValue;
            if (match.Value.StartsWith("X", StringComparison.OrdinalIgnoreCase))
            {
                int x;
                return int.TryParse(match.Value.Substring(1), out x) ? x : int.MaxValue;
            }
            int year;
            return int.TryParse(match.Value, out year) ? year : int.MaxValue;
        }

        private static string ReadString(RegistryKey key, string name)
        {
            var value = key.GetValue(name);
            return value == null ? "" : Convert.ToString(value);
        }
    }

    internal sealed class InstallerForm : Form
    {
        private readonly FlowLayoutPanel _versionPanel;
        private readonly Label _statusLabel;
        private readonly Button _installButton;
        private readonly Button _uninstallButton;
        private readonly List<CheckBox> _versionChecks = new List<CheckBox>();
        private List<CdrInstallation> _installations = new List<CdrInstallation>();

        public InstallerForm()
        {
            Text = "AI矢量助手安装程序 v2.3.9";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(470, 390);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            BackColor = Color.White;

            var title = new Label
            {
                AutoSize = true,
                Text = "AI矢量助手 v2.3.9",
                Font = new Font("Microsoft YaHei UI", 18F, FontStyle.Bold),
                ForeColor = Color.FromArgb(244, 61, 111),
                Location = new Point(150, 12)
            };
            Controls.Add(title);

            var subtitle = new Label
            {
                AutoSize = true,
                Text = "CorelDRAW 插件安装 / 卸载",
                ForeColor = Color.FromArgb(90, 90, 90),
                Location = new Point(145, 43)
            };
            Controls.Add(subtitle);

            var group = new GroupBox
            {
                Text = "请选择已安装的 CDR 版本",
                Location = new Point(14, 70),
                Size = new Size(442, 205),
                Padding = new Padding(10, 22, 10, 8)
            };
            Controls.Add(group);

            _versionPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(2),
                BackColor = Color.White
            };
            group.Controls.Add(_versionPanel);

            _installButton = MakeButton("安装", new Point(286, 290), InstallSelected);
            _uninstallButton = MakeButton("卸载选中", new Point(172, 290), UninstallSelected);
            Controls.Add(_uninstallButton);
            Controls.Add(_installButton);

            var refresh = MakeButton("重新扫描", new Point(58, 290), (s, e) => RefreshVersions());
            Controls.Add(refresh);

            var help = MakeButton("安装说明", new Point(168, 337), ShowHelp);
            Controls.Add(help);
            var exit = MakeButton("退出", new Point(286, 337), (s, e) => Close());
            Controls.Add(exit);

            _statusLabel = new Label
            {
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = Color.FromArgb(80, 80, 80),
                Location = new Point(18, 315),
                Size = new Size(430, 20)
            };
            Controls.Add(_statusLabel);

            Shown += (s, e) => RefreshVersions();
        }

        private Button MakeButton(string text, Point location, EventHandler handler)
        {
            var button = new Button
            {
                Text = text,
                Location = location,
                Size = new Size(92, 32),
                FlatStyle = FlatStyle.Standard
            };
            button.Click += handler;
            return button;
        }

        private void RefreshVersions()
        {
            _versionPanel.Controls.Clear();
            _versionChecks.Clear();
            _installations = CdrDetector.Detect();
            foreach (var installation in _installations)
            {
                var check = new CheckBox
                {
                    Text = installation.DisplayName,
                    Tag = installation,
                    Checked = true,
                    AutoSize = false,
                    Width = 195,
                    Height = 34,
                    Margin = new Padding(2, 2, 2, 2),
                    FlatStyle = FlatStyle.Standard
                };
                check.MouseEnter += (s, e) =>
                {
                    var item = ((CheckBox)s).Tag as CdrInstallation;
                    ((CheckBox)s).ToolTipText(item == null ? "" : item.ProgramsDirectory);
                };
                _versionChecks.Add(check);
                _versionPanel.Controls.Add(check);
            }

            if (_installations.Count == 0)
            {
                var empty = new Label
                {
                    AutoSize = true,
                    Text = "未自动找到 CDR。请确认已安装 CorelDRAW 后重新扫描。",
                    ForeColor = Color.DarkRed,
                    Padding = new Padding(2, 8, 2, 2)
                };
                _versionPanel.Controls.Add(empty);
                _statusLabel.Text = "未找到可用的 CDR 安装。";
            }
            else
            {
                _statusLabel.Text = "已找到 " + _installations.Count + " 个 CDR 安装，可多选。";
            }
        }

        private List<CdrInstallation> SelectedInstallations()
        {
            return _versionChecks
                .Where(x => x.Checked && x.Tag is CdrInstallation)
                .Select(x => (CdrInstallation)x.Tag)
                .ToList();
        }

        private void InstallSelected(object sender, EventArgs e)
        {
            var selected = SelectedInstallations();
            if (selected.Count == 0)
            {
                MessageBox.Show(this, "请至少选择一个 CDR 版本。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (CorelDrawRunning())
            {
                MessageBox.Show(this, "请先关闭所有 CorelDRAW 窗口，再执行安装。", "无法安装", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var payload = ExtractPayload();
                foreach (var installation in selected)
                {
                    InstallTo(installation.AddonDirectory, payload);
                }
                _statusLabel.Text = "安装完成，请重新启动 CorelDRAW。";
                MessageBox.Show(this, "安装完成。\r\n\r\n请重新启动 CorelDRAW 使插件生效。", "安装成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _statusLabel.Text = "安装失败：" + ex.Message;
                MessageBox.Show(this, ex.Message, "安装失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UninstallSelected(object sender, EventArgs e)
        {
            var selected = SelectedInstallations();
            if (selected.Count == 0)
            {
                MessageBox.Show(this, "请至少选择一个 CDR 版本。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (CorelDrawRunning())
            {
                MessageBox.Show(this, "请先关闭所有 CorelDRAW 窗口，再执行卸载。", "无法卸载", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var confirm = MessageBox.Show(
                this,
                "将从以下 CDR 版本删除 AI矢量助手插件：\r\n\r\n"
                + string.Join("\r\n", selected.Select(x => "• " + x.DisplayName))
                + "\r\n\r\n插件配置和历史记录也会被删除，是否继续？",
                "确认卸载",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            try
            {
                foreach (var installation in selected)
                {
                    if (Directory.Exists(installation.AddonDirectory))
                    {
                        Directory.Delete(installation.AddonDirectory, true);
                    }
                }
                _statusLabel.Text = "卸载完成，请重新启动 CorelDRAW。";
                MessageBox.Show(this, "卸载完成。", "卸载成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _statusLabel.Text = "卸载失败：" + ex.Message;
                MessageBox.Show(this, ex.Message, "卸载失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static bool CorelDrawRunning()
        {
            return Process.GetProcessesByName("CorelDRW").Length > 0;
        }

        private static void InstallTo(string addonDirectory, string payloadDirectory)
        {
            Directory.CreateDirectory(addonDirectory);
            var staleSuperSvg = Path.Combine(addonDirectory, "tools", "supersvg");
            if (Directory.Exists(staleSuperSvg)) Directory.Delete(staleSuperSvg, true);

            foreach (var file in Directory.EnumerateFiles(payloadDirectory, "*", SearchOption.AllDirectories))
            {
                var relative = file.Substring(payloadDirectory.Length).TrimStart('\\', '/');
                var destination = Path.Combine(addonDirectory, relative);
                var parent = Path.GetDirectoryName(destination);
                if (!string.IsNullOrWhiteSpace(parent)) Directory.CreateDirectory(parent);
                File.Copy(file, destination, true);
            }
        }

        private static string ExtractPayload()
        {
            var temp = Path.Combine(Path.GetTempPath(), "AIVectorInstaller-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(temp);
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("AIVectorInstaller.Payload.zip"))
            {
                if (stream == null) throw new InvalidOperationException("安装包中缺少插件 payload。");
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Read, false))
                {
                    foreach (var entry in archive.Entries)
                    {
                        var destination = Path.GetFullPath(Path.Combine(temp, entry.FullName));
                        if (!destination.StartsWith(Path.GetFullPath(temp) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("安装包 payload 路径无效。");
                        if (string.IsNullOrEmpty(entry.Name))
                        {
                            Directory.CreateDirectory(destination);
                            continue;
                        }
                        var parent = Path.GetDirectoryName(destination);
                        if (!string.IsNullOrWhiteSpace(parent)) Directory.CreateDirectory(parent);
                        using (var input = entry.Open())
                        using (var output = File.Create(destination))
                        {
                            input.CopyTo(output);
                        }
                    }
                }
            }
            return temp;
        }

        private void ShowHelp(object sender, EventArgs e)
        {
            MessageBox.Show(
                this,
                "1. 安装程序会自动扫描已安装的 CorelDRAW 版本。\r\n"
                + "2. 勾选需要安装的版本，点击“安装”。\r\n"
                + "3. 卸载时勾选目标版本，点击“卸载选中”。\r\n"
                + "4. 安装或卸载前请关闭 CorelDRAW。\r\n"
                + "5. 完成后重新启动 CorelDRAW。",
                "安装说明",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
    }

    internal static class CheckBoxExtensions
    {
        private static readonly ToolTip ToolTip = new ToolTip();

        public static void ToolTipText(this CheckBox checkBox, string text)
        {
            ToolTip.SetToolTip(checkBox, text ?? "");
        }
    }
}
