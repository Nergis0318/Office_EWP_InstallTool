using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OfficeInstallTool
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    internal sealed class MainForm : Form
    {
        private const string OfficeDir = @"C:\Office";
        private const string SetupUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe";
        private const string ScrubberUrl = "https://gitlab.com/-/project/11037551/uploads/f49f0d69e0aaf92e740a1f694d0438b9/OfficeScrubber_14.zip";

        private const string ConfigurationXml =
@"<Configuration ID=""5d2abb4b-ab45-4fc7-bc44-31a2957cbcb2"">
  <Add OfficeClientEdition=""64"" Channel=""Current"" MigrateArch=""TRUE"">
    <Product ID=""O365ProPlusEEANoTeamsRetail"">
      <Language ID=""ko-kr"" />
      <Language ID=""MatchPreviousMSI"" />
      <ExcludeApp ID=""Access"" />
      <ExcludeApp ID=""Groove"" />
      <ExcludeApp ID=""Lync"" />
      <ExcludeApp ID=""OneDrive"" />
      <ExcludeApp ID=""OneNote"" />
      <ExcludeApp ID=""Outlook"" />
      <ExcludeApp ID=""OutlookForWindows"" />
      <ExcludeApp ID=""Publisher"" />
    </Product>
  </Add>
  <Updates Enabled=""TRUE"" />
  <RemoveMSI />
  <AppSettings>
    <Setup Name=""Company"" Value=""Nergis"" />
    <User
            Key=""software\microsoft\office\16.0\excel\options""
            Name=""defaultformat""
            Value=""51""
            Type=""REG_DWORD""
            App=""excel16""
            Id=""L_SaveExcelfilesas""
        />
    <User
            Key=""software\microsoft\office\16.0\powerpoint\options""
            Name=""defaultformat""
            Value=""27""
            Type=""REG_DWORD""
            App=""ppt16""
            Id=""L_SavePowerPointfilesas""
        />
    <User
            Key=""software\microsoft\office\16.0\word\options""
            Name=""defaultformat""
            Value=""""
            Type=""REG_SZ""
            App=""word16""
            Id=""L_SaveWordfilesas""
        />
  </AppSettings>
  <Display Level=""Full"" AcceptEULA=""TRUE"" />
</Configuration>
";

        private readonly TextBox logBox;
        private readonly Button openAppsButton;
        private readonly Button scrubberButton;
        private readonly Button prepareButton;
        private readonly Button installButton;
        private readonly Label adminLabel;

        public MainForm()
        {
            Text = "Excel / Word / PowerPoint 설치 도구";
            Width = 760;
            Height = 560;
            MinimumSize = new Size(680, 500);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Malgun Gothic", 9F);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5,
                Padding = new Padding(16)
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var title = new Label
            {
                Text = "Office 설치 준비 및 설치",
                Font = new Font(Font.FontFamily, 15F, FontStyle.Bold),
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 4)
            };
            root.Controls.Add(title, 0, 0);

            var description = new Label
            {
                Text = "이 도구는 기존 Office 정리 후 C:\\Office에 ODT와 Configuration.xml을 준비하고 Excel, Word, PowerPoint를 설치합니다.",
                AutoSize = true,
                MaximumSize = new Size(700, 0),
                Margin = new Padding(0, 0, 0, 12)
            };
            root.Controls.Add(description, 0, 1);

            adminLabel = new Label
            {
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 12)
            };
            root.Controls.Add(adminLabel, 0, 2);

            var mainSplit = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 185,
                Panel1MinSize = 150,
                Panel2MinSize = 150
            };

            var buttonGrid = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2
            };
            buttonGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            buttonGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            buttonGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            buttonGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

            openAppsButton = MakeStepButton("1. 앱 및 기능 열기", "기존 Office가 있으면 Windows 설정에서 먼저 제거합니다.");
            scrubberButton = MakeStepButton("2. OfficeScrubber 실행", "[R] Remove all Licenses 옵션을 선택해 라이선스를 정리합니다.");
            prepareButton = MakeStepButton("3. 설치 파일 준비", "C:\\Office 생성, ODT 다운로드, Configuration.xml 생성을 수행합니다.");
            installButton = MakeStepButton("4. Office 설치 및 정리", "관리자 권한 cmd에서 Office 설치 후 C:\\Office를 삭제합니다.");

            buttonGrid.Controls.Add(openAppsButton, 0, 0);
            buttonGrid.Controls.Add(scrubberButton, 1, 0);
            buttonGrid.Controls.Add(prepareButton, 0, 1);
            buttonGrid.Controls.Add(installButton, 1, 1);

            mainSplit.Panel1.Controls.Add(buttonGrid);

            logBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                BackColor = Color.White,
                Font = new Font("Consolas", 9F)
            };
            mainSplit.Panel2.Controls.Add(logBox);
            root.Controls.Add(mainSplit, 0, 3);

            var footer = new Label
            {
                Text = "참고: OfficeScrubber는 외부 도구이며, 실행 후 표시되는 명령창에서 직접 [R] 옵션을 선택해야 합니다.",
                AutoSize = true,
                MaximumSize = new Size(700, 0),
                ForeColor = Color.DimGray,
                Margin = new Padding(0, 10, 0, 0)
            };
            root.Controls.Add(footer, 0, 4);

            Controls.Add(root);

            openAppsButton.Click += (sender, args) => OpenAppsAndFeatures();
            scrubberButton.Click += async (sender, args) => await RunStepAsync(scrubberButton, DownloadAndRunScrubberAsync);
            prepareButton.Click += async (sender, args) => await RunStepAsync(prepareButton, PrepareOfficeFilesAsync);
            installButton.Click += (sender, args) => StartOfficeInstall();

            RefreshAdminStatus();
            Log("도구가 준비되었습니다.");
        }

        private Button MakeStepButton(string title, string body)
        {
            return new Button
            {
                Text = title + Environment.NewLine + body,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(12),
                Margin = new Padding(6),
                Height = 80
            };
        }

        private void RefreshAdminStatus()
        {
            bool admin = IsAdministrator();
            adminLabel.Text = admin
                ? "관리자 권한으로 실행 중입니다."
                : "관리자 권한이 아닙니다. C:\\Office 생성, scrubber 실행, Office 설치는 관리자 권한이 필요할 수 있습니다.";
            adminLabel.ForeColor = admin ? Color.DarkGreen : Color.DarkRed;
        }

        private static bool IsAdministrator()
        {
            using (var identity = WindowsIdentity.GetCurrent())
            {
                var principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        private void OpenAppsAndFeatures()
        {
            try
            {
                Process.Start(new ProcessStartInfo("ms-settings:appsfeatures") { UseShellExecute = true });
                Log("Windows 설정의 앱 및 기능 화면을 열었습니다. 기존 Office가 있으면 제거하세요.");
            }
            catch (Exception ex)
            {
                Log("앱 및 기능 화면을 열지 못했습니다: " + ex.Message);
            }
        }

        private async Task RunStepAsync(Button button, Func<Task> action)
        {
            SetButtonsEnabled(false);
            button.Enabled = false;
            try
            {
                await action();
            }
            catch (Exception ex)
            {
                Log("오류: " + ex.Message);
                MessageBox.Show(this, ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetButtonsEnabled(true);
            }
        }

        private void SetButtonsEnabled(bool enabled)
        {
            openAppsButton.Enabled = enabled;
            scrubberButton.Enabled = enabled;
            prepareButton.Enabled = enabled;
            installButton.Enabled = enabled;
        }

        private async Task DownloadAndRunScrubberAsync()
        {
            EnsureOfficeDirectory();
            string zipPath = Path.Combine(OfficeDir, "OfficeScrubber_14.zip");
            string extractDir = Path.Combine(OfficeDir, "OfficeScrubber");

            if (!File.Exists(zipPath))
            {
                Log("OfficeScrubber 다운로드 중...");
                await DownloadFileAsync(ScrubberUrl, zipPath);
                Log("다운로드 완료: " + zipPath);
            }
            else
            {
                Log("기존 OfficeScrubber zip을 사용합니다: " + zipPath);
            }

            if (Directory.Exists(extractDir))
            {
                Directory.Delete(extractDir, true);
            }

            Log("OfficeScrubber 압축 해제 중...");
            ZipFile.ExtractToDirectory(zipPath, extractDir);

            string cmdFile = FindFile(extractDir, "OfficeScrubber.cmd");
            if (cmdFile == null)
            {
                throw new FileNotFoundException("압축 파일 안에서 OfficeScrubber.cmd를 찾지 못했습니다.");
            }

            Log("OfficeScrubber.cmd 실행: " + cmdFile);
            Log("명령창이 열리면 [R] Remove all Licenses 옵션을 선택하세요.");
            Process.Start(new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = "/k \"" + cmdFile + "\"",
                WorkingDirectory = Path.GetDirectoryName(cmdFile),
                UseShellExecute = true,
                Verb = "runas"
            });
        }

        private async Task PrepareOfficeFilesAsync()
        {
            EnsureOfficeDirectory();
            string setupPath = Path.Combine(OfficeDir, "setup.exe");
            string configPath = Path.Combine(OfficeDir, "Configuration.xml");

            if (!File.Exists(setupPath))
            {
                Log("Office Deployment Tool 다운로드 중...");
                await DownloadFileAsync(SetupUrl, setupPath);
                Log("다운로드 완료: " + setupPath);
            }
            else
            {
                Log("기존 setup.exe를 사용합니다: " + setupPath);
            }

            File.WriteAllText(configPath, ConfigurationXml, new UTF8Encoding(false));
            Log("Configuration.xml 생성 완료: " + configPath);
            Log("설치 준비가 완료되었습니다.");
        }

        private void StartOfficeInstall()
        {
            string setupPath = Path.Combine(OfficeDir, "setup.exe");
            string configPath = Path.Combine(OfficeDir, "Configuration.xml");

            if (!File.Exists(setupPath) || !File.Exists(configPath))
            {
                MessageBox.Show(this, "먼저 [3. 설치 파일 준비]를 실행하세요.", "준비 필요", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                Log("Office 설치 및 C:\\Office 정리 명령 실행 중...");
                Process.Start(new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = "/k cd /d " + Quote(OfficeDir) + " && setup.exe /configure Configuration.xml && cd /d C:\\ && rmdir /s /q " + Quote(OfficeDir) + " && echo C:\\Office 삭제 완료",
                    WorkingDirectory = OfficeDir,
                    UseShellExecute = true,
                    Verb = "runas"
                });
                Log("관리자 권한 명령 프롬프트에서 Office 설치를 시작했습니다. 설치가 정상 종료되면 C:\\Office가 삭제됩니다.");
            }
            catch (Exception ex)
            {
                Log("설치를 시작하지 못했습니다: " + ex.Message);
                MessageBox.Show(this, ex.Message, "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void EnsureOfficeDirectory()
        {
            Directory.CreateDirectory(OfficeDir);
        }

        private static async Task DownloadFileAsync(string url, string path)
        {
            using (var client = new WebClient())
            {
                await client.DownloadFileTaskAsync(new Uri(url), path);
            }
        }

        private static string FindFile(string root, string fileName)
        {
            foreach (var file in Directory.GetFiles(root, fileName, SearchOption.AllDirectories))
            {
                return file;
            }

            return null;
        }

        private static string Quote(string value)
        {
            return "\"" + value + "\"";
        }

        private void Log(string message)
        {
            string line = "[" + DateTime.Now.ToString("HH:mm:ss") + "] " + message + Environment.NewLine;
            logBox.AppendText(line);
        }
    }
}
