using System;
using System.IO;
using System.Net;
using System.Reflection;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32;
using System.Threading.Tasks;
using System.Text;

namespace EZBlocker
{
    public partial class Main : Form
    {
        private bool _muted = false;
        private string _lastMessage = string.Empty;
        private readonly ToolTip _artistTooltip = new ToolTip();
        private readonly string _volumeMixerPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "SndVol.exe");
        private readonly string _hostsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), @"drivers\etc\hosts");

        private static readonly string[] AdHosts = 
        { 
            "pubads.g.doubleclick.net", 
            "securepubads.g.doubleclick.net", 
            "www.googletagservices.com", 
            "gads.pubmatic.com", 
            "ads.pubmatic.com", 
            "tpc.googlesyndication.com", 
            "pagead2.googlesyndication.com", 
            "googleads.g.doubleclick.net" 
        };

        public const string WebsiteUrl = @"https://www.ericzhang.me/projects/spotify-ad-blocker-ezblocker/";

        private Analytics _analytics;
        private DateTime _lastRequest;
        private string _lastAction = string.Empty;
        private MediaHook _hook;

        private const int DefaultMainTimerInterval = 1000;
        private const int PlayingMainTimerInterval = 200;

        public Main()
        {
            Thread.CurrentThread.CurrentUICulture = Thread.CurrentThread.CurrentCulture;
            InitializeComponent();
        }

        private async void MainTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (!_hook.IsRunning()) 
                {
                    UpdateStatus(Properties.strings.StatusNotFound, string.Empty);
                    SetMainTimerInterval(DefaultMainTimerInterval);
                    return;
                }

                Debug.WriteLine(_hook.IsAdPlaying ? "Ad is playing" : "Normal music is playing");

                if (_hook.IsAdPlaying)
                {
                    HandleAdPlaying();
                }
                else if (_hook.IsPlaying)
                {
                    HandleNormalPlaying();
                }
                else
                {
                    UpdateStatus(Properties.strings.StatusPaused, string.Empty);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void HandleAdPlaying()
        {
            SetMainTimerInterval(DefaultMainTimerInterval);

            if (!_muted) Mute(true);

            if (!_hook.IsPlaying)
            {
                _hook.SendNextTrack();
                Thread.Sleep(500);
            }

            string artist = _hook.CurrentArtist;
            string message = $"{Properties.strings.StatusMuting} {Truncate(artist)}";
            UpdateStatus(message, artist);

            LogAction($"/mute/{artist}");
        }

        private void HandleNormalPlaying()
        {
            if (_muted)
            {
                Thread.Sleep(200);
                Mute(false);
            }

            SetMainTimerInterval(PlayingMainTimerInterval);

            string artist = _hook.CurrentArtist;
            string message = $"{Properties.strings.StatusPlaying} {Truncate(artist)}";
            UpdateStatus(message, artist);

            LogAction($"/play/{artist}");
        }

        private void SetMainTimerInterval(int interval)
        {
            if (MainTimer.Interval != interval)
                MainTimer.Interval = interval;
        }

        private void Mute(bool mute)
        {
            AudioUtils.SetSpotifyMute(mute);
            _muted = mute;
        }

        private string Truncate(string name)
        {
            return name.Length > 9 ? name.Substring(0, 9) + "..." : name;
        }

        private async Task CheckUpdate()
        {
            try
            {
                using (var webClient = new WebClient())
                {
                    webClient.Headers.Add("user-agent", $"EZBlocker {Assembly.GetExecutingAssembly().GetName().Version} {Environment.OSVersion}");

                    string versionString = await webClient.DownloadStringTaskAsync("https://www.ericzhang.me/dl/?file=EZBlocker-version.txt");
                    int latestVersion = Convert.ToInt32(versionString);
                    int currentVersion = Convert.ToInt32(Assembly.GetExecutingAssembly().GetName().Version.ToString().Replace(".", ""));

                    if (latestVersion > currentVersion && 
                        MessageBox.Show(Properties.strings.UpgradeMessageBox, "EZBlocker", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        Process.Start(WebsiteUrl);
                        Application.Exit();
                    }
                }
            }
            catch (WebException)
            {
                MessageBox.Show(Properties.strings.UpgradeErrorMessageBox, "EZBlocker");
            }
        }

        private void LogAction(string action)
        {
            if (_lastAction.Equals(action) && DateTime.Now - _lastRequest < TimeSpan.FromMinutes(5)) 
                return;

            Task.Run(() => _analytics.LogAction(action));
            _lastAction = action;
            _lastRequest = DateTime.Now;
        }

        private void Heartbeat_Tick(object sender, EventArgs e)
        {
            if (DateTime.Now - _lastRequest > TimeSpan.FromMinutes(5))
            {
                LogAction("/heartbeat");
            }
        }

        private void Main_Load(object sender, EventArgs e)
        {
            PerformInitialSetup();

            // Set up UI and Spotify Hook
            SetupUI();
            _hook = new MediaHook();
            MainTimer.Enabled = true;

            LogAction("/launch");
            Task.Run(() => CheckUpdate());
        }

        private void PerformInitialSetup()
        {
            if (Properties.Settings.Default.UpdateSettings)
            {
                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.UpdateSettings = false;
                Properties.Settings.Default.Save();
            }

            string spotifyPath = GetSpotifyPath();
            Properties.Settings.Default.SpotifyPath = spotifyPath;
            Properties.Settings.Default.Save();

            StartSpotifyIfNeeded();
        }

        private void StartSpotifyIfNeeded()
        {
            try
            {
                Process.GetCurrentProcess().PriorityClass = ProcessPriorityClass.High;

                if (Properties.Settings.Default.StartSpotify &&
                    File.Exists(Properties.Settings.Default.SpotifyPath) &&
                    Process.GetProcessesByName("spotify").Length < 1)
                {
                    Process.Start(Properties.Settings.Default.SpotifyPath);
                }
            }
            catch (InvalidOperationException) { }
        }

        private void SetupUI()
        {
            if (File.Exists(_hostsPath))
            {
                string hostsFile = File.ReadAllText(_hostsPath);
                BlockBannersCheckbox.Checked = AdHosts.All(host => hostsFile.Contains("0.0.0.0 " + host));
            }

            SetupStartupCheckbox();
            SpotifyCheckbox.Checked = Properties.Settings.Default.StartSpotify;

            SetupAnalytics();
        }

        private void SetupStartupCheckbox()
        {
            using (RegistryKey startupKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                string exePath = $"\"{Application.ExecutablePath}\"";
                string startupValue = startupKey.GetValue("EZBlocker")?.ToString();

                if (startupValue != null)
                {
                    StartupCheckbox.Checked = startupValue == exePath;
                    this.WindowState = FormWindowState.Minimized;
                }
                else if (startupValue != exePath)
                {
                    startupKey.DeleteValue("EZBlocker");
                }
            }
        }

        private void SetupAnalytics()
        {
            if (string.IsNullOrEmpty(Properties.Settings.Default.CID))
            {
                Properties.Settings.Default.CID = Analytics.GenerateCID();
                Properties.Settings.Default.Save();
            }
            _analytics = new Analytics(Properties.Settings.Default.CID, Assembly.GetExecutingAssembly().GetName().Version.ToString());
        }

        private string GetSpotifyPath()
        {
            var process = Process.GetProcessesByName("spotify").FirstOrDefault(p => p.MainWindowTitle.Length > 1);
            return process?.MainModule.FileName ?? string.Empty;
        }

        private void RestoreFromTray()
        {
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
        }

        private void Notify(string message)
        {
            NotifyIcon.ShowBalloonTip(5000, "EZBlocker", message, ToolTipIcon.None);
        }

        private void NotifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (!this.ShowInTaskbar && e.Button == MouseButtons.Left)
            {
                RestoreFromTray();
            }
        }

        private void NotifyIcon_BalloonTipClicked(object sender, EventArgs e)
        {
            RestoreFromTray();
        }

        private void Form_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                this.ShowInTaskbar = false;
                this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
                Notify(Properties.strings.HiddenNotify);
            }
        }

        private void SkipAdsCheckbox_Click(object sender, EventArgs e)
        {
            if (!MainTimer.Enabled) return;

            if (!IsUserAnAdmin())
            {
                MessageBox.Show(Properties.strings.BlockBannersUAC, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                BlockBannersCheckbox.Checked = !BlockBannersCheckbox.Checked;
                return;
            }

            try
            {
                UpdateHostsFile();
                MessageBox.Show(Properties.strings.BlockBannersRestart, "EZBlocker", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LogAction($"/settings/blockBanners/{BlockBannersCheckbox.Checked}");
            }
            catch (IOException ex)
            {
                Debug.WriteLine(ex);
            }
        }

        private void UpdateHostsFile()
        {
            if (!File.Exists(_hostsPath))
            {
                File.Create(_hostsPath).Close();
            }

            var hostsFileLines = File.ReadAllLines(_hostsPath)
                                     .Where(line => !AdHosts.Contains(line.Replace("0.0.0.0 ", "")) && line.Length > 0)
                                     .ToArray();

            File.WriteAllLines(_hostsPath, hostsFileLines);

            if (BlockBannersCheckbox.Checked)
            {
                using (var sw = File.AppendText(_hostsPath))
                {
                    sw.WriteLine();
                    foreach (string host in AdHosts)
                    {
                        sw.WriteLine("0.0.0.0 " + host);
                    }
                }
            }
        }

        private void StartupCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            if (!MainTimer.Enabled) return;

            using (RegistryKey startupKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true))
            {
                if (StartupCheckbox.Checked)
                {
                    startupKey.SetValue("EZBlocker", $"\"{Application.ExecutablePath}\"");
                }
                else
                {
                    startupKey.DeleteValue("EZBlocker");
                }
            }

            LogAction($"/settings/startup/{StartupCheckbox.Checked}");
        }

        private void SpotifyCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            if (!MainTimer.Enabled) return;

            Properties.Settings.Default.StartSpotify = SpotifyCheckbox.Checked;
            Properties.Settings.Default.Save();
            LogAction($"/settings/startSpotify/{SpotifyCheckbox.Checked}");
        }

        private void VolumeMixerButton_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(_volumeMixerPath);
                LogAction("/button/volumeMixer");
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show(Properties.strings.VolumeMixerOpenError, "EZBlocker");
            }
        }

        private void WebsiteLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (File.Exists(Properties.Settings.Default.SpotifyPath))
            {
                string message = Properties.strings.ReportProblemMessageBox
                                 .Replace("{0}", Assembly.GetExecutingAssembly().GetName().Version.ToString())
                                 .Replace("{1}", FileVersionInfo.GetVersionInfo(Properties.Settings.Default.SpotifyPath).FileVersion);
                MessageBox.Show(message, "EZBlocker");

                Clipboard.SetText(Properties.strings.ReportProblemClipboard
                                  .Replace("{0}", Assembly.GetExecutingAssembly().GetName().Version.ToString())
                                  .Replace("{1}", FileVersionInfo.GetVersionInfo(Properties.Settings.Default.SpotifyPath).FileVersion));
            }

            Process.Start(WebsiteUrl);
            LogAction("/button/website");
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RestoreFromTray();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void websiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start(WebsiteUrl);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!MainTimer.Enabled) return;

            if (!Properties.Settings.Default.UserEducated)
            {
                var result = MessageBox.Show(Properties.strings.OnExitMessageBox, "EZBlocker", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                e.Cancel = (result == DialogResult.No);

                if (result == DialogResult.Yes)
                {
                    Properties.Settings.Default.UserEducated = true;
                    Properties.Settings.Default.Save();
                }
            }
        }

        private void UpdateStatus(string statusMessage, string tooltipMessage)
        {
            if (_lastMessage == statusMessage) return;

            _lastMessage = statusMessage;
            StatusLabel.Text = statusMessage;
            _artistTooltip.SetToolTip(StatusLabel, tooltipMessage);
        }

        [DllImport("shell32.dll")]
        public static extern bool IsUserAnAdmin();
    }
}
