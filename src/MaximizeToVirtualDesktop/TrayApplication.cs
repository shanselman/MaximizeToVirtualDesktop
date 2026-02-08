using System.Diagnostics;
using System.Runtime.InteropServices;
using MaximizeToVirtualDesktop.Interop;
using Updatum;

namespace MaximizeToVirtualDesktop;

/// <summary>
/// System tray application. Hosts the NotifyIcon, handles the global hotkey,
/// and owns the lifecycle of all components.
/// </summary>
internal sealed class TrayApplication : Form
{
    private const int HOTKEY_ID = 0x1;
    private const int HOTKEY_PIN_ID = 0x2;
    private uint _shellRestartMessage;
    private bool _comInitialized;

    private readonly NotifyIcon _trayIcon;
    private readonly VirtualDesktopService _vds;
    private readonly FullScreenTracker _tracker;
    private readonly FullScreenManager _manager;
    private readonly WindowMonitor _monitor;
    private readonly MaximizeButtonHook _mouseHook;
    private readonly System.Windows.Forms.Timer _cleanupTimer;
    private System.Windows.Forms.Timer? _retryTimer;

    internal static readonly UpdatumManager Updater = new("shanselman", "MaximizeToVirtualDesktop")
    {
        FetchOnlyLatestRelease = true,
        InstallUpdateSingleFileExecutableName = "MaximizeToVirtualDesktop",
    };

    public TrayApplication()
    {
        // Make the form invisible — we're a tray-only app
        ShowInTaskbar = false;
        WindowState = FormWindowState.Minimized;
        FormBorderStyle = FormBorderStyle.None;
        Opacity = 0;
        Size = new Size(0, 0);

        // Initialize components
        _vds = new VirtualDesktopService();
        _tracker = new FullScreenTracker();
        _manager = new FullScreenManager(_vds, _tracker);
        _monitor = new WindowMonitor(_manager, _tracker, this);
        _mouseHook = new MaximizeButtonHook(_manager, this);

        // System tray icon
        _trayIcon = new NotifyIcon
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? SystemIcons.Application,
            Text = "Maximize to Virtual Desktop\nCtrl+Alt+Shift+X | Shift+Click | Ctrl+Alt+Shift+P to pin",
            Visible = true,
            ContextMenuStrip = BuildContextMenu()
        };

        // Periodic cleanup of stale entries (every 30 seconds)
        _cleanupTimer = new System.Windows.Forms.Timer { Interval = 30_000 };
        _cleanupTimer.Tick += (_, _) => _manager.CleanupStaleEntries();
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);

        // Check Windows version before attempting COM init
        var buildNumber = GetWindowsBuildNumber();
        Trace.WriteLine($"TrayApplication: Windows build {buildNumber}");

        if (buildNumber < 22000)
        {
            // Not Windows 11 at all
            MessageBox.Show(
                "MaximizeToVirtualDesktop requires Windows 11.\n\n" +
                $"Your system is running Windows build {buildNumber}.\n" +
                "Virtual Desktop APIs needed by this app are not available on Windows 10.",
                "MaximizeToVirtualDesktop — Unsupported Windows Version",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            Application.Exit();
            return;
        }

        // Initialize COM (adapter auto-selects based on build number)
        _comInitialized = _vds.Initialize(buildNumber);
        if (!_comInitialized)
        {
            Trace.WriteLine("TrayApplication: COM initialization failed — entering degraded mode.");
            _trayIcon.Text = "Maximize to Virtual Desktop\n⚠️ COM failed — checking for updates...";
            _trayIcon.BalloonTipTitle = "Maximize to Virtual Desktop";
            _trayIcon.BalloonTipText =
                "Virtual Desktop COM interface failed to initialize.\n" +
                "This usually means Windows updated and broke the internal APIs.\n" +
                "Checking for an updated version now...";
            _trayIcon.BalloonTipIcon = ToolTipIcon.Warning;
            _trayIcon.ShowBalloonTip(5000);

            // Immediately check for updates, then retry every 5 minutes
            _ = CheckForUpdatesAsync(userInitiated: false, comFailure: true);
            _retryTimer = new System.Windows.Forms.Timer { Interval = 5 * 60 * 1000 };
            _retryTimer.Tick += async (_, _) =>
            {
                // Try reinitializing COM in case an in-place Windows update fixed it
                if (_vds.Reinitialize())
                {
                    Trace.WriteLine("TrayApplication: COM reinitialized successfully!");
                    _comInitialized = true;
                    _retryTimer!.Stop();
                    _retryTimer.Dispose();
                    _retryTimer = null;
                    _trayIcon.Text = "Maximize to Virtual Desktop\nCtrl+Alt+Shift+X | Shift+Click | Ctrl+Alt+Shift+P to pin";
                    StartMonitoring();
                    return;
                }
                await CheckForUpdatesAsync(userInitiated: false, comFailure: true);
            };
            _retryTimer.Start();

            // Register for Explorer restart — COM might work after Explorer restarts
            _shellRestartMessage = NativeMethods.RegisterWindowMessage("TaskbarCreated");
            return;
        }

        // Register for Explorer restart notification
        _shellRestartMessage = NativeMethods.RegisterWindowMessage("TaskbarCreated");

        // Recover orphaned desktops from a previous crash
        RecoverOrphanedDesktops();

        // Start monitoring
        StartMonitoring();

        Trace.WriteLine("TrayApplication: Started.");

        // Show first-run balloon tip
        ShowFirstRunBalloon();

        // Check for updates asynchronously
        _ = CheckForUpdatesAsync();
    }

    private void StartMonitoring()
    {
        _monitor.Start();
        _mouseHook.Install();
        _cleanupTimer.Start();

        // Register hotkey if not already registered
        if (!NativeMethods.RegisterHotKey(Handle, HOTKEY_ID,
            NativeMethods.MOD_CONTROL | NativeMethods.MOD_ALT | NativeMethods.MOD_SHIFT | NativeMethods.MOD_NOREPEAT,
            NativeMethods.VK_X))
        {
            Trace.WriteLine("TrayApplication: Failed to register hotkey (may already be registered).");
        }
        else
        {
            Trace.WriteLine("TrayApplication: Registered hotkey Ctrl+Alt+Shift+X");
        }

        if (!NativeMethods.RegisterHotKey(Handle, HOTKEY_PIN_ID,
            NativeMethods.MOD_CONTROL | NativeMethods.MOD_ALT | NativeMethods.MOD_SHIFT | NativeMethods.MOD_NOREPEAT,
            NativeMethods.VK_P))
        {
            Trace.WriteLine("TrayApplication: Failed to register pin hotkey.");
        }
        else
        {
            Trace.WriteLine("TrayApplication: Registered hotkey Ctrl+Alt+Shift+P");
        }
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == (int)NativeMethods.WM_HOTKEY && m.WParam == (IntPtr)HOTKEY_ID)
        {
            OnHotkeyPressed();
            return;
        }

        if (m.Msg == (int)NativeMethods.WM_HOTKEY && m.WParam == (IntPtr)HOTKEY_PIN_ID)
        {
            OnPinHotkeyPressed();
            return;
        }

        // Explorer restart: COM objects are now invalid, reinitialize
        if (_shellRestartMessage != 0 && m.Msg == (int)_shellRestartMessage)
        {
            Trace.WriteLine("TrayApplication: Explorer restarted, reinitializing COM...");

            // Windows destroys all virtual desktops on Explorer restart —
            // our tracked COM refs are now stale and must be released.
            _tracker.ClearAll();

            if (_vds.Reinitialize() && !_comInitialized)
            {
                // Recovered from degraded mode!
                Trace.WriteLine("TrayApplication: COM recovered after Explorer restart!");
                _comInitialized = true;
                _retryTimer?.Stop();
                _retryTimer?.Dispose();
                _retryTimer = null;
                _trayIcon.Text = "Maximize to Virtual Desktop\nCtrl+Alt+Shift+X | Shift+Click | Ctrl+Alt+Shift+P to pin";
                StartMonitoring();
            }
            return;
        }

        base.WndProc(ref m);
    }

    private void OnHotkeyPressed()
    {
        if (!_comInitialized)
        {
            Trace.WriteLine("TrayApplication: Hotkey pressed but COM not initialized.");
            return;
        }

        var hwnd = NativeMethods.GetForegroundWindow();
        if (hwnd == IntPtr.Zero || hwnd == Handle)
        {
            Trace.WriteLine("TrayApplication: Hotkey pressed but no valid foreground window.");
            return;
        }

        Trace.WriteLine($"TrayApplication: Hotkey pressed, toggling window {hwnd}");
        _manager.Toggle(hwnd);
    }

    private void OnPinHotkeyPressed()
    {
        if (!_comInitialized)
        {
            Trace.WriteLine("TrayApplication: Pin hotkey pressed but COM not initialized.");
            return;
        }

        var hwnd = NativeMethods.GetForegroundWindow();
        if (hwnd == IntPtr.Zero || hwnd == Handle)
        {
            Trace.WriteLine("TrayApplication: Pin hotkey pressed but no valid foreground window.");
            return;
        }

        Trace.WriteLine($"TrayApplication: Pin hotkey pressed, toggling pin for window {hwnd}");
        _manager.PinToggle(hwnd);
    }

    private ContextMenuStrip BuildContextMenu()
    {
        var menu = new ContextMenuStrip();

        var statusItem = new ToolStripMenuItem("No windows tracked") { Enabled = false };
        menu.Opening += (_, _) =>
        {
            var count = _tracker.Count;
            statusItem.Text = count == 0
                ? "No windows tracked"
                : $"{count} window(s) on virtual desktops";
        };
        menu.Items.Add(statusItem);
        menu.Items.Add(new ToolStripSeparator());

        var restoreAllItem = new ToolStripMenuItem("Restore All", null, (_, _) =>
        {
            _manager.RestoreAll();
        });
        menu.Items.Add(restoreAllItem);

        menu.Items.Add(new ToolStripSeparator());

        var howToUseItem = new ToolStripMenuItem("How to Use", null, (_, _) =>
        {
            ShowUsageInfo();
        });
        menu.Items.Add(howToUseItem);

        menu.Items.Add(new ToolStripSeparator());

        var updateItem = new ToolStripMenuItem("Check for Updates...", null, async (_, _) =>
        {
            await CheckForUpdatesAsync(userInitiated: true);
        });
        menu.Items.Add(updateItem);

        menu.Items.Add(new ToolStripSeparator());

        var exitItem = new ToolStripMenuItem("Exit", null, (_, _) =>
        {
            Application.Exit();
        });
        menu.Items.Add(exitItem);

        return menu;
    }

    private async Task CheckForUpdatesAsync(bool userInitiated = false, bool comFailure = false)
    {
        try
        {
            if (!userInitiated && !comFailure) await Task.Delay(5000);

            var updateFound = await Updater.CheckForUpdatesAsync();

            if (!updateFound)
            {
                if (userInitiated)
                    MessageBox.Show("You're running the latest version.", "No Updates",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (comFailure)
                    _trayIcon.Text = "Maximize to Virtual Desktop\n⚠️ COM failed — no update available yet";
                return;
            }

            var release = Updater.LatestRelease!;
            var changelog = Updater.GetChangelog(true) ?? "No release notes available.";

            var message = comFailure
                ? $"A fix may be available! Version {release.TagName} is ready.\n\n{changelog}\n\nDownload and install?"
                : $"Version {release.TagName} is available.\n\n{changelog}\n\nDownload and install?";

            var result = MessageBox.Show(message,
                comFailure ? "Update Available — May Fix COM Issue" : "Update Available",
                MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            if (result == DialogResult.Yes)
            {
                var asset = await Updater.DownloadUpdateAsync();
                if (asset != null)
                {
                    await Updater.InstallUpdateAsync(asset);
                }
                else if (userInitiated)
                {
                    MessageBox.Show("Failed to download the update.", "Update Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"TrayApplication: Update check failed: {ex.Message}");
            if (userInitiated)
                MessageBox.Show($"Update check failed: {ex.Message}", "Update Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static readonly string FirstRunMarker = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "MaximizeToVirtualDesktop", ".firstrun");

    private void ShowFirstRunBalloon()
    {
        try
        {
            if (File.Exists(FirstRunMarker)) return;

            Directory.CreateDirectory(Path.GetDirectoryName(FirstRunMarker)!);
            File.WriteAllText(FirstRunMarker, "");

            _trayIcon.BalloonTipTitle = "Maximize to Virtual Desktop";
            _trayIcon.BalloonTipText =
                "Press Ctrl+Alt+Shift+X or Shift+Click the maximize button " +
                "to maximize a window to its own virtual desktop.\n" +
                "Press Ctrl+Alt+Shift+P to pin a window to all desktops.";
            _trayIcon.BalloonTipIcon = ToolTipIcon.Info;
            _trayIcon.ShowBalloonTip(5000);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"TrayApplication: First-run balloon failed: {ex.Message}");
        }
    }

    private static void ShowUsageInfo()
    {
        using var form = new Form
        {
            Text = "How to Use — Maximize to Virtual Desktop",
            Size = new Size(620, 580),
            StartPosition = FormStartPosition.CenterScreen,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            ShowInTaskbar = false,
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath),
        };

        var rtb = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            BorderStyle = BorderStyle.None,
            BackColor = SystemColors.Window,
        };

        // Add inner padding via a wrapper panel
        var contentPanel = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(20, 16, 20, 0),
        };
        contentPanel.Controls.Add(rtb);

        // Build RTF content
        rtb.SelectionFont = new Font("Segoe UI Variable Display", 14f, FontStyle.Bold);
        rtb.AppendText("Maximize to Virtual Desktop\n\n");

        AppendSection(rtb, "Maximize a Window to Its Own Desktop",
            ("Ctrl + Alt + Shift + X", "Toggles the focused window to/from its own virtual desktop."),
            ("Shift + Click maximize button", "Hold Shift and click any window's maximize button."));

        AppendSection(rtb, "Pin a Window to All Desktops",
            ("Ctrl + Alt + Shift + P", "Toggles pin/unpin so the focused window appears on every desktop."));

        AppendSection(rtb, "How It Works",
            ("Maximize", "The window moves to a new virtual desktop and is maximized full-screen."),
            ("Restore", "Close or restore the window to automatically return to your original desktop."),
            ("Restore All", "Use the tray menu to bring all windows back at once."));

        var panel = new Panel { Dock = DockStyle.Bottom, Height = 50 };
        var okButton = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Size = new Size(90, 32),
            FlatStyle = FlatStyle.System,
        };
        okButton.Location = new Point((panel.Width - okButton.Width) / 2, 9);
        panel.Resize += (_, _) => okButton.Location = new Point((panel.Width - okButton.Width) / 2, 9);
        panel.Controls.Add(okButton);

        form.AcceptButton = okButton;
        form.Controls.Add(contentPanel);
        form.Controls.Add(panel);
        form.ShowDialog();
    }

    private static void AppendSection(RichTextBox rtb, string heading, params (string key, string desc)[] items)
    {
        rtb.SelectionFont = new Font("Segoe UI Variable Display", 10.5f, FontStyle.Bold);
        rtb.AppendText(heading + "\n");

        foreach (var (key, desc) in items)
        {
            rtb.SelectionFont = new Font("Segoe UI Variable Display", 9.5f, FontStyle.Regular);
            rtb.AppendText("  ");
            rtb.SelectionFont = new Font("Consolas", 9f, FontStyle.Bold);
            rtb.AppendText(key);
            rtb.SelectionFont = new Font("Segoe UI Variable Display", 9.5f, FontStyle.Regular);
            rtb.AppendText("  —  " + desc + "\n");
        }

        rtb.AppendText("\n");
    }

    private static int GetWindowsBuildNumber()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion");
            var build = key?.GetValue("CurrentBuildNumber")?.ToString();
            return int.TryParse(build, out var num) ? num : 0;
        }
        catch
        {
            return 0;
        }
    }

    private void RecoverOrphanedDesktops()
    {
        var persisted = TrackerPersistence.Load();
        if (persisted.Count == 0) return;

        Trace.WriteLine($"TrayApplication: Found {persisted.Count} orphaned desktop(s) from previous session.");

        foreach (var entry in persisted)
        {
            var desktop = _vds.FindDesktop(entry.TempDesktopId);
            if (desktop != null)
            {
                Trace.WriteLine($"TrayApplication: Removing orphaned desktop {entry.TempDesktopId} ({entry.ProcessName ?? "unknown"})");
                _vds.RemoveDesktop(desktop);
                Marshal.ReleaseComObject(desktop);
            }
        }

        TrackerPersistence.Delete();
        Trace.WriteLine("TrayApplication: Orphaned desktop recovery complete.");
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        Trace.WriteLine("TrayApplication: Shutting down...");

        _retryTimer?.Stop();
        _retryTimer?.Dispose();

        _cleanupTimer.Stop();
        _cleanupTimer.Dispose();

        // Restore all tracked windows before exiting
        _manager.RestoreAll();

        // Clean up native resources
        NativeMethods.UnregisterHotKey(Handle, HOTKEY_ID);
        NativeMethods.UnregisterHotKey(Handle, HOTKEY_PIN_ID);
        _mouseHook.Dispose();
        _monitor.Dispose();
        _vds.Dispose();

        _trayIcon.Visible = false;
        _trayIcon.Dispose();

        Trace.WriteLine("TrayApplication: Shutdown complete.");
        base.OnFormClosing(e);
    }
}
