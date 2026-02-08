using System.Diagnostics;
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
    private uint _shellRestartMessage;

    private readonly NotifyIcon _trayIcon;
    private readonly VirtualDesktopService _vds;
    private readonly FullScreenTracker _tracker;
    private readonly FullScreenManager _manager;
    private readonly WindowMonitor _monitor;
    private readonly MaximizeButtonHook _mouseHook;
    private readonly System.Windows.Forms.Timer _cleanupTimer;

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
            Text = "Maximize to Virtual Desktop\nCtrl+Alt+Shift+X or Shift+Click maximize button",
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

        // Initialize COM
        if (!_vds.Initialize())
        {
            MessageBox.Show(
                "Failed to initialize Virtual Desktop COM interface.\n\n" +
                "This may mean your Windows version is not supported.\n" +
                "MaximizeToVirtualDesktop requires Windows 11 24H2 or later.",
                "MaximizeToVirtualDesktop",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            Application.Exit();
            return;
        }

        // Register global hotkey: Ctrl+Alt+Shift+X
        if (!NativeMethods.RegisterHotKey(Handle, HOTKEY_ID,
            NativeMethods.MOD_CONTROL | NativeMethods.MOD_ALT | NativeMethods.MOD_SHIFT | NativeMethods.MOD_NOREPEAT,
            NativeMethods.VK_X))
        {
            MessageBox.Show(
                "Failed to register hotkey Ctrl+Alt+Shift+X.\n\n" +
                "Another application may already be using this key combination.",
                "MaximizeToVirtualDesktop",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            // Continue anyway — the app is still useful if we add Shift+Click later
        }
        else
        {
            Trace.WriteLine("TrayApplication: Registered hotkey Ctrl+Alt+Shift+X");
        }

        // Register for Explorer restart notification
        _shellRestartMessage = NativeMethods.RegisterWindowMessage("TaskbarCreated");

        // Start monitoring
        _monitor.Start();
        _mouseHook.Install();
        _cleanupTimer.Start();

        Trace.WriteLine("TrayApplication: Started.");

        // Show first-run balloon tip
        ShowFirstRunBalloon();

        // Check for updates asynchronously
        _ = CheckForUpdatesAsync();
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == (int)NativeMethods.WM_HOTKEY && m.WParam == (IntPtr)HOTKEY_ID)
        {
            OnHotkeyPressed();
            return;
        }

        // Explorer restart: COM objects are now invalid, reinitialize
        if (_shellRestartMessage != 0 && m.Msg == (int)_shellRestartMessage)
        {
            Trace.WriteLine("TrayApplication: Explorer restarted, reinitializing COM...");
            _vds.Reinitialize();
            return;
        }

        base.WndProc(ref m);
    }

    private void OnHotkeyPressed()
    {
        var hwnd = NativeMethods.GetForegroundWindow();
        if (hwnd == IntPtr.Zero || hwnd == Handle)
        {
            Trace.WriteLine("TrayApplication: Hotkey pressed but no valid foreground window.");
            return;
        }

        Trace.WriteLine($"TrayApplication: Hotkey pressed, toggling window {hwnd}");
        _manager.Toggle(hwnd);
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

    private async Task CheckForUpdatesAsync(bool userInitiated = false)
    {
        try
        {
            if (!userInitiated) await Task.Delay(3000);

            var updateFound = await Updater.CheckForUpdatesAsync();

            if (!updateFound)
            {
                if (userInitiated)
                    MessageBox.Show("You're running the latest version.", "No Updates",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var release = Updater.LatestRelease!;
            var changelog = Updater.GetChangelog(true) ?? "No release notes available.";
            var result = MessageBox.Show(
                $"Version {release.TagName} is available.\n\n{changelog}\n\nDownload and install?",
                "Update Available",
                MessageBoxButtons.YesNo, MessageBoxIcon.Information);

            if (result == DialogResult.Yes)
            {
                var asset = await Updater.DownloadUpdateAsync();
                if (asset != null)
                {
                    await Updater.InstallUpdateAsync(asset);
                }
                else
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
                "to maximize a window to its own virtual desktop.";
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
        MessageBox.Show(
            "Maximize to Virtual Desktop\n\n" +
            "Two ways to maximize a window to its own virtual desktop:\n\n" +
            "  • Hotkey: Ctrl+Alt+Shift+X\n" +
            "    Toggles the active window to/from a virtual desktop.\n\n" +
            "  • Shift+Click the maximize button\n" +
            "    Hold Shift and click any window's maximize button.\n\n" +
            "The window is moved to a new virtual desktop and maximized.\n" +
            "Close or restore the window to return to your original desktop.\n\n" +
            "Use \"Restore All\" in the tray menu to bring everything back.",
            "How to Use — Maximize to Virtual Desktop",
            MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        Trace.WriteLine("TrayApplication: Shutting down...");

        _cleanupTimer.Stop();
        _cleanupTimer.Dispose();

        // Restore all tracked windows before exiting
        _manager.RestoreAll();

        // Clean up native resources
        NativeMethods.UnregisterHotKey(Handle, HOTKEY_ID);
        _mouseHook.Dispose();
        _monitor.Dispose();
        _vds.Dispose();

        _trayIcon.Visible = false;
        _trayIcon.Dispose();

        Trace.WriteLine("TrayApplication: Shutdown complete.");
        base.OnFormClosing(e);
    }
}
