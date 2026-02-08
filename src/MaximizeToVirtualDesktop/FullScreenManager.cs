using System.Diagnostics;
using System.Runtime.InteropServices;
using MaximizeToVirtualDesktop.Interop;

namespace MaximizeToVirtualDesktop;

/// <summary>
/// Orchestrates the "maximize to virtual desktop" and "restore from virtual desktop" flows.
/// Every mutating step has rollback if the next step fails.
/// </summary>
internal sealed class FullScreenManager
{
    private readonly VirtualDesktopService _vds;
    private readonly FullScreenTracker _tracker;
    private readonly HashSet<IntPtr> _inFlight = new();

    public FullScreenManager(VirtualDesktopService vds, FullScreenTracker tracker)
    {
        _vds = vds;
        _tracker = tracker;
    }

    /// <summary>
    /// Toggle: if window is tracked, restore it. Otherwise, maximize it to a new desktop.
    /// </summary>
    public void Toggle(IntPtr hwnd)
    {
        if (!NativeMethods.IsWindow(hwnd))
        {
            Trace.WriteLine($"FullScreenManager: hwnd {hwnd} is not a valid window, ignoring.");
            return;
        }

        if (!_inFlight.Add(hwnd))
        {
            Trace.WriteLine($"FullScreenManager: hwnd {hwnd} already in-flight, ignoring.");
            return;
        }

        try
        {
            if (_tracker.IsTracked(hwnd))
            {
                Restore(hwnd);
            }
            else
            {
                MaximizeToDesktop(hwnd);
            }
        }
        finally
        {
            _inFlight.Remove(hwnd);
        }
    }

    /// <summary>
    /// Send a window to a new virtual desktop, maximized.
    /// </summary>
    public void MaximizeToDesktop(IntPtr hwnd)
    {
        if (!NativeMethods.IsWindow(hwnd))
        {
            Trace.WriteLine($"FullScreenManager: hwnd {hwnd} is not valid, aborting maximize.");
            return;
        }

        if (_tracker.IsTracked(hwnd))
        {
            Trace.WriteLine($"FullScreenManager: hwnd {hwnd} already tracked, toggling to restore.");
            Restore(hwnd);
            return;
        }

        // 1. Record original state
        var originalDesktopId = _vds.GetDesktopIdForWindow(hwnd);
        if (originalDesktopId == null)
        {
            Trace.WriteLine("FullScreenManager: Could not determine original desktop, aborting.");
            return;
        }

        var originalPlacement = NativeMethods.WINDOWPLACEMENT.Default;
        if (!NativeMethods.GetWindowPlacement(hwnd, ref originalPlacement))
        {
            Trace.WriteLine("FullScreenManager: Could not get window placement, aborting.");
            return;
        }

        // 1b. Determine target monitor and collect companion windows
        var targetMonitor = GetTargetMonitor(hwnd);
        var companions = GetCompanionWindows(targetMonitor, hwnd);
        Trace.WriteLine($"FullScreenManager: Target monitor={targetMonitor.DeviceName}, {companions.Count} companion window(s).");

        // 2. Create new virtual desktop
        var (tempDesktop, tempDesktopId) = _vds.CreateDesktop();
        if (tempDesktop == null || tempDesktopId == null)
        {
            Trace.WriteLine("FullScreenManager: Failed to create desktop, aborting.");
            return;
        }

        // 3. Name the desktop after the window title (or process name as fallback)
        string? processName = null;
        try
        {
            NativeMethods.GetWindowThreadProcessId(hwnd, out int processId);
            using var process = Process.GetProcessById(processId);
            processName = !string.IsNullOrWhiteSpace(process.MainWindowTitle)
                ? process.MainWindowTitle
                : process.ProcessName;
            _vds.SetDesktopName(tempDesktop, $"[MVD] {processName}");
        }
        catch
        {
            // Non-critical, continue
        }

        // 4. Move window to new desktop
        if (!_vds.MoveWindowToDesktop(hwnd, tempDesktop))
        {
            Trace.WriteLine("FullScreenManager: Failed to move window, rolling back desktop creation.");
            _vds.RemoveDesktop(tempDesktop);
            Marshal.ReleaseComObject(tempDesktop);
            return;
        }

        // 4b. Move companion windows to new desktop
        int movedCompanions = 0;
        foreach (var companion in companions)
        {
            if (_vds.MoveWindowToDesktop(companion, tempDesktop))
                movedCompanions++;
            else
                Trace.WriteLine($"FullScreenManager: Failed to move companion {companion}, skipping.");
        }
        if (companions.Count > 0)
            Trace.WriteLine($"FullScreenManager: Moved {movedCompanions}/{companions.Count} companion(s) to new desktop.");

        // 5. Switch to the new desktop
        if (!_vds.SwitchToDesktop(tempDesktop))
        {
            // Rollback: move window back, remove desktop
            Trace.WriteLine("FullScreenManager: Failed to switch desktop, rolling back.");
            var origDesktop = _vds.FindDesktop(originalDesktopId.Value);
            try
            {
                if (origDesktop != null) _vds.MoveWindowToDesktop(hwnd, origDesktop);
            }
            finally
            {
                if (origDesktop != null) Marshal.ReleaseComObject(origDesktop);
            }
            _vds.RemoveDesktop(tempDesktop);
            Marshal.ReleaseComObject(tempDesktop);
            return;
        }

        // 5b. Move window to target monitor (multi-monitor setups)
        if (Screen.AllScreens.Length > 1)
            MoveWindowToMonitor(hwnd, targetMonitor);

        // 6. Maximize the window — delay lets desktop switch animation finish first
        bool elevated = NativeMethods.IsWindowElevated(hwnd);
        if (elevated)
        {
            Trace.WriteLine("FullScreenManager: Window is elevated, cannot maximize via UIPI.");
            NotificationOverlay.ShowNotification("⚠ Elevated Window",
                "Press Win+↑ to maximize", hwnd);
        }
        else
        {
            Thread.Sleep(250);
            NativeMethods.ShowWindow(hwnd, NativeMethods.SW_MAXIMIZE);
        }
        NativeMethods.SetForegroundWindow(hwnd);

        // 7. Track it
        _tracker.Track(hwnd, originalDesktopId.Value, tempDesktopId.Value, tempDesktop, processName, originalPlacement);

        NotificationOverlay.ShowNotification("→ Virtual Desktop", processName ?? "", hwnd);
        Trace.WriteLine($"FullScreenManager: Successfully maximized {hwnd} to desktop {tempDesktopId}");
    }

    /// <summary>
    /// Restore a tracked window: move it back to its original desktop, restore window state,
    /// switch back, and remove the temp desktop.
    /// </summary>
    public void Restore(IntPtr hwnd)
    {
        var entry = _tracker.Get(hwnd);
        if (entry == null)
        {
            Trace.WriteLine($"FullScreenManager: hwnd {hwnd} not tracked, ignoring restore.");
            return;
        }

        // Untrack first to prevent reentrant calls from WindowMonitor
        _tracker.Untrack(hwnd);

        var windowStillExists = NativeMethods.IsWindow(hwnd);

        // 1. Restore window placement (before moving, so it's sized correctly)
        if (windowStillExists)
        {
            var placement = entry.OriginalPlacement;
            NativeMethods.SetWindowPlacement(hwnd, ref placement);
        }

        // 2. Move window back to original desktop and switch back
        var origDesktop = _vds.FindDesktop(entry.OriginalDesktopId);
        try
        {
            if (origDesktop != null)
            {
                if (windowStillExists) _vds.MoveWindowToDesktop(hwnd, origDesktop);
                _vds.SwitchToDesktop(origDesktop);
            }
            else
            {
                Trace.WriteLine("FullScreenManager: Original desktop no longer exists, leaving window on current.");
            }
        }
        finally
        {
            if (origDesktop != null) Marshal.ReleaseComObject(origDesktop);
        }

        // 3. Remove temp desktop and release its COM reference
        _vds.RemoveDesktop(entry.TempDesktop);
        Marshal.ReleaseComObject(entry.TempDesktop);

        // 4. Set focus on the restored window
        if (windowStillExists)
        {
            NativeMethods.SetForegroundWindow(hwnd);
        }

        NotificationOverlay.ShowNotification("← Restored", entry.ProcessName ?? "", hwnd);
        Trace.WriteLine($"FullScreenManager: Restored {hwnd} to original desktop.");
    }

    /// <summary>
    /// Called when a tracked window is destroyed (closed). Clean up its temp desktop.
    /// </summary>
    public void HandleWindowDestroyed(IntPtr hwnd)
    {
        var entry = _tracker.Untrack(hwnd);
        if (entry == null) return;

        Trace.WriteLine($"FullScreenManager: Tracked window {hwnd} destroyed, cleaning up.");

        // Switch back to original desktop first
        var origDesktop = _vds.FindDesktop(entry.OriginalDesktopId);
        try
        {
            if (origDesktop != null) _vds.SwitchToDesktop(origDesktop);
        }
        finally
        {
            if (origDesktop != null) Marshal.ReleaseComObject(origDesktop);
        }

        // Then remove the temp desktop and release its COM reference
        _vds.RemoveDesktop(entry.TempDesktop);
        Marshal.ReleaseComObject(entry.TempDesktop);
    }

    /// <summary>
    /// Clean up all tracked windows — called on app exit.
    /// </summary>
    public void RestoreAll()
    {
        var entries = _tracker.GetAll();
        Trace.WriteLine($"FullScreenManager: Restoring {entries.Count} tracked window(s) on exit.");

        foreach (var entry in entries)
        {
            try
            {
                Restore(entry.Hwnd);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"FullScreenManager: Error restoring {entry.Hwnd}: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Remove stale entries for windows that no longer exist.
    /// </summary>
    public void CleanupStaleEntries()
    {
        var stale = _tracker.GetStaleHandles();
        foreach (var hwnd in stale)
        {
            HandleWindowDestroyed(hwnd);
        }
    }

    /// <summary>
    /// Toggle pin/unpin of a window to all virtual desktops.
    /// </summary>
    public void PinToggle(IntPtr hwnd)
    {
        if (!NativeMethods.IsWindow(hwnd))
        {
            Trace.WriteLine($"FullScreenManager: hwnd {hwnd} is not valid, ignoring pin toggle.");
            return;
        }

        string? processName = null;
        try
        {
            NativeMethods.GetWindowThreadProcessId(hwnd, out int pid);
            using var process = System.Diagnostics.Process.GetProcessById(pid);
            processName = !string.IsNullOrWhiteSpace(process.MainWindowTitle)
                ? process.MainWindowTitle
                : process.ProcessName;
        }
        catch { }

        if (_vds.IsWindowPinned(hwnd))
        {
            if (_vds.UnpinWindow(hwnd))
                NotificationOverlay.ShowNotification("📌 Unpinned", processName ?? "", hwnd);
            else
                NotificationOverlay.ShowNotification("⚠ Unpin Failed", processName ?? "", hwnd);
        }
        else
        {
            if (_vds.PinWindow(hwnd))
                NotificationOverlay.ShowNotification("📌 Pinned to All Desktops", processName ?? "", hwnd);
            else
                NotificationOverlay.ShowNotification("⚠ Pin Failed", processName ?? "", hwnd);
        }
    }

    // ---- Multi-monitor helpers ----

    private static readonly HashSet<string> _shellClassNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "Shell_TrayWnd", "Shell_SecondaryTrayWnd", "Progman", "WorkerW"
    };

    /// <summary>
    /// Determines the target monitor for maximization.
    /// Default: the horizontally-middle monitor. Single-monitor: current monitor.
    /// </summary>
    private static Screen GetTargetMonitor(IntPtr hwnd)
    {
        var screens = Screen.AllScreens.OrderBy(s => s.Bounds.X).ToArray();
        if (screens.Length < 2)
            return Screen.FromHandle(hwnd);
        return screens[screens.Length / 2];
    }

    /// <summary>
    /// Enumerates visible top-level windows on monitors OTHER than the target.
    /// These "companion" windows will be brought along to the new virtual desktop.
    /// </summary>
    private List<IntPtr> GetCompanionWindows(Screen targetMonitor, IntPtr primaryHwnd)
    {
        var companions = new List<IntPtr>();
        var classNameBuf = new char[256];

        NativeMethods.EnumWindows((hwnd, _) =>
        {
            if (hwnd == primaryHwnd) return true;
            if (!NativeMethods.IsWindowVisible(hwnd)) return true;
            if (!NativeMethods.IsWindow(hwnd)) return true;
            if (NativeMethods.GetWindowTextLength(hwnd) == 0) return true;

            // Skip known shell / desktop windows
            int len = NativeMethods.GetClassName(hwnd, classNameBuf, classNameBuf.Length);
            if (len > 0)
            {
                var className = new string(classNameBuf, 0, len);
                if (_shellClassNames.Contains(className)) return true;
            }

            // Skip tool windows (unless they also carry WS_EX_APPWINDOW)
            int exStyle = NativeMethods.GetWindowLong(hwnd, NativeMethods.GWL_EXSTYLE);
            if ((exStyle & NativeMethods.WS_EX_TOOLWINDOW) != 0
                && (exStyle & NativeMethods.WS_EX_APPWINDOW) == 0)
                return true;

            // Skip minimized windows
            var placement = NativeMethods.WINDOWPLACEMENT.Default;
            if (!NativeMethods.GetWindowPlacement(hwnd, ref placement)) return true;
            if (placement.showCmd == NativeMethods.SW_SHOWMINIMIZED) return true;

            // Skip windows on the target monitor — they stay behind on the original desktop
            var screen = Screen.FromHandle(hwnd);
            if (screen.DeviceName == targetMonitor.DeviceName) return true;

            // Skip already-pinned windows (they're visible on all desktops anyway)
            if (_vds.IsWindowPinned(hwnd)) return true;

            companions.Add(hwnd);
            return true;
        }, IntPtr.Zero);

        return companions;
    }

    /// <summary>
    /// Repositions a window to the center of the target monitor.
    /// Restores a maximized window first since maximized windows can't be freely repositioned.
    /// </summary>
    private static void MoveWindowToMonitor(IntPtr hwnd, Screen targetMonitor)
    {
        var currentScreen = Screen.FromHandle(hwnd);
        if (currentScreen.DeviceName == targetMonitor.DeviceName)
            return;

        // Restore if maximized — maximized windows are locked to their monitor
        var placement = NativeMethods.WINDOWPLACEMENT.Default;
        NativeMethods.GetWindowPlacement(hwnd, ref placement);
        if (placement.showCmd == NativeMethods.SW_SHOWMAXIMIZED)
            NativeMethods.ShowWindow(hwnd, NativeMethods.SW_RESTORE);

        // Get current window size
        if (!NativeMethods.GetWindowRect(hwnd, out var rect))
            return;
        int width = rect.Right - rect.Left;
        int height = rect.Bottom - rect.Top;

        // Center on target monitor's working area
        var wa = targetMonitor.WorkingArea;
        int x = wa.Left + (wa.Width - width) / 2;
        int y = wa.Top + (wa.Height - height) / 2;

        NativeMethods.SetWindowPos(hwnd, IntPtr.Zero, x, y, 0, 0,
            NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE);

        Trace.WriteLine($"FullScreenManager: Repositioned {hwnd} to monitor {targetMonitor.DeviceName} at ({x},{y})");
    }
}
