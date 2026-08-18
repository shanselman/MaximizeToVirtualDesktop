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
    private readonly AppSettings _settings;
    private readonly HashSet<IntPtr> _inFlight = new();

    public FullScreenManager(VirtualDesktopService vds, FullScreenTracker tracker, AppSettings settings)
    {
        _vds = vds;
        _tracker = tracker;
        _settings = settings;
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
    /// Only the clicked window is moved; other windows from the same process are not affected.
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

        // 2. Create new virtual desktop
        var (tempDesktop, tempDesktopId) = _vds.CreateDesktop();
        if (tempDesktop == null || tempDesktopId == null)
        {
            Trace.WriteLine("FullScreenManager: Failed to create desktop, aborting.");
            return;
        }

        // 3. Name the desktop
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
        catch { /* Non-critical */ }

        // 4. Maximize, move, and switch
        bool elevated = NativeMethods.IsWindowElevated(hwnd);

        if (!elevated && NativeMethods.IsWindow(hwnd))
        {
            NativeMethods.ShowWindow(hwnd, NativeMethods.SW_MAXIMIZE);
            Thread.Sleep(250);
        }
        else if (elevated)
        {
            Trace.WriteLine("FullScreenManager: Window is elevated, cannot maximize via UIPI.");
            if (_settings.ShowSwitchPopup)
            {
                NotificationOverlay.ShowNotification("⚠ Elevated Window",
                    "Press Win+↑ to maximize", hwnd);
            }
        }

        if (!_vds.MoveWindowToDesktop(hwnd, tempDesktop))
        {
            RollbackMaximize(tempDesktop, originalDesktopId.Value, hwnd,
                originalPlacement, elevated, windowMoved: false);
            return;
        }

        if (!_vds.SwitchToDesktop(tempDesktop))
        {
            RollbackMaximize(tempDesktop, originalDesktopId.Value, hwnd,
                originalPlacement, elevated, windowMoved: true);
            return;
        }

        NativeMethods.SetForegroundWindow(hwnd);

        // 5. Track
        _tracker.Track(hwnd, originalDesktopId.Value, tempDesktopId.Value, tempDesktop, processName, originalPlacement);

        if (_settings.ShowSwitchPopup)
        {
            NotificationOverlay.ShowNotification("→ Virtual Desktop", processName ?? "", hwnd);
        }
        Trace.WriteLine($"FullScreenManager: Successfully moved window to desktop {tempDesktopId}");
    }

    /// <summary>
    /// Roll back a failed maximize operation and remove the temporary desktop.
    /// </summary>
    private void RollbackMaximize(IVirtualDesktop tempDesktop, Guid originalDesktopId,
        IntPtr hwnd, NativeMethods.WINDOWPLACEMENT originalPlacement, bool elevated,
        bool windowMoved)
    {
        Trace.WriteLine("FullScreenManager: Maximize operation failed, rolling back.");
        var origDesktop = _vds.FindDesktop(originalDesktopId);
        try
        {
            if (windowMoved && origDesktop != null)
                _vds.MoveWindowToDesktop(hwnd, origDesktop);

            if (origDesktop != null && _vds.GetCurrentDesktopId() != originalDesktopId)
                _vds.SwitchToDesktop(origDesktop);
        }
        finally { if (origDesktop != null) Marshal.ReleaseComObject(origDesktop); }

        if (!elevated && NativeMethods.IsWindow(hwnd))
        {
            var placement = originalPlacement;
            NativeMethods.SetWindowPlacement(hwnd, ref placement);
        }

        if (!_vds.RemoveDesktop(tempDesktop))
            Trace.WriteLine("FullScreenManager: Failed to remove temporary desktop during rollback.");
        Marshal.ReleaseComObject(tempDesktop);
    }

    /// <summary>
    /// Restore a tracked window: move it back to its original desktop, restore window state,
    /// switch back, and remove the temp desktop.
    /// When <paramref name="keepMinimized"/> or <paramref name="keepHidden"/> is true,
    /// preserve that visibility state while restoring the original placement.
    /// </summary>
    public void Restore(IntPtr hwnd, bool keepMinimized = false, bool keepHidden = false)
    {
        var entry = _tracker.Get(hwnd);
        if (entry == null)
        {
            Trace.WriteLine($"FullScreenManager: hwnd {hwnd} not tracked, ignoring restore.");
            return;
        }

        Trace.WriteLine($"FullScreenManager: Restoring window {hwnd} from temp desktop {entry.TempDesktopId}");

        var origDesktop = _vds.FindDesktop(entry.OriginalDesktopId);
        bool restored = true;
        try
        {
            if (origDesktop != null && NativeMethods.IsWindow(hwnd))
                restored = _vds.MoveWindowToDesktop(hwnd, origDesktop);

            var currentDesktopId = _vds.GetCurrentDesktopId();
            if (origDesktop != null
                && (currentDesktopId == null || currentDesktopId == entry.TempDesktopId))
            {
                restored = _vds.SwitchToDesktop(origDesktop) && restored;
            }

            if (restored && NativeMethods.IsWindow(hwnd))
            {
                RestoreWindowPlacement(hwnd, entry, keepMinimized, keepHidden);
            }
        }
        finally
        {
            if (origDesktop != null) Marshal.ReleaseComObject(origDesktop);
        }

        if (!restored)
        {
            Trace.WriteLine($"FullScreenManager: Could not restore window {hwnd}; keeping it tracked for retry.");
            return;
        }

        if (!_vds.RemoveDesktop(entry.TempDesktop))
        {
            Trace.WriteLine($"FullScreenManager: Could not remove desktop {entry.TempDesktopId}; keeping it tracked for recovery.");
            return;
        }

        Marshal.ReleaseComObject(entry.TempDesktop);
        _tracker.Untrack(hwnd);

        if (!keepMinimized && !keepHidden && NativeMethods.IsWindow(hwnd))
        {
            NativeMethods.SetForegroundWindow(hwnd);
        }

        if (_settings.ShowSwitchPopup && !keepHidden)
        {
            NotificationOverlay.ShowNotification("← Restored", entry.ProcessName ?? "", hwnd);
        }
        Trace.WriteLine($"FullScreenManager: Restored window to original desktop.");
    }

    private static void RestoreWindowPlacement(IntPtr hwnd, TrackingEntry entry,
        bool keepMinimized, bool keepHidden)
    {
        if (keepMinimized || keepHidden)
        {
            var placement = entry.OriginalPlacement;
            if (entry.OriginalPlacement.showCmd != NativeMethods.SW_MAXIMIZE)
                placement.flags &= ~NativeMethods.WPF_RESTORETOMAXIMIZED;
            placement.showCmd = keepHidden
                ? NativeMethods.SW_HIDE
                : NativeMethods.SW_SHOWMINNOACTIVE;
            NativeMethods.SetWindowPlacement(hwnd, ref placement);
            return;
        }

        var current = NativeMethods.WINDOWPLACEMENT.Default;
        if (NativeMethods.GetWindowPlacement(hwnd, ref current)
            && current.showCmd == NativeMethods.SW_MAXIMIZE)
        {
            var placement = entry.OriginalPlacement;
            NativeMethods.SetWindowPlacement(hwnd, ref placement);
        }
        NativeMethods.ShowWindow(hwnd, (int)NativeMethods.SW_SHOWNORMAL);
    }

    /// <summary>
    /// Called when a tracked window is destroyed (closed). Clean up its temp desktop.
    /// </summary>
    public void HandleWindowDestroyed(IntPtr hwnd)
    {
        var entry = _tracker.Get(hwnd);
        if (entry == null) return;

        Trace.WriteLine($"FullScreenManager: Tracked window {hwnd} destroyed.");

        Trace.WriteLine($"FullScreenManager: Cleaning up temp desktop {entry.TempDesktopId}");

        var origDesktop = _vds.FindDesktop(entry.OriginalDesktopId);
        bool switched = true;
        try
        {
            var currentDesktopId = _vds.GetCurrentDesktopId();
            if (origDesktop != null
                && (currentDesktopId == null || currentDesktopId == entry.TempDesktopId))
            {
                switched = _vds.SwitchToDesktop(origDesktop);
            }
        }
        finally
        {
            if (origDesktop != null) Marshal.ReleaseComObject(origDesktop);
        }

        if (!switched || !_vds.RemoveDesktop(entry.TempDesktop))
        {
            Trace.WriteLine($"FullScreenManager: Could not clean up desktop {entry.TempDesktopId}; keeping it tracked for retry.");
            return;
        }

        Marshal.ReleaseComObject(entry.TempDesktop);
        _tracker.Untrack(hwnd);
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

}
