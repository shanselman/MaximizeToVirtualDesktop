using System.Diagnostics;
using MaximizeToVirtualDesktop.Interop;
using WindowsDesktop;

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
            return;
        }

        // 5. Switch to the new desktop
        if (!_vds.SwitchToDesktop(tempDesktop))
        {
            // Rollback: move window back, remove desktop
            Trace.WriteLine("FullScreenManager: Failed to switch desktop, rolling back.");
            var origDesktop = _vds.FindDesktop(originalDesktopId.Value);
            if (origDesktop != null) _vds.MoveWindowToDesktop(hwnd, origDesktop);
            _vds.RemoveDesktop(tempDesktop);
            return;
        }

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
        if (origDesktop != null)
        {
            if (windowStillExists) _vds.MoveWindowToDesktop(hwnd, origDesktop);
            _vds.SwitchToDesktop(origDesktop);
        }
        else
        {
            Trace.WriteLine("FullScreenManager: Original desktop no longer exists, leaving window on current.");
        }

        // 3. Remove temp desktop
        _vds.RemoveDesktop(entry.TempDesktop);

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
        if (origDesktop != null) _vds.SwitchToDesktop(origDesktop);

        // Then remove the temp desktop
        _vds.RemoveDesktop(entry.TempDesktop);
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
