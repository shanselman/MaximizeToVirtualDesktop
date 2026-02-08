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
    /// If the window belongs to a multi-window process (e.g., VS Code with detached tabs),
    /// moves all visible windows from that process together.
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

        // 1. Find all windows from the same process
        var allWindows = GetAllProcessWindows(hwnd);
        if (allWindows.Count == 0)
        {
            Trace.WriteLine("FullScreenManager: No valid windows found for process, aborting.");
            return;
        }

        // Ensure the original window is first in the list (will be the one we maximize)
        allWindows.Remove(hwnd);
        allWindows.Insert(0, hwnd);

        // 2. Record original state for all windows
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

        // Store placements for all windows
        var windowPlacements = new Dictionary<IntPtr, NativeMethods.WINDOWPLACEMENT>();
        foreach (var window in allWindows)
        {
            var placement = NativeMethods.WINDOWPLACEMENT.Default;
            if (NativeMethods.GetWindowPlacement(window, ref placement))
            {
                windowPlacements[window] = placement;
            }
        }

        // 3. Create new virtual desktop
        var (tempDesktop, tempDesktopId) = _vds.CreateDesktop();
        if (tempDesktop == null || tempDesktopId == null)
        {
            Trace.WriteLine("FullScreenManager: Failed to create desktop, aborting.");
            return;
        }

        // 4. Name the desktop after the window title (or process name as fallback)
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

        // 5. Move all windows to new desktop
        var movedWindows = new List<IntPtr>();
        foreach (var window in allWindows)
        {
            if (_vds.MoveWindowToDesktop(window, tempDesktop))
            {
                movedWindows.Add(window);
            }
            else
            {
                Trace.WriteLine($"FullScreenManager: Failed to move window {window}, rolling back.");
                // Rollback: move already-moved windows back
                var origDesktop = _vds.FindDesktop(originalDesktopId.Value);
                try
                {
                    if (origDesktop != null)
                    {
                        foreach (var movedWindow in movedWindows)
                        {
                            _vds.MoveWindowToDesktop(movedWindow, origDesktop);
                        }
                    }
                }
                finally
                {
                    if (origDesktop != null) Marshal.ReleaseComObject(origDesktop);
                }
                _vds.RemoveDesktop(tempDesktop);
                Marshal.ReleaseComObject(tempDesktop);
                return;
            }
        }

        // 6. Switch to the new desktop
        if (!_vds.SwitchToDesktop(tempDesktop))
        {
            // Rollback: move all windows back, remove desktop
            Trace.WriteLine("FullScreenManager: Failed to switch desktop, rolling back.");
            var origDesktop = _vds.FindDesktop(originalDesktopId.Value);
            try
            {
                if (origDesktop != null)
                {
                    foreach (var window in movedWindows)
                    {
                        _vds.MoveWindowToDesktop(window, origDesktop);
                    }
                }
            }
            finally
            {
                if (origDesktop != null) Marshal.ReleaseComObject(origDesktop);
            }
            _vds.RemoveDesktop(tempDesktop);
            Marshal.ReleaseComObject(tempDesktop);
            return;
        }

        // 7. Maximize the primary window — delay lets desktop switch animation finish first
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

        // 8. Track all windows
        // Note: All windows share the same tempDesktop COM reference.
        // When restoring/destroying, we only release the COM reference once per desktop.
        foreach (var window in movedWindows)
        {
            var placement = windowPlacements.ContainsKey(window) 
                ? windowPlacements[window] 
                : NativeMethods.WINDOWPLACEMENT.Default;
            _tracker.Track(window, originalDesktopId.Value, tempDesktopId.Value, tempDesktop, processName, placement);
        }

        var windowCount = movedWindows.Count;
        var message = windowCount > 1 
            ? $"→ Virtual Desktop ({windowCount} windows)" 
            : "→ Virtual Desktop";
        NotificationOverlay.ShowNotification(message, processName ?? "", hwnd);
        Trace.WriteLine($"FullScreenManager: Successfully moved {windowCount} window(s) to desktop {tempDesktopId}");
    }

    /// <summary>
    /// Restore a tracked window: move it back to its original desktop, restore window state,
    /// switch back, and remove the temp desktop.
    /// If multiple windows share the same temp desktop, restores all of them together.
    /// </summary>
    public void Restore(IntPtr hwnd)
    {
        var entry = _tracker.Get(hwnd);
        if (entry == null)
        {
            Trace.WriteLine($"FullScreenManager: hwnd {hwnd} not tracked, ignoring restore.");
            return;
        }

        // Find all windows on the same temp desktop (may be multiple if from same process)
        var relatedWindows = _tracker.GetAll()
            .Where(e => e.TempDesktopId == entry.TempDesktopId)
            .ToList();

        Trace.WriteLine($"FullScreenManager: Restoring {relatedWindows.Count} window(s) from temp desktop {entry.TempDesktopId}");

        // Untrack all related windows first to prevent reentrant calls
        foreach (var relatedEntry in relatedWindows)
        {
            _tracker.Untrack(relatedEntry.Hwnd);
        }

        var origDesktop = _vds.FindDesktop(entry.OriginalDesktopId);
        try
        {
            // Restore all windows
            foreach (var relatedEntry in relatedWindows)
            {
                var windowStillExists = NativeMethods.IsWindow(relatedEntry.Hwnd);
                
                // Restore window placement
                if (windowStillExists)
                {
                    var placement = relatedEntry.OriginalPlacement;
                    NativeMethods.SetWindowPlacement(relatedEntry.Hwnd, ref placement);
                }

                // Move window back to original desktop
                if (origDesktop != null && windowStillExists)
                {
                    _vds.MoveWindowToDesktop(relatedEntry.Hwnd, origDesktop);
                }
            }

            // Switch back to original desktop
            if (origDesktop != null)
            {
                _vds.SwitchToDesktop(origDesktop);
            }
            else
            {
                Trace.WriteLine("FullScreenManager: Original desktop no longer exists, leaving windows on current.");
            }
        }
        finally
        {
            if (origDesktop != null) Marshal.ReleaseComObject(origDesktop);
        }

        // Remove temp desktop and release its COM reference (only once for all windows)
        // All related windows share the same TempDesktop COM reference, so we only release it once
        _vds.RemoveDesktop(entry.TempDesktop);
        Marshal.ReleaseComObject(entry.TempDesktop);

        // Set focus on the primary restored window
        if (NativeMethods.IsWindow(hwnd))
        {
            NativeMethods.SetForegroundWindow(hwnd);
        }

        var windowCount = relatedWindows.Count;
        var message = windowCount > 1 ? $"← Restored ({windowCount} windows)" : "← Restored";
        NotificationOverlay.ShowNotification(message, entry.ProcessName ?? "", hwnd);
        Trace.WriteLine($"FullScreenManager: Restored {windowCount} window(s) to original desktop.");
    }

    /// <summary>
    /// Called when a tracked window is destroyed (closed). Clean up its temp desktop.
    /// If multiple windows share the same temp desktop, only removes the desktop when
    /// the last window is destroyed.
    /// </summary>
    public void HandleWindowDestroyed(IntPtr hwnd)
    {
        var entry = _tracker.Untrack(hwnd);
        if (entry == null) return;

        Trace.WriteLine($"FullScreenManager: Tracked window {hwnd} destroyed.");

        // Check if other windows are still on the same temp desktop
        var remainingWindows = _tracker.GetAll()
            .Where(e => e.TempDesktopId == entry.TempDesktopId)
            .ToList();

        if (remainingWindows.Count > 0)
        {
            Trace.WriteLine($"FullScreenManager: {remainingWindows.Count} window(s) still on temp desktop {entry.TempDesktopId}, not removing desktop yet.");
            // Don't release the COM reference or remove desktop - other windows still need it
            return;
        }

        Trace.WriteLine($"FullScreenManager: Last window on temp desktop {entry.TempDesktopId}, cleaning up.");

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

        // Remove the temp desktop and release its COM reference (only once, now that it's the last window)
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

    /// <summary>
    /// Get all visible windows belonging to the same process as the given window.
    /// Returns windows that are:
    /// - Visible
    /// - Not owned by another window (not dialogs/popups)
    /// - Belong to the same process
    /// - Have a window title (filters out background/helper windows)
    /// 
    /// This is used to find sibling windows (e.g., VS Code detached tabs) that should
    /// be moved together to the virtual desktop.
    /// </summary>
    private List<IntPtr> GetAllProcessWindows(IntPtr hwnd)
    {
        NativeMethods.GetWindowThreadProcessId(hwnd, out int targetPid);
        var windows = new List<IntPtr>();

        NativeMethods.EnumWindows((enumHwnd, _) =>
        {
            // Skip if not visible
            if (!NativeMethods.IsWindowVisible(enumHwnd))
                return true;

            // Skip if owned by another window (dialogs, popups)
            if (NativeMethods.GetWindow(enumHwnd, NativeMethods.GW_OWNER) != IntPtr.Zero)
                return true;

            // Skip if different process
            NativeMethods.GetWindowThreadProcessId(enumHwnd, out int enumPid);
            if (enumPid != targetPid)
                return true;

            // Skip if no title - this filters out background/helper windows.
            // Top-level application windows (main windows, detached tabs) have titles.
            // Note: This may miss windows that are temporarily title-less during initialization,
            // but that's acceptable for this use case.
            int textLength = NativeMethods.GetWindowTextLength(enumHwnd);
            if (textLength == 0)
                return true;

            windows.Add(enumHwnd);
            return true; // Continue enumeration
        }, IntPtr.Zero);

        Trace.WriteLine($"FullScreenManager: Found {windows.Count} window(s) for process {targetPid}");
        return windows;
    }
}
