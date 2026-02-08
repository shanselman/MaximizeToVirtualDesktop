using System.Diagnostics;
using MaximizeToVirtualDesktop.Interop;
using WindowsDesktop;

namespace MaximizeToVirtualDesktop;

internal sealed record TrackingEntry(
    IntPtr Hwnd,
    Guid OriginalDesktopId,
    Guid TempDesktopId,
    VirtualDesktop TempDesktop,
    string? ProcessName,
    NativeMethods.WINDOWPLACEMENT OriginalPlacement);

/// <summary>
/// Thread-safe tracker for windows we've sent to virtual desktops.
/// All mutations happen on the UI thread via Control.BeginInvoke,
/// but we lock defensively anyway.
/// </summary>
internal sealed class FullScreenTracker
{
    private readonly Dictionary<IntPtr, TrackingEntry> _entries = new();
    private readonly object _lock = new();

    public bool IsTracked(IntPtr hwnd)
    {
        lock (_lock) return _entries.ContainsKey(hwnd);
    }

    public TrackingEntry? Get(IntPtr hwnd)
    {
        lock (_lock) return _entries.GetValueOrDefault(hwnd);
    }

    public void Track(IntPtr hwnd, Guid originalDesktopId, Guid tempDesktopId,
        VirtualDesktop tempDesktop, string? processName,
        NativeMethods.WINDOWPLACEMENT originalPlacement)
    {
        lock (_lock)
        {
            _entries[hwnd] = new TrackingEntry(hwnd, originalDesktopId, tempDesktopId,
                tempDesktop, processName, originalPlacement);
            Trace.WriteLine($"FullScreenTracker: Now tracking {hwnd} (total: {_entries.Count})");
        }
        PersistToDisk();
    }

    public TrackingEntry? Untrack(IntPtr hwnd)
    {
        TrackingEntry? entry;
        lock (_lock)
        {
            if (_entries.Remove(hwnd, out entry))
            {
                Trace.WriteLine($"FullScreenTracker: Untracked {hwnd} (total: {_entries.Count})");
            }
            else
            {
                return null;
            }
        }
        PersistToDisk();
        return entry;
    }

    /// <summary>
    /// Clear all entries, releasing COM references. Called when Explorer restarts
    /// (Windows destroys all virtual desktops, making our tracked refs stale).
    /// </summary>
    public void ClearAll()
    {
        lock (_lock)
        {
            var count = _entries.Count;
            _entries.Clear();
            Trace.WriteLine($"FullScreenTracker: Cleared {count} stale entries (Explorer restart).");
        }
        TrackerPersistence.Delete();
    }

    /// <summary>Returns all tracked entries (snapshot).</summary>
    public IReadOnlyList<TrackingEntry> GetAll()
    {
        lock (_lock) return _entries.Values.ToList();
    }

    /// <summary>Returns tracked window handles whose windows no longer exist.</summary>
    public IReadOnlyList<IntPtr> GetStaleHandles()
    {
        lock (_lock)
        {
            return _entries.Keys
                .Where(hwnd => !NativeMethods.IsWindow(hwnd))
                .ToList();
        }
    }

    public int Count
    {
        get { lock (_lock) return _entries.Count; }
    }

    private void PersistToDisk()
    {
        List<TrackerPersistence.PersistedEntry> snapshot;
        lock (_lock)
        {
            snapshot = _entries.Values.Select(e =>
                new TrackerPersistence.PersistedEntry(e.TempDesktopId, e.ProcessName, DateTime.UtcNow)).ToList();
        }
        TrackerPersistence.Save(snapshot);
    }
}
