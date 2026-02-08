using System.Diagnostics;
using MaximizeToVirtualDesktop.Interop;

namespace MaximizeToVirtualDesktop;

internal sealed record TrackingEntry(
    IntPtr Hwnd,
    Guid OriginalDesktopId,
    IVirtualDesktop TempDesktop,
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

    public void Track(IntPtr hwnd, Guid originalDesktopId, IVirtualDesktop tempDesktop,
        NativeMethods.WINDOWPLACEMENT originalPlacement)
    {
        lock (_lock)
        {
            _entries[hwnd] = new TrackingEntry(hwnd, originalDesktopId, tempDesktop, originalPlacement);
            Trace.WriteLine($"FullScreenTracker: Now tracking {hwnd} (total: {_entries.Count})");
        }
    }

    public TrackingEntry? Untrack(IntPtr hwnd)
    {
        lock (_lock)
        {
            if (_entries.Remove(hwnd, out var entry))
            {
                Trace.WriteLine($"FullScreenTracker: Untracked {hwnd} (total: {_entries.Count})");
                return entry;
            }
            return null;
        }
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
}
