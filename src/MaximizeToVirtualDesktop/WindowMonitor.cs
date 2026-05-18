using System.Diagnostics;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using MaximizeToVirtualDesktop.Interop;

namespace MaximizeToVirtualDesktop;

/// <summary>
/// Uses SetWinEventHook to monitor tracked windows for state changes (un-maximize, close).
/// All callbacks are marshaled to the UI thread.
/// </summary>
internal sealed class WindowMonitor : IDisposable
{
    private const int RdpGeometryTolerance = 2;
    private const int StandardFallbackDelayMs = 200;
    private const int RdpFallbackDelayMs = 1000;
    private const int StandardRestoreDelayMs = 100;
    private const int RdpRestoreDelayMs = 1000;

    private enum FullscreenState
    {
        None,
        Maximized,
        RdpFullscreen
    }

    private readonly FullScreenManager _manager;
    private readonly FullScreenTracker _tracker;
    private readonly Control _syncControl;
    private readonly AppSettings _settings;
    private IntPtr _locationChangeHook;
    private IntPtr _destroyHook;
    private bool _disposed;

    // Must be stored as fields to prevent GC collection of the delegate
    private readonly NativeMethods.WinEventProc _locationChangeProc;
    private readonly NativeMethods.WinEventProc _destroyProc;
    private readonly NativeMethods.WinEventProc _moveSizeEndProc;
    private IntPtr _moveSizeEndHook;
    // Track windows that have been maximized but need to wait for resize end
    // Access to this set must happen only on the UI thread.
    private readonly Dictionary<IntPtr, FullscreenState> _pendingMaximize = new();
    private readonly HashSet<IntPtr> _pendingRdpMaximize = new();
    private readonly HashSet<IntPtr> _pendingRestore = new();
    private bool _rdpMaximizeScheduled;

    public WindowMonitor(FullScreenManager manager, FullScreenTracker tracker, Control syncControl, AppSettings settings)
    {
        _manager = manager;
        _tracker = tracker;
        _syncControl = syncControl;
        _settings = settings;

        _locationChangeProc = OnLocationChange;
        _destroyProc = OnDestroy;
        _moveSizeEndProc = OnMoveSizeEnd;
    }

    public void Start()
    {
        if (_locationChangeHook != IntPtr.Zero) return;

        // EVENT_OBJECT_LOCATIONCHANGE fires when window state changes (including maximize/restore)
        _locationChangeHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_OBJECT_LOCATIONCHANGE,
            NativeMethods.EVENT_OBJECT_LOCATIONCHANGE,
            IntPtr.Zero, _locationChangeProc,
            0, 0, NativeMethods.WINEVENT_OUTOFCONTEXT);

        // EVENT_SYSTEM_MOVESIZEEND fires after a window finishes moving or resizing (including maximize via shortcuts)
        _moveSizeEndHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_SYSTEM_MOVESIZEEND,
            NativeMethods.EVENT_SYSTEM_MOVESIZEEND,
            IntPtr.Zero, _moveSizeEndProc,
            0, 0, NativeMethods.WINEVENT_OUTOFCONTEXT);

        // EVENT_OBJECT_DESTROY fires when a window is closed
        _destroyHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_OBJECT_DESTROY,
            NativeMethods.EVENT_OBJECT_DESTROY,
            IntPtr.Zero, _destroyProc,
            0, 0, NativeMethods.WINEVENT_OUTOFCONTEXT);

        if (_locationChangeHook == IntPtr.Zero || _destroyHook == IntPtr.Zero)
        {
            Trace.WriteLine("WindowMonitor: Failed to set one or more WinEvent hooks.");
        }
        else
        {
            Trace.WriteLine("WindowMonitor: Started monitoring.");
        }
    }

    private void OnLocationChange(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint idEventThread, uint dwmsEventTime)
    {
        // Only care about top-level window changes (OBJID_WINDOW)
        if (idObject != NativeMethods.OBJID_WINDOW || idChild != 0) return;

        // If the window is already tracked, check if it is being restored (i.e., no longer maximized)
        var trackedEntry = _tracker.Get(hwnd);
        if (trackedEntry != null)
        {
            if (!trackedEntry.IsRdpFullscreen && GetFullscreenState(hwnd) == FullscreenState.None)
            {
                Trace.WriteLine($"WindowMonitor: Tracked window {hwnd} restored via location change.");
                QueueRestore(hwnd, StandardRestoreDelayMs);
            }

            // Still maximized, or an RDP fullscreen window with noisy geometry changes.
            return;
        }

        // Not tracked yet: check for a new maximize event (including via shortcut)
        var fullscreenState = GetFullscreenState(hwnd);
        if (fullscreenState == FullscreenState.None) return;
        if (!IsForegroundWindowOrOwner(hwnd)) return;

        bool shiftHeld = (NativeMethods.GetAsyncKeyState(NativeMethods.VK_SHIFT) & 0x8000) != 0;
        bool triggerVirtualDesktop = _settings.InvertShiftClick ? !shiftHeld : shiftHeld;
        if (triggerVirtualDesktop)
        {
            if (fullscreenState == FullscreenState.RdpFullscreen)
            {
                QueueRdpMaximize(hwnd);
                return;
            }

            // Defer maximization until after the resize operation completes
            MarshalToUiThread(() =>
            {
                if (!_pendingMaximize.ContainsKey(hwnd))
                {
                    _pendingMaximize[hwnd] = fullscreenState;
                    Trace.WriteLine($"WindowMonitor: Queued maximize for window {hwnd} after resize end.");
                    int delayMs = fullscreenState == FullscreenState.RdpFullscreen
                        ? RdpFallbackDelayMs
                        : StandardFallbackDelayMs;

                    // Schedule a fallback in case MoveSizeEnd does not fire (e.g., keyboard shortcut)
                    _ = Task.Run(async () =>
                    {
                        await Task.Delay(delayMs);
                        // Marshal the check/remove back onto the UI thread
                        MarshalToUiThread(() =>
                        {
                            if (_pendingMaximize.Remove(hwnd, out var pendingState))
                            {
                                var currentState = GetFullscreenState(hwnd);
                                if (currentState != FullscreenState.None && IsForegroundWindowOrOwner(hwnd))
                                {
                                    Trace.WriteLine($"WindowMonitor: Fallback processing for pending maximize window {hwnd}.");
                                    _manager.MaximizeToDesktop(hwnd, pendingState == FullscreenState.RdpFullscreen);
                                }
                                else
                                {
                                    Trace.WriteLine($"WindowMonitor: Fallback detected pending window {hwnd} is no longer fullscreen.");
                                }
                            }
                        });
                    });
                }
            });
        }
    }

    private void QueueRdpMaximize(IntPtr hwnd)
    {
        MarshalToUiThread(() =>
        {
            _pendingRdpMaximize.Add(hwnd);
            Trace.WriteLine($"WindowMonitor: Queued RDP fullscreen candidate {hwnd}.");

            if (_rdpMaximizeScheduled) return;
            _rdpMaximizeScheduled = true;

            _ = Task.Run(async () =>
            {
                await Task.Delay(RdpFallbackDelayMs);

                MarshalToUiThread(() =>
                {
                    var candidates = _pendingRdpMaximize.ToList();
                    _pendingRdpMaximize.Clear();
                    _rdpMaximizeScheduled = false;

                    var selected = candidates
                        .Where(hwnd => GetFullscreenState(hwnd) == FullscreenState.RdpFullscreen)
                        .OrderByDescending(IsForegroundWindowOrOwner)
                        .FirstOrDefault(_manager.CanManageWindow);

                    if (selected == IntPtr.Zero)
                    {
                        Trace.WriteLine($"WindowMonitor: No manageable RDP fullscreen candidate among {candidates.Count} window(s).");
                        return;
                    }

                    Trace.WriteLine($"WindowMonitor: Processing RDP fullscreen candidate {selected} from {candidates.Count} window(s).");
                    _manager.MaximizeToDesktop(selected, isRdpFullscreen: true);
                });
            });
        });
    }

    private async void OnMoveSizeEnd(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint idEventThread, uint dwmsEventTime)
    {
        // Only care about top-level window changes (OBJID_WINDOW)
        if (idObject != NativeMethods.OBJID_WINDOW || idChild != 0) return;

        // If this window was pending maximize, handle it now
        bool wasPending = false;
        FullscreenState pendingState = FullscreenState.None;
        if (!_syncControl.IsDisposed && _syncControl.IsHandleCreated)
        {
            wasPending = (bool)_syncControl.Invoke(new Func<bool>(() =>
            {
                if (!_pendingMaximize.Remove(hwnd, out pendingState)) return false;
                return true;
            }));
        }
        if (wasPending)
        {
            Trace.WriteLine($"WindowMonitor: MoveSizeEnd triggered for pending maximize window {hwnd}.");
            MarshalToUiThread(() =>
            {
                if (GetFullscreenState(hwnd) != FullscreenState.None)
                {
                    _manager.MaximizeToDesktop(hwnd, pendingState == FullscreenState.RdpFullscreen);
                }
                else
                {
                    Trace.WriteLine($"WindowMonitor: MoveSizeEnd ignored pending window {hwnd} because it is no longer fullscreen.");
                }
            });
            return;
        }

        // Handle only tracked windows that are being restored (not maximized)
        var trackedEntry = _tracker.Get(hwnd);
        if (trackedEntry == null) return;
        if (GetFullscreenState(hwnd) == FullscreenState.None)
        {
            Trace.WriteLine($"WindowMonitor: Tracked window {hwnd} un-maximized via move/size, restoring.");
            await Task.Delay(trackedEntry.IsRdpFullscreen ? RdpRestoreDelayMs : StandardRestoreDelayMs);
            QueueRestore(hwnd, 0);
        }
    }

    private void OnDestroy(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint idEventThread, uint dwmsEventTime)
    {
        if (idObject != NativeMethods.OBJID_WINDOW || idChild != 0) return;
        if (!_tracker.IsTracked(hwnd)) return;

        Trace.WriteLine($"WindowMonitor: Tracked window {hwnd} destroyed.");
        MarshalToUiThread(() => _manager.HandleWindowDestroyed(hwnd));
    }

    private void MarshalToUiThread(Action action)
    {
        if (_syncControl.IsDisposed || !_syncControl.IsHandleCreated) return;

        try
        {
            _syncControl.BeginInvoke(action);
        }
        catch (ObjectDisposedException)
        {
            // App is shutting down
        }
    }

    private void QueueRestore(IntPtr hwnd, int delayMs)
    {
        MarshalToUiThread(() =>
        {
            if (!_pendingRestore.Add(hwnd)) return;

            _ = Task.Run(async () =>
            {
                if (delayMs > 0) await Task.Delay(delayMs);

                MarshalToUiThread(() =>
                {
                    _pendingRestore.Remove(hwnd);

                    var entry = _tracker.Get(hwnd);
                    if (entry == null) return;
                    if (GetFullscreenState(hwnd) != FullscreenState.None)
                    {
                        Trace.WriteLine($"WindowMonitor: Restore ignored for {hwnd} because it is fullscreen again.");
                        return;
                    }

                    _manager.Restore(hwnd);
                });
            });
        });
    }

    private static FullscreenState GetFullscreenState(IntPtr hwnd)
    {
        var placement = NativeMethods.WINDOWPLACEMENT.Default;
        if (NativeMethods.GetWindowPlacement(hwnd, ref placement) && placement.showCmd == NativeMethods.SW_MAXIMIZE)
        {
            return FullscreenState.Maximized;
        }

        return IsRdpFullscreenTransition(hwnd)
            ? FullscreenState.RdpFullscreen
            : FullscreenState.None;
    }

    private static bool IsRdpFullscreenTransition(IntPtr hwnd)
    {
        if (!IsRdpWindow(hwnd)) return false;
        if (!NativeMethods.GetWindowRect(hwnd, out var windowRect)) return false;
        if (!IsBorderlessWindow(hwnd)) return false;

        var monitor = NativeMethods.MonitorFromWindow(hwnd, NativeMethods.MONITOR_DEFAULTTONEAREST);
        if (monitor == IntPtr.Zero) return false;

        var monitorInfo = NativeMethods.MONITORINFO.Default;
        if (!NativeMethods.GetMonitorInfo(monitor, ref monitorInfo)) return false;

        return IsRectFullscreen(windowRect, monitorInfo.rcMonitor, GetVirtualScreenRect());
    }

    private static bool IsRdpWindow(IntPtr hwnd)
    {
        if (NativeMethods.GetWindowThreadProcessId(hwnd, out int processId) == 0) return false;
        if (processId <= 0) return false;

        try
        {
            using var process = Process.GetProcessById(processId);
            return process.ProcessName.Equals("mstsc", StringComparison.OrdinalIgnoreCase)
                || process.ProcessName.Equals("msrdc", StringComparison.OrdinalIgnoreCase);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"WindowMonitor: Failed to inspect process for hwnd {hwnd}: {ex.Message}");
            return false;
        }
    }

    private static bool IsBorderlessWindow(IntPtr hwnd)
    {
        NativeMethods.SetLastError(0);
        long style = NativeMethods.GetWindowLongPtr(hwnd, NativeMethods.GWL_STYLE).ToInt64();
        if (style == 0 && Marshal.GetLastWin32Error() != 0) return false;

        return (style & (NativeMethods.WS_CAPTION | NativeMethods.WS_THICKFRAME)) == 0;
    }

    private static bool IsForegroundWindowOrOwner(IntPtr hwnd)
    {
        var foreground = NativeMethods.GetForegroundWindow();
        if (foreground == IntPtr.Zero) return false;
        if (foreground == hwnd) return true;

        return NativeMethods.GetAncestor(hwnd, NativeMethods.GA_ROOTOWNER) == foreground
            || NativeMethods.GetAncestor(foreground, NativeMethods.GA_ROOTOWNER) == hwnd;
    }

    private static NativeMethods.RECT GetVirtualScreenRect()
    {
        int left = NativeMethods.GetSystemMetrics(NativeMethods.SM_XVIRTUALSCREEN);
        int top = NativeMethods.GetSystemMetrics(NativeMethods.SM_YVIRTUALSCREEN);
        int width = NativeMethods.GetSystemMetrics(NativeMethods.SM_CXVIRTUALSCREEN);
        int height = NativeMethods.GetSystemMetrics(NativeMethods.SM_CYVIRTUALSCREEN);
        return new NativeMethods.RECT
        {
            Left = left,
            Top = top,
            Right = left + width,
            Bottom = top + height
        };
    }

    internal static bool IsRectFullscreen(NativeMethods.RECT windowRect, NativeMethods.RECT monitorRect, NativeMethods.RECT virtualScreenRect)
    {
        return RectEquals(windowRect, monitorRect, RdpGeometryTolerance)
            || RectEquals(windowRect, virtualScreenRect, RdpGeometryTolerance);
    }

    internal static bool RectEquals(NativeMethods.RECT a, NativeMethods.RECT b, int tolerance)
    {
        return Math.Abs(a.Left - b.Left) <= tolerance
            && Math.Abs(a.Top - b.Top) <= tolerance
            && Math.Abs(a.Right - b.Right) <= tolerance
            && Math.Abs(a.Bottom - b.Bottom) <= tolerance;
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_locationChangeHook != IntPtr.Zero)
        {
            NativeMethods.UnhookWinEvent(_locationChangeHook);
            _locationChangeHook = IntPtr.Zero;
        }
        if (_destroyHook != IntPtr.Zero)
        {
            NativeMethods.UnhookWinEvent(_destroyHook);
            _destroyHook = IntPtr.Zero;
        }
        if (_moveSizeEndHook != IntPtr.Zero)
        {
            NativeMethods.UnhookWinEvent(_moveSizeEndHook);
            _moveSizeEndHook = IntPtr.Zero;
        }

        Trace.WriteLine("WindowMonitor: Disposed.");
    }
}
