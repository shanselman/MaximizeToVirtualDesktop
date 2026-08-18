using System.Diagnostics;
using System.Collections.Generic;
using System.Threading.Tasks;
using MaximizeToVirtualDesktop.Interop;

namespace MaximizeToVirtualDesktop;

/// <summary>
/// Uses SetWinEventHook to monitor tracked windows for state changes (un-maximize, close).
/// All callbacks are marshaled to the UI thread.
/// </summary>
internal sealed class WindowMonitor : IDisposable
{
    private readonly FullScreenManager _manager;
    private readonly FullScreenTracker _tracker;
    private readonly Control _syncControl;
    private readonly AppSettings _settings;
    private IntPtr _locationChangeHook;
    private IntPtr _destroyHook;
    private IntPtr _hideHook;
    private bool _disposed;

    // Must be stored as fields to prevent GC collection of the delegate
    private readonly NativeMethods.WinEventProc _locationChangeProc;
    private readonly NativeMethods.WinEventProc _destroyProc;
    private readonly NativeMethods.WinEventProc _hideProc;
    private readonly NativeMethods.WinEventProc _moveSizeEndProc;
    private IntPtr _moveSizeEndHook;
    // Track windows that have been maximized but need to wait for resize end
    private readonly HashSet<IntPtr> _pendingMaximize = new();
    private readonly Dictionary<IntPtr, long> _pendingHide = new();
    private readonly System.Windows.Forms.Timer _hideTimer;

    public WindowMonitor(FullScreenManager manager, FullScreenTracker tracker, Control syncControl, AppSettings settings)
    {
        _manager = manager;
        _tracker = tracker;
        _syncControl = syncControl;
        _settings = settings;

        _locationChangeProc = OnLocationChange;
        _destroyProc = OnDestroy;
        _hideProc = OnHide;
        _moveSizeEndProc = OnMoveSizeEnd;

        _hideTimer = new System.Windows.Forms.Timer { Interval = 100 };
        _hideTimer.Tick += (_, _) => ProcessPendingHides();
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

        // EVENT_OBJECT_HIDE fires when a window is hidden (closed to tray, or closing)
        _hideHook = NativeMethods.SetWinEventHook(
            NativeMethods.EVENT_OBJECT_HIDE,
            NativeMethods.EVENT_OBJECT_HIDE,
            IntPtr.Zero, _hideProc,
            0, 0, NativeMethods.WINEVENT_OUTOFCONTEXT);

        if (_locationChangeHook == IntPtr.Zero || _moveSizeEndHook == IntPtr.Zero
            || _destroyHook == IntPtr.Zero || _hideHook == IntPtr.Zero)
        {
            Trace.WriteLine("WindowMonitor: Failed to set one or more WinEvent hooks.");
        }
        else
        {
            Trace.WriteLine("WindowMonitor: Started monitoring.");
        }
    }

    private bool ShouldTriggerVirtualDesktop()
    {
        bool shiftHeld = (NativeMethods.GetAsyncKeyState(NativeMethods.VK_SHIFT) & 0x8000) != 0;
        return _settings.InvertShiftClick ? !shiftHeld : shiftHeld;
    }

    private void OnLocationChange(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint idEventThread, uint dwmsEventTime)
    {
        // Only care about top-level window changes (OBJID_WINDOW)
        if (idObject != NativeMethods.OBJID_WINDOW || idChild != 0) return;

        // If the window is already tracked, check if it is being restored (i.e., no longer maximized)
        if (_tracker.IsTracked(hwnd))
        {
            var placement = NativeMethods.WINDOWPLACEMENT.Default;
            if (NativeMethods.GetWindowPlacement(hwnd, ref placement))
            {
                if (placement.showCmd != NativeMethods.SW_MAXIMIZE)
                {
                    // If left mouse button is down, user is dragging — let Windows handle the resize.
                    // OnMoveSizeEnd will fire after release and trigger restore.
                    bool isDragging = (NativeMethods.GetAsyncKeyState(NativeMethods.VK_LBUTTON) & 0x8000) != 0;
                    if (isDragging) return;

                    bool isMinimized = NativeMethods.IsIconic(hwnd);
                    Trace.WriteLine($"WindowMonitor: Tracked window {hwnd} un-maximized (minimized={isMinimized}).");
                    MarshalToUiThread(() => _manager.Restore(hwnd, keepMinimized: isMinimized));
                    return;
                }
            }
            // Still maximized; let MoveSizeEnd handle pending maximize.
            return;
        }

        // Not tracked yet: check for a new maximize event (including via shortcut)
        var newPlacement = NativeMethods.WINDOWPLACEMENT.Default;
        if (!NativeMethods.GetWindowPlacement(hwnd, ref newPlacement)) return;
        if (newPlacement.showCmd != NativeMethods.SW_MAXIMIZE) return;

        bool triggerVirtualDesktop = ShouldTriggerVirtualDesktop();
        if (triggerVirtualDesktop)
        {
            // Defer maximization until after the resize operation completes
            MarshalToUiThread(() =>
            {
                if (_pendingMaximize.Add(hwnd))
                {
                    Trace.WriteLine($"WindowMonitor: Queued maximize for window {hwnd} after resize end.");
                    // Schedule a fallback in case MoveSizeEnd does not fire (e.g., keyboard shortcut)
                    _ = Task.Run(async () =>
                    {
                        await Task.Delay(50);
                        // Marshal the check/remove back onto the UI thread
                        MarshalToUiThread(() =>
                        {
                            if (_pendingMaximize.Contains(hwnd))
                            {
                                _pendingMaximize.Remove(hwnd);
                                // If already tracked, the hook path already handled it — don't double-trigger
                                if (_tracker.IsTracked(hwnd))
                                {
                                    Trace.WriteLine($"WindowMonitor: Pending window {hwnd} already tracked, skipping fallback.");
                                    return;
                                }
                                var placement = NativeMethods.WINDOWPLACEMENT.Default;
                                bool isMaximized = NativeMethods.GetWindowPlacement(hwnd, ref placement) && placement.showCmd == NativeMethods.SW_MAXIMIZE;
                                if (isMaximized)
                                {
                                    Trace.WriteLine($"WindowMonitor: Fallback processing for pending maximize window {hwnd}.");
                                    _manager.Toggle(hwnd);
                                }
                                else
                                {
                                    Trace.WriteLine($"WindowMonitor: Fallback detected restore for pending window {hwnd}.");
                                    _manager.Restore(hwnd, keepMinimized: NativeMethods.IsIconic(hwnd));
                                }
                            }
                        });
                    });
                }
            });
        }
    }

    private async void OnMoveSizeEnd(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint idEventThread, uint dwmsEventTime)
    {
        // Only care about top-level window changes (OBJID_WINDOW)
        if (idObject != NativeMethods.OBJID_WINDOW || idChild != 0) return;

        // If this window was pending maximize, handle it now
        bool wasPending = false;
        if (!_syncControl.IsDisposed && _syncControl.IsHandleCreated)
        {
            wasPending = (bool)_syncControl.Invoke(new Func<bool>(() => _pendingMaximize.Remove(hwnd)));
        }
        if (wasPending)
        {
            Trace.WriteLine($"WindowMonitor: MoveSizeEnd triggered for pending maximize window {hwnd}.");
            MarshalToUiThread(() => _manager.Toggle(hwnd));
            return;
        }

        // Handle only tracked windows that are being restored (not maximized)
        if (!_tracker.IsTracked(hwnd)) return;
        var placement = NativeMethods.WINDOWPLACEMENT.Default;
        if (!NativeMethods.GetWindowPlacement(hwnd, ref placement)) return;
        Trace.WriteLine($"WindowMonitor: MoveSizeEnd: tracked window {hwnd} showCmd={placement.showCmd}.");
        if (placement.showCmd != NativeMethods.SW_MAXIMIZE)
        {
            Trace.WriteLine($"WindowMonitor: Tracked window {hwnd} un-maximized via move/size, restoring.");
            await Task.Delay(100);
            MarshalToUiThread(() => _manager.Restore(hwnd, keepMinimized: NativeMethods.IsIconic(hwnd)));
        }
    }

    private void OnDestroy(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint idEventThread, uint dwmsEventTime)
    {
        if (idObject != NativeMethods.OBJID_WINDOW || idChild != 0) return;
        if (!_tracker.IsTracked(hwnd)) return;

        Trace.WriteLine($"WindowMonitor: Tracked window {hwnd} destroyed.");
        MarshalToUiThread(() =>
        {
            _pendingHide.Remove(hwnd);
            _manager.HandleWindowDestroyed(hwnd);
        });
    }

    private void OnHide(IntPtr hWinEventHook, uint eventType, IntPtr hwnd,
        int idObject, int idChild, uint idEventThread, uint dwmsEventTime)
    {
        if (idObject != NativeMethods.OBJID_WINDOW || idChild != 0) return;
        if (!_tracker.IsTracked(hwnd)) return;

        MarshalToUiThread(() =>
        {
            _pendingHide[hwnd] = Environment.TickCount64 + 300;
            _hideTimer.Start();
        });
    }

    private void ProcessPendingHides()
    {
        var now = Environment.TickCount64;
        var due = _pendingHide
            .Where(item => item.Value <= now)
            .Select(item => item.Key)
            .ToList();

        foreach (var hwnd in due)
        {
            _pendingHide.Remove(hwnd);
            if (!_tracker.IsTracked(hwnd) || NativeMethods.IsWindowVisible(hwnd)) continue;

            if (NativeMethods.IsWindow(hwnd))
            {
                Trace.WriteLine($"WindowMonitor: Tracked window {hwnd} remained hidden, restoring in place.");
                _manager.Restore(hwnd, keepHidden: true);
                if (_tracker.IsTracked(hwnd))
                    _pendingHide[hwnd] = Environment.TickCount64 + 5000;
            }
            else
            {
                _manager.HandleWindowDestroyed(hwnd);
            }
        }

        if (_pendingHide.Count == 0)
            _hideTimer.Stop();
    }

    private void MarshalToUiThread(Action action)
    {
        if (_disposed || _syncControl.IsDisposed || !_syncControl.IsHandleCreated) return;

        try
        {
            _syncControl.BeginInvoke(action);
        }
        catch (ObjectDisposedException)
        {
            // App is shutting down
        }
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
        if (_hideHook != IntPtr.Zero)
        {
            NativeMethods.UnhookWinEvent(_hideHook);
            _hideHook = IntPtr.Zero;
        }
        if (_moveSizeEndHook != IntPtr.Zero)
        {
            NativeMethods.UnhookWinEvent(_moveSizeEndHook);
            _moveSizeEndHook = IntPtr.Zero;
        }

        _hideTimer.Stop();
        _hideTimer.Dispose();
        _pendingHide.Clear();

        Trace.WriteLine("WindowMonitor: Disposed.");
    }
}
