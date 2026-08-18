using System.Diagnostics;
using System.Runtime.InteropServices;
using MaximizeToVirtualDesktop.Interop;

namespace MaximizeToVirtualDesktop;

/// <summary>
/// Low-level mouse hook that detects Shift+Click on a window's maximize button.
/// Suppresses the Windows maximize and triggers virtual-desktop maximize instead.
/// </summary>
internal sealed class MaximizeButtonHook : IDisposable
{
    private readonly FullScreenManager _manager;
    private readonly Control _syncControl;
    private readonly AppSettings _settings;
    private IntPtr _hookHandle;
    private bool _disposed;

    // Must be stored as a field to prevent GC collection
    private readonly NativeMethods.LowLevelHookProc _hookProc;

    public MaximizeButtonHook(FullScreenManager manager, Control syncControl, AppSettings settings)
    {
        _manager = manager;
        _syncControl = syncControl;
        _settings = settings;
        _hookProc = HookCallback;
    }

    public void Install()
    {
        if (_hookHandle != IntPtr.Zero) return;

        _hookHandle = NativeMethods.SetWindowsHookEx(
            NativeMethods.WH_MOUSE_LL,
            _hookProc,
            NativeMethods.GetModuleHandle(null),
            0);

        if (_hookHandle == IntPtr.Zero)
        {
            Trace.WriteLine("MaximizeButtonHook: Failed to install mouse hook.");
        }
        else
        {
            Trace.WriteLine("MaximizeButtonHook: Installed.");
        }
    }

    private IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= NativeMethods.HC_ACTION && wParam == (IntPtr)NativeMethods.WM_LBUTTONDOWN)
        {
            if (IsTriggerActive())
            {
                var hookStruct = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);
                var hwnd = NativeMethods.WindowFromPoint(hookStruct.pt);
                if (hwnd != IntPtr.Zero && IsClickOnMaximizeButton(hwnd, hookStruct.pt))
                {
                    var topLevel = NativeMethods.GetAncestor(hwnd, NativeMethods.GA_ROOT);
                    if (topLevel != IntPtr.Zero)
                    {
                        PostToggle(topLevel);
                        return (IntPtr)1;
                    }
                }
            }
        }

        return NativeMethods.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
    }

    private bool IsTriggerActive()
    {
        bool shiftHeld = (NativeMethods.GetAsyncKeyState(NativeMethods.VK_SHIFT) & 0x8000) != 0;
        return _settings.InvertShiftClick ? !shiftHeld : shiftHeld;
    }

    private void PostToggle(IntPtr topLevel)
    {
        try
        {
            if (!_syncControl.IsDisposed && _syncControl.IsHandleCreated)
            {
                _syncControl.BeginInvoke(() => _manager.Toggle(topLevel));
            }
        }
        catch (ObjectDisposedException)
        {
            // App is shutting down
        }
    }

    private static bool IsClickOnMaximizeButton(IntPtr hwnd, NativeMethods.POINT pt)
    {
        try
        {
            IntPtr lParam = (IntPtr)((pt.Y << 16) | (pt.X & 0xFFFF));
            IntPtr result = NativeMethods.SendMessageTimeout(
                hwnd, NativeMethods.WM_NCHITTEST, IntPtr.Zero, lParam,
                NativeMethods.SMTO_ABORTIFHUNG, 100, out IntPtr hitResult);
            return result != IntPtr.Zero && hitResult == (IntPtr)NativeMethods.HTMAXBUTTON;
        }
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        if (_hookHandle != IntPtr.Zero)
        {
            NativeMethods.UnhookWindowsHookEx(_hookHandle);
            _hookHandle = IntPtr.Zero;
            Trace.WriteLine("MaximizeButtonHook: Uninstalled.");
        }
    }
}
