using System.Diagnostics;
using System.Runtime.InteropServices;
using MaximizeToVirtualDesktop.Interop;

namespace MaximizeToVirtualDesktop;

/// <summary>
/// Low-level mouse hook that detects Shift+Click on a window's maximize button.
/// Works wherever Windows 11 Snap Layouts works (apps that return HTMAXBUTTON from WM_NCHITTEST).
/// </summary>
internal sealed class MaximizeButtonHook : IDisposable
{
    private readonly FullScreenManager _manager;
    private IntPtr _hookHandle;
    private bool _disposed;

    // Must be stored as a field to prevent GC collection
    private readonly NativeMethods.LowLevelHookProc _hookProc;

    public MaximizeButtonHook(FullScreenManager manager)
    {
        _manager = manager;
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
            // Is Shift held?
            if ((NativeMethods.GetAsyncKeyState(NativeMethods.VK_SHIFT) & 0x8000) != 0)
            {
                var hookStruct = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);
                var hwnd = NativeMethods.WindowFromPoint(hookStruct.pt);

                if (hwnd != IntPtr.Zero && IsClickOnMaximizeButton(hwnd, hookStruct.pt))
                {
                    Trace.WriteLine($"MaximizeButtonHook: Shift+Click on maximize button of {hwnd}");

                    // Find the top-level window (the click target might be a child)
                    var topLevel = GetTopLevelWindow(hwnd);
                    if (topLevel != IntPtr.Zero)
                    {
                        _manager.Toggle(topLevel);

                        // Suppress the click — return non-zero to prevent it reaching the window
                        return (IntPtr)1;
                    }
                }
            }
        }

        return NativeMethods.CallNextHookEx(_hookHandle, nCode, wParam, lParam);
    }

    private static bool IsClickOnMaximizeButton(IntPtr hwnd, NativeMethods.POINT pt)
    {
        try
        {
            // Send WM_NCHITTEST to determine what part of the window the click is on.
            // This is the same mechanism Windows 11 uses for Snap Layouts.
            IntPtr lParam = (IntPtr)((pt.Y << 16) | (pt.X & 0xFFFF));
            IntPtr hitResult = NativeMethods.SendMessage(hwnd, NativeMethods.WM_NCHITTEST, IntPtr.Zero, lParam);
            return hitResult == (IntPtr)NativeMethods.HTMAXBUTTON;
        }
        catch
        {
            return false;
        }
    }

    private static IntPtr GetTopLevelWindow(IntPtr hwnd)
    {
        // Walk up the parent chain to find the top-level window
        IntPtr current = hwnd;
        IntPtr parent;
        while ((parent = GetParent(current)) != IntPtr.Zero)
        {
            current = parent;
        }
        return current;
    }

    [DllImport("user32.dll")]
    private static extern IntPtr GetParent(IntPtr hWnd);

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
