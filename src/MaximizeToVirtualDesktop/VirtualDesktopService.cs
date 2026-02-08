using System.Diagnostics;
using MaximizeToVirtualDesktop.Interop;
using WindowsDesktop;

namespace MaximizeToVirtualDesktop;

/// <summary>
/// Wraps the Jack251970/VirtualDesktop library (Slions.VirtualDesktop NuGet package).
/// This library handles COM interface versioning automatically — GUIDs and vtable layouts
/// are resolved at runtime based on the Windows build, eliminating the need to manually
/// update vendored COM declarations when Windows updates break things.
///
/// All methods are defensive — they catch exceptions and return success/failure rather
/// than throwing.
/// </summary>
internal sealed class VirtualDesktopService : IDisposable
{
    private bool _disposed;

    public bool IsInitialized { get; private set; }

    public bool Initialize()
    {
        try
        {
            // The library auto-detects the Windows build and selects the correct
            // COM interface GUIDs and vtable layouts at initialization time.
            // It uses runtime compilation via Roslyn to generate a version-specific
            // interop assembly, cached on disk for subsequent launches.
            var desktops = VirtualDesktop.GetDesktops();
            Trace.WriteLine($"VirtualDesktopService: Initialized successfully ({desktops.Length} desktops).");
            IsInitialized = true;
            return true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: Initialization failed: {ex.Message}");
            IsInitialized = false;
            return false;
        }
    }

    /// <summary>
    /// Reinitialize after Explorer restart or COM failure.
    /// </summary>
    public bool Reinitialize()
    {
        IsInitialized = false;
        return Initialize();
    }

    public Guid? GetCurrentDesktopId()
    {
        try
        {
            return VirtualDesktop.Current.Id;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: GetCurrentDesktopId failed: {ex.Message}");
            return null;
        }
    }

    public Guid? GetDesktopIdForWindow(IntPtr hwnd)
    {
        try
        {
            var desktop = VirtualDesktop.FromHwnd(hwnd);
            return desktop?.Id;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: GetDesktopIdForWindow failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Creates a new virtual desktop. Returns the desktop and its ID, or null on failure.
    /// </summary>
    public (VirtualDesktop? desktop, Guid? id) CreateDesktop()
    {
        try
        {
            var desktop = VirtualDesktop.Create();
            Trace.WriteLine($"VirtualDesktopService: Created desktop {desktop.Id}");
            return (desktop, desktop.Id);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: CreateDesktop failed: {ex.Message}");
            return (null, null);
        }
    }

    /// <summary>
    /// Moves a window to the specified desktop. The library handles cross-process
    /// windows internally (using MoveViewToDesktop for other-process windows).
    /// </summary>
    public bool MoveWindowToDesktop(IntPtr hwnd, VirtualDesktop desktop)
    {
        try
        {
            VirtualDesktop.MoveToDesktop(hwnd, desktop);
            Trace.WriteLine($"VirtualDesktopService: Moved window {hwnd} to desktop {desktop.Id}");
            return true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: MoveWindowToDesktop failed: {ex.Message}");

            // Second attempt: try main window of the process
            try
            {
                NativeMethods.GetWindowThreadProcessId(hwnd, out int processId);
                using var process = Process.GetProcessById(processId);
                if (process.MainWindowHandle != IntPtr.Zero && process.MainWindowHandle != hwnd)
                {
                    VirtualDesktop.MoveToDesktop(process.MainWindowHandle, desktop);
                    Trace.WriteLine($"VirtualDesktopService: Moved main window instead for process {processId}");
                    return true;
                }
            }
            catch (Exception ex2)
            {
                Trace.WriteLine($"VirtualDesktopService: Fallback move also failed: {ex2.Message}");
            }

            return false;
        }
    }

    public bool SwitchToDesktop(VirtualDesktop desktop)
    {
        try
        {
            // Activate the taskbar to prevent flashing icons (from MScholtes)
            var taskbarHwnd = NativeMethods.FindWindow("Shell_TrayWnd", null);
            if (taskbarHwnd != IntPtr.Zero)
            {
                NativeMethods.GetWindowThreadProcessId(taskbarHwnd, out _);
                var foregroundHwnd = NativeMethods.GetForegroundWindow();
                uint desktopThreadId = NativeMethods.GetWindowThreadProcessId(taskbarHwnd, out _);
                uint foregroundThreadId = NativeMethods.GetWindowThreadProcessId(foregroundHwnd, out _);
                uint currentThreadId = NativeMethods.GetCurrentThreadId();

                if (desktopThreadId != 0 && foregroundThreadId != 0 && foregroundThreadId != currentThreadId)
                {
                    NativeMethods.AttachThreadInput(desktopThreadId, currentThreadId, true);
                    NativeMethods.AttachThreadInput(foregroundThreadId, currentThreadId, true);
                    NativeMethods.SetForegroundWindow(taskbarHwnd);
                    NativeMethods.AttachThreadInput(foregroundThreadId, currentThreadId, false);
                    NativeMethods.AttachThreadInput(desktopThreadId, currentThreadId, false);
                }
            }

            desktop.Switch();

            Trace.WriteLine($"VirtualDesktopService: Switched to desktop {desktop.Id}");
            return true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: SwitchToDesktop failed: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Removes a virtual desktop. Windows on it move to an adjacent desktop.
    /// </summary>
    public bool RemoveDesktop(VirtualDesktop desktop)
    {
        try
        {
            desktop.Remove();
            Trace.WriteLine($"VirtualDesktopService: Removed desktop {desktop.Id}");
            return true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: RemoveDesktop failed: {ex.Message}");
            return false;
        }
    }

    public bool SetDesktopName(VirtualDesktop desktop, string name)
    {
        try
        {
            desktop.Name = name;
            return true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: SetDesktopName failed: {ex.Message}");
            return false;
        }
    }

    public VirtualDesktop? FindDesktop(Guid desktopId)
    {
        try
        {
            return VirtualDesktop.FromId(desktopId);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: FindDesktop failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Returns true if the window is pinned to all virtual desktops.
    /// </summary>
    public bool IsWindowPinned(IntPtr hwnd)
    {
        try
        {
            return VirtualDesktop.IsPinnedWindow(hwnd);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: IsWindowPinned failed: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Pin a window to all virtual desktops.
    /// </summary>
    public bool PinWindow(IntPtr hwnd)
    {
        try
        {
            var result = VirtualDesktop.PinWindow(hwnd);
            if (result) Trace.WriteLine($"VirtualDesktopService: Pinned window {hwnd}");
            return result;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: PinWindow failed: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Unpin a window from all virtual desktops.
    /// </summary>
    public bool UnpinWindow(IntPtr hwnd)
    {
        try
        {
            var result = VirtualDesktop.UnpinWindow(hwnd);
            if (result) Trace.WriteLine($"VirtualDesktopService: Unpinned window {hwnd}");
            return result;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: UnpinWindow failed: {ex.Message}");
            return false;
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            // The library manages its own COM lifecycle internally.
            // No manual COM object release needed.
            _disposed = true;
        }
    }
}
