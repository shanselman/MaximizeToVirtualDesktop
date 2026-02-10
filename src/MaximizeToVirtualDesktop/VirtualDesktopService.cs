using System.Diagnostics;
using System.Runtime.InteropServices;
using MaximizeToVirtualDesktop.Interop;

namespace MaximizeToVirtualDesktop;

/// <summary>
/// Wraps the COM virtual desktop APIs. All methods are defensive — they catch COM
/// exceptions and return success/failure rather than throwing.
/// </summary>
internal sealed class VirtualDesktopService : IDisposable
{
    private DesktopManagerAdapter? _managerInternal;
    private IVirtualDesktopManager? _manager;
    private IApplicationViewCollection? _viewCollection;
    private IVirtualDesktopPinnedApps? _pinnedApps;
    private int _buildNumber;
    private bool _disposed;

    public bool IsInitialized => _managerInternal != null && _manager != null;

    public bool Initialize(int windowsBuildNumber)
    {
        _buildNumber = windowsBuildNumber;
        IServiceProvider10? shell = null;
        try
        {
            shell = (IServiceProvider10)Activator.CreateInstance(
                Type.GetTypeFromCLSID(ComGuids.CLSID_ImmersiveShell)!)!;

            var mgrInternalGuid = ComGuids.IID_VirtualDesktopManagerInternal;
            var mgrInternalRaw = shell.QueryService(
                ref Unsafe.AsRef(ComGuids.CLSID_VirtualDesktopManagerInternal), ref mgrInternalGuid);

            _managerInternal = DesktopManagerAdapter.Create(mgrInternalRaw, windowsBuildNumber);

            _manager = (IVirtualDesktopManager)Activator.CreateInstance(
                Type.GetTypeFromCLSID(ComGuids.CLSID_VirtualDesktopManager)!)!;

            var viewCollGuid = typeof(IApplicationViewCollection).GUID;
            _viewCollection = (IApplicationViewCollection)shell.QueryService(
                ref viewCollGuid, ref viewCollGuid);

            // Pin support — query IVirtualDesktopPinnedApps
            try
            {
                var pinnedGuid = typeof(IVirtualDesktopPinnedApps).GUID;
                _pinnedApps = (IVirtualDesktopPinnedApps)shell.QueryService(
                    ref Unsafe.AsRef(ComGuids.CLSID_VirtualDesktopPinnedApps), ref pinnedGuid);
            }
            catch
            {
                Trace.WriteLine("VirtualDesktopService: IVirtualDesktopPinnedApps not available.");
            }

            Trace.WriteLine("VirtualDesktopService: COM initialized successfully.");
            return true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: COM initialization failed: {ex.Message}");
            ReleaseComObjects();
            return false;
        }
        finally
        {
            if (shell != null) Marshal.ReleaseComObject(shell);
        }
    }

    /// <summary>
    /// Reinitialize COM objects (e.g. after Explorer restart).
    /// </summary>
    public bool Reinitialize()
    {
        ReleaseComObjects();
        return Initialize(_buildNumber);
    }

    public Guid? GetCurrentDesktopId()
    {
        IVirtualDesktop? desktop = null;
        try
        {
            desktop = _managerInternal?.GetCurrentDesktop();
            return desktop?.GetId();
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: GetCurrentDesktopId failed: {ex.Message}");
            return null;
        }
        finally
        {
            if (desktop != null) Marshal.ReleaseComObject(desktop);
        }
    }

    public Guid? GetDesktopIdForWindow(IntPtr hwnd)
    {
        try
        {
            return _manager?.GetWindowDesktopId(hwnd);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: GetDesktopIdForWindow failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Creates a new virtual desktop. Returns its ID, or null on failure.
    /// </summary>
    public (IVirtualDesktop? desktop, Guid? id) CreateDesktop()
    {
        try
        {
            var desktop = _managerInternal!.CreateDesktop();
            var id = desktop.GetId();
            Trace.WriteLine($"VirtualDesktopService: Created desktop {id}");
            return (desktop, id);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: CreateDesktop failed: {ex.Message}");
            return (null, null);
        }
    }

    /// <summary>
    /// Moves a window to the specified desktop. Uses MoveViewToDesktop for cross-process windows.
    /// </summary>
    public bool MoveWindowToDesktop(IntPtr hwnd, IVirtualDesktop desktop)
    {
        IApplicationView? view = null;
        NativeMethods.GetWindowThreadProcessId(hwnd, out int processId);
        string processName = string.Empty;
        
        // Get process name for diagnostics
        if (processId > 0)
        {
            try
            {
                using var process = Process.GetProcessById(processId);
                processName = process.ProcessName; // Keep original casing
            }
            catch (ArgumentException)
            {
                // Process has exited
                Trace.WriteLine($"VirtualDesktopService: Process {processId} has exited");
            }
            catch (InvalidOperationException ex)
            {
                // Process has no associated process (e.g., system process)
                Trace.WriteLine($"VirtualDesktopService: Cannot get process info for {processId}: {ex.Message}");
            }
        }
        else
        {
            Trace.WriteLine($"VirtualDesktopService: Invalid process ID for hwnd {hwnd}");
        }

        try
        {
            // Attempt 1: Use IApplicationView (standard cross-process method)
            int hr = _viewCollection?.GetViewForHwnd(hwnd, out view) ?? unchecked((int)0x80004003); // E_POINTER if null
            
            if (view != null)
            {
                _managerInternal!.MoveViewToDesktop(view, desktop);
                Trace.WriteLine($"VirtualDesktopService: Moved window {hwnd} ({processName}) to desktop {desktop.GetId()}");
                return true;
            }
            
            // Attempt 2: Try documented API as fallback
            // This typically only works for windows owned by this process
            try
            {
                var desktopId = desktop.GetId();
                _manager!.MoveWindowToDesktop(hwnd, ref desktopId);
                Trace.WriteLine($"VirtualDesktopService: Moved window {hwnd} ({processName}) via IVirtualDesktopManager to desktop {desktopId}");
                return true;
            }
            catch (COMException comEx) when (comEx.HResult == unchecked((int)0x80070005)) // E_ACCESSDENIED
            {
                // Expected for windows from other processes
                Trace.WriteLine($"VirtualDesktopService: IVirtualDesktopManager access denied (expected for cross-process), hr={hr:X8}");
            }
            
            Trace.WriteLine($"VirtualDesktopService: GetViewForHwnd returned view=null for hwnd={hwnd}, process={processName}, hr={hr:X8}");
            return false;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: MoveWindowToDesktop failed for {processName} hwnd={hwnd}: {ex.Message}");
            return false;
        }
        finally
        {
            if (view != null) Marshal.ReleaseComObject(view);
        }
    }

    public bool SwitchToDesktop(IVirtualDesktop desktop)
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

            _managerInternal!.SwitchDesktopWithAnimation(desktop);

            Trace.WriteLine($"VirtualDesktopService: Switched to desktop {desktop.GetId()}");
            return true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: SwitchToDesktop failed: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Removes a virtual desktop. Windows on it move to the fallback desktop.
    /// </summary>
    public bool RemoveDesktop(IVirtualDesktop desktop)
    {
        IVirtualDesktop? current = null;
        IVirtualDesktop? adjacent = null;
        try
        {
            // Find a fallback desktop (the current one, or adjacent)
            current = _managerInternal!.GetCurrentDesktop();
            IVirtualDesktop fallback;

            if (current.GetId() == desktop.GetId())
            {
                // We're removing the current desktop — find an adjacent one
                int hr = _managerInternal.GetAdjacentDesktop(desktop, 3, out fallback); // 3 = Left
                if (hr != 0)
                {
                    hr = _managerInternal.GetAdjacentDesktop(desktop, 4, out fallback); // 4 = Right
                    if (hr != 0)
                    {
                        Trace.WriteLine("VirtualDesktopService: No adjacent desktop for fallback, cannot remove.");
                        return false;
                    }
                }
                adjacent = fallback;
            }
            else
            {
                fallback = current;
            }

            _managerInternal.RemoveDesktop(desktop, fallback);
            Trace.WriteLine($"VirtualDesktopService: Removed desktop {desktop.GetId()}");
            return true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: RemoveDesktop failed: {ex.Message}");
            return false;
        }
        finally
        {
            if (adjacent != null) Marshal.ReleaseComObject(adjacent);
            if (current != null) Marshal.ReleaseComObject(current);
        }
    }

    public bool SetDesktopName(IVirtualDesktop desktop, string name)
    {
        IntPtr hstring = IntPtr.Zero;
        try
        {
            int hr = NativeMethods.WindowsCreateString(name, name.Length, out hstring);
            if (hr != 0) return false;

            _managerInternal!.SetDesktopName(desktop, hstring);
            return true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: SetDesktopName failed: {ex.Message}");
            return false;
        }
        finally
        {
            if (hstring != IntPtr.Zero) NativeMethods.WindowsDeleteString(hstring);
        }
    }

    public IVirtualDesktop? FindDesktop(Guid desktopId)
    {
        try
        {
            return _managerInternal?.FindDesktop(ref desktopId);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: FindDesktop failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Returns true if the window's view is pinned to all virtual desktops.
    /// </summary>
    public bool IsWindowPinned(IntPtr hwnd)
    {
        IApplicationView? view = null;
        try
        {
            if (_pinnedApps == null || _viewCollection == null) return false;
            _viewCollection.GetViewForHwnd(hwnd, out view);
            return view != null && _pinnedApps.IsViewPinned(view);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: IsWindowPinned failed: {ex.Message}");
            return false;
        }
        finally
        {
            if (view != null) Marshal.ReleaseComObject(view);
        }
    }

    /// <summary>
    /// Pin a window's view to all virtual desktops.
    /// </summary>
    public bool PinWindow(IntPtr hwnd)
    {
        IApplicationView? view = null;
        try
        {
            if (_pinnedApps == null || _viewCollection == null) return false;
            _viewCollection.GetViewForHwnd(hwnd, out view);
            if (view == null) return false;
            _pinnedApps.PinView(view);
            Trace.WriteLine($"VirtualDesktopService: Pinned window {hwnd}");
            return true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: PinWindow failed: {ex.Message}");
            return false;
        }
        finally
        {
            if (view != null) Marshal.ReleaseComObject(view);
        }
    }

    /// <summary>
    /// Unpin a window's view from all virtual desktops.
    /// </summary>
    public bool UnpinWindow(IntPtr hwnd)
    {
        IApplicationView? view = null;
        try
        {
            if (_pinnedApps == null || _viewCollection == null) return false;
            _viewCollection.GetViewForHwnd(hwnd, out view);
            if (view == null) return false;
            _pinnedApps.UnpinView(view);
            Trace.WriteLine($"VirtualDesktopService: Unpinned window {hwnd}");
            return true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"VirtualDesktopService: UnpinWindow failed: {ex.Message}");
            return false;
        }
        finally
        {
            if (view != null) Marshal.ReleaseComObject(view);
        }
    }

    private void ReleaseComObjects()
    {
        if (_managerInternal != null) { _managerInternal.Dispose(); _managerInternal = null; }
        if (_manager != null) { Marshal.ReleaseComObject(_manager); _manager = null; }
        if (_viewCollection != null) { Marshal.ReleaseComObject(_viewCollection); _viewCollection = null; }
        if (_pinnedApps != null) { Marshal.ReleaseComObject(_pinnedApps); _pinnedApps = null; }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            ReleaseComObjects();
            _disposed = true;
        }
    }
}

// Helper to pass readonly Guid fields by ref to COM
internal static class Unsafe
{
    internal static ref Guid AsRef(in Guid guid)
    {
        // We need to pass a readonly static Guid by ref to COM QueryService.
        // This is safe because COM only reads the value.
        unsafe
        {
            fixed (Guid* p = &guid)
            {
                return ref *p;
            }
        }
    }
}
