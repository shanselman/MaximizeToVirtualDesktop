// Virtual Desktop COM interface declarations for Windows 11 24H2+
// Vendored from MScholtes/VirtualDesktop v1.21 (MIT License)
// https://github.com/MScholtes/VirtualDesktop
//
// These GUIDs are UNDOCUMENTED and change with Windows builds.
// When they break, update from MScholtes' latest VirtualDesktop11-24H2.cs.
// Last verified: Windows 11 24H2 (build 26100), August 2025.

using System.Runtime.InteropServices;

namespace MaximizeToVirtualDesktop.Interop;

internal static class ComGuids
{
    public static readonly Guid CLSID_ImmersiveShell = new("C2F03A33-21F5-47FA-B4BB-156362A2F239");
    public static readonly Guid CLSID_VirtualDesktopManagerInternal = new("C5E0CDCA-7B6E-41B2-9FC4-D93975CC467B");
    public static readonly Guid CLSID_VirtualDesktopManager = new("AA509086-5CA9-4C25-8F95-589D3C07B48A");
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
internal interface IServiceProvider10
{
    [return: MarshalAs(UnmanagedType.IUnknown)]
    object QueryService(ref Guid service, ref Guid riid);
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("92CA9DCD-5622-4BBA-A805-5E9F541BD8C9")]
internal interface IObjectArray
{
    void GetCount(out int count);
    void GetAt(int index, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object obj);
}

// Stable, documented interface — GUID has not changed since Windows 10.
[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("A5CD92FF-29BE-454C-8D04-D82879FB3F1B")]
internal interface IVirtualDesktopManager
{
    bool IsWindowOnCurrentVirtualDesktop(IntPtr topLevelWindow);
    Guid GetWindowDesktopId(IntPtr topLevelWindow);
    void MoveWindowToDesktop(IntPtr topLevelWindow, ref Guid desktopId);
}

// Undocumented — GUID changes with Windows builds.
[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("3F07F4BE-B107-441A-AF0F-39D82529072C")]
internal interface IVirtualDesktop
{
    bool IsViewVisible(IApplicationView view);
    Guid GetId();
    [return: MarshalAs(UnmanagedType.HString)]
    string GetName();
    [return: MarshalAs(UnmanagedType.HString)]
    string GetWallpaperPath();
    bool IsRemote();
}

// Undocumented — GUID changes with Windows builds.
[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("53F5CA0B-158F-4124-900C-057158060B27")]
internal interface IVirtualDesktopManagerInternal
{
    int GetCount();
    void MoveViewToDesktop(IApplicationView view, IVirtualDesktop desktop);
    bool CanViewMoveDesktops(IApplicationView view);
    IVirtualDesktop GetCurrentDesktop();
    void GetDesktops(out IObjectArray desktops);
    [PreserveSig]
    int GetAdjacentDesktop(IVirtualDesktop from, int direction, out IVirtualDesktop desktop);
    void SwitchDesktop(IVirtualDesktop desktop);
    void SwitchDesktopAndMoveForegroundView(IVirtualDesktop desktop);
    IVirtualDesktop CreateDesktop();
    void MoveDesktop(IVirtualDesktop desktop, int nIndex);
    void RemoveDesktop(IVirtualDesktop desktop, IVirtualDesktop fallback);
    IVirtualDesktop FindDesktop(ref Guid desktopId);
    void GetDesktopSwitchIncludeExcludeViews(IVirtualDesktop desktop, out IObjectArray unknown1, out IObjectArray unknown2);
    void SetDesktopName(IVirtualDesktop desktop, [MarshalAs(UnmanagedType.HString)] string name);
    void SetDesktopWallpaper(IVirtualDesktop desktop, [MarshalAs(UnmanagedType.HString)] string path);
    void UpdateWallpaperPathForAllDesktops([MarshalAs(UnmanagedType.HString)] string path);
    void CopyDesktopState(IApplicationView pView0, IApplicationView pView1);
    void CreateRemoteDesktop([MarshalAs(UnmanagedType.HString)] string path, out IVirtualDesktop desktop);
    void SwitchRemoteDesktop(IVirtualDesktop desktop, IntPtr switchtype);
    void SwitchDesktopWithAnimation(IVirtualDesktop desktop);
    void GetLastActiveDesktop(out IVirtualDesktop desktop);
    void WaitForAnimationToComplete();
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIInspectable)]
[Guid("372E1D3B-38D3-42E4-A15B-8AB2B178F513")]
internal interface IApplicationView
{
    int SetFocus();
    int SwitchTo();
    int TryInvokeBack(IntPtr callback);
    int GetThumbnailWindow(out IntPtr hwnd);
    int GetMonitor(out IntPtr immersiveMonitor);
    int GetVisibility(out int visibility);
    int SetCloak(int cloakType, int unknown);
    int GetPosition(ref Guid guid, out IntPtr position);
    int SetPosition(ref IntPtr position);
    int InsertAfterWindow(IntPtr hwnd);
    int GetExtendedFramePosition(out Rect rect);
    int GetAppUserModelId([MarshalAs(UnmanagedType.LPWStr)] out string id);
    int SetAppUserModelId(string id);
    int IsEqualByAppUserModelId(string id, out int result);
    int GetViewState(out uint state);
    int SetViewState(uint state);
    int GetNeediness(out int neediness);
    int GetLastActivationTimestamp(out ulong timestamp);
    int SetLastActivationTimestamp(ulong timestamp);
    int GetVirtualDesktopId(out Guid guid);
    int SetVirtualDesktopId(ref Guid guid);
    int GetShowInSwitchers(out int flag);
    int SetShowInSwitchers(int flag);
    int GetScaleFactor(out int factor);
    int CanReceiveInput(out bool canReceiveInput);
    int GetCompatibilityPolicyType(out int flags);
    int SetCompatibilityPolicyType(int flags);
    // remaining methods omitted — we don't call them
}

[StructLayout(LayoutKind.Sequential)]
internal struct Rect
{
    public int Left, Top, Right, Bottom;
}

[ComImport]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
[Guid("1841C6D7-4F9D-42C0-AF41-8747538F10E5")]
internal interface IApplicationViewCollection
{
    int GetViews(out IObjectArray array);
    int GetViewsByZOrder(out IObjectArray array);
    int GetViewsByAppUserModelId(string id, out IObjectArray array);
    int GetViewForHwnd(IntPtr hwnd, out IApplicationView view);
    int GetViewForApplication(object application, out IApplicationView view);
    int GetViewForAppUserModelId(string id, out IApplicationView view);
    int GetViewInFocus(out IntPtr view);
    int Unknown1(out IntPtr view);
    void RefreshCollection();
    int RegisterForApplicationViewChanges(object listener, out int cookie);
    int UnregisterForApplicationViewChanges(int cookie);
}
