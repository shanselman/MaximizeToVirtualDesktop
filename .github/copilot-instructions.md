# MaximizeToVirtualDesktop — Copilot Instructions

## Project Overview
A Windows 11 system tray app that replicates macOS's green-button behavior: maximize a window to its own virtual desktop, with automatic cleanup on restore/close. C# / .NET 8 WinForms, zero NuGet dependencies.

## Architecture
- **8 source files** in `src/MaximizeToVirtualDesktop/`, flat structure
- **TrayApplication** owns the lifecycle of all components
- **FullScreenManager** is the orchestrator — never bypass it to call VirtualDesktopService directly
- **FullScreenTracker** is the single source of truth for tracked windows
- **VirtualDesktopService** wraps COM; every method catches exceptions and returns success/failure
- **WindowMonitor** uses SetWinEventHook (out-of-process, no DLL injection)
- **MaximizeButtonHook** uses WH_MOUSE_LL (in-process, no DLL injection)
- **Interop/** contains P/Invoke declarations and vendored COM interfaces

## Key Rules
1. **Zero NuGet dependencies.** COM interop is vendored from MScholtes/VirtualDesktop. Do not add packages.
2. **Every mutating step needs rollback.** If step N+1 fails, undo step N. See MaximizeToDesktop() for the pattern.
3. **All WinEvent/hook callbacks marshal to UI thread** via Control.BeginInvoke. Never mutate FullScreenTracker from a callback thread.
4. **Defensive COM calls.** Every VirtualDesktopService method wraps calls in try/catch. COM objects can become invalid at any time (Explorer restart, desktop removal).
5. **IDisposable on every native resource.** Hooks, COM objects, timers — all cleaned up in OnFormClosing.

## COM Interface Maintenance
`Interop/VirtualDesktopCom.cs` contains undocumented Windows COM GUIDs that change with major Windows updates. When the app breaks after a Windows update:
1. Check [MScholtes/VirtualDesktop](https://github.com/MScholtes/VirtualDesktop) for updated GUIDs
2. Update only `VirtualDesktopCom.cs` — the rest of the code is stable
3. The key interfaces are: `IVirtualDesktop`, `IVirtualDesktopManagerInternal`, `IApplicationView`, `IApplicationViewCollection`

## Testing
- Build: `dotnet build`
- The app must be tested manually (system tray + virtual desktop interactions)
- Test matrix: hotkey toggle, Shift+Click, window close cleanup, un-maximize cleanup, Explorer restart, app exit cleanup

## Style
- Minimal comments — only where the "why" isn't obvious
- `Trace.WriteLine` for diagnostic output (visible in debugger output window)
- No `async/await` — everything runs on the UI thread's message pump
