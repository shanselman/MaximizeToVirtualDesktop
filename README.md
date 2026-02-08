# MaximizeToVirtualDesktop

Bring macOS's green-button "maximize to full-screen virtual desktop" behavior to Windows 11.

When triggered, the foreground window is moved to a brand-new virtual desktop and maximized. Closing, un-maximizing, or toggling the hotkey again restores everything — the window returns to its original desktop and size, and the temporary desktop is removed.

## Usage

| Trigger | How |
|---------|-----|
| **Hotkey** | `Ctrl + Alt + Shift + X` — toggles the foreground window |
| **Shift + Click** | Hold `Shift` and click a window's maximize button |

The app runs in the system tray. Right-click the tray icon for options:
- **Restore All** — brings all windows back and removes temp desktops
- **Exit** — restores all windows, then exits

## How It Works

1. **Maximize** — creates a new virtual desktop, moves the window there, switches to it, and maximizes the window. The desktop is named after the process.
2. **Auto-restore on un-maximize** — if you restore/un-maximize a tracked window, it's automatically sent back to its original desktop and the temp desktop is cleaned up.
3. **Auto-restore on close** — closing a tracked window triggers the same cleanup.
4. **Toggle** — pressing the hotkey on an already-tracked window restores it.

## Requirements

- **Windows 11 24H2** or later
- .NET 8 Runtime (or use the self-contained publish)

### Shift+Click Compatibility

Shift+Click works wherever Windows 11 Snap Layouts works — apps that correctly handle `WM_NCHITTEST` and return `HTMAXBUTTON`. This includes:

| App | Works? |
|-----|--------|
| Notepad, Explorer, traditional Win32 | ✅ |
| Windows Terminal (WinUI) | ✅ |
| VS Code (Electron 13+) | ✅ |
| Visual Studio | ✅ |
| Apps with fully custom title bars | ❌ (use hotkey) |

The hotkey **always** works, regardless of the window type.

## Building

```bash
dotnet build
```

### Single-file publish

```bash
dotnet publish -c Release -r win-x64 --self-contained -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true
```

## Architecture

```
src/MaximizeToVirtualDesktop/
├── Program.cs                  Entry point, single-instance Mutex
├── TrayApplication.cs          System tray, hotkey, lifecycle owner
├── FullScreenManager.cs        Orchestrator with rollback on every step
├── FullScreenTracker.cs        Thread-safe hwnd → tracking state map
├── VirtualDesktopService.cs    COM wrapper, 4 operations, Explorer restart recovery
├── WindowMonitor.cs            SetWinEventHook for close/un-maximize detection
├── MaximizeButtonHook.cs       WH_MOUSE_LL for Shift+Click on maximize button
└── Interop/
    ├── NativeMethods.cs        P/Invoke declarations
    └── VirtualDesktopCom.cs    Vendored COM interfaces (from MScholtes/VirtualDesktop)
```

**Zero NuGet dependencies.** COM interop declarations are vendored from [MScholtes/VirtualDesktop](https://github.com/MScholtes/VirtualDesktop) (MIT license, actively maintained).

## Design Principles

1. **Reliable over featureful** — every code path handles failure. No crashes, no orphaned desktops, no stuck windows. If something goes wrong, roll back to the state before we touched anything.
2. **Tight** — one project, zero packages, 9 files. Each file has one job.
3. **Clean** — `IDisposable` on every native resource. No fire-and-forget. All WinEvent callbacks marshal to the UI thread.

## Known Limitations

- **Windows version dependency** — the undocumented `IVirtualDesktopManagerInternal` COM interface changes GUIDs with major Windows updates (~2-3x/year). When this happens, update `Interop/VirtualDesktopCom.cs` from [MScholtes' latest](https://github.com/MScholtes/VirtualDesktop).
- **Elevated windows** — cannot move windows running as Administrator from a non-elevated instance.
- **App crash** — if the app crashes, temporary desktops may remain. They're named after the process for easy identification.

## Prior Art

- [Peach](https://peachapp.net) — MS Store app with the same hotkey UX
- [MScholtes/VirtualDesktop](https://github.com/MScholtes/VirtualDesktop) — the COM interface source we vendor
- PowerToys feature requests [#13993](https://github.com/microsoft/PowerToys/issues/13993), [#21597](https://github.com/microsoft/PowerToys/issues/21597)

## License

MIT

---

*Credit to Kieran Mockford for the original idea. May he rest in peace.*
