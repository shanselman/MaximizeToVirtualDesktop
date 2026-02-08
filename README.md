# MaximizeToVirtualDesktop

Bring macOS's green-button "maximize to full-screen virtual desktop" behavior to Windows 11.

When triggered, the foreground window is moved to a brand-new virtual desktop and maximized. Closing, un-maximizing, or toggling the hotkey again restores everything — the window returns to its original desktop and size, and the temporary desktop is removed.

![MaximizeToVirtualDesktop demo](img/maximizetovd_compressed.gif)

## Usage

| Trigger | How |
|---------|-----|
| **Hotkey** | `Ctrl + Alt + Shift + X` — toggles the foreground window |
| **Shift + Click** | Hold `Shift` and click a window's maximize button |

The app runs in the system tray. Right-click the tray icon for options:
- **Restore All** — brings all windows back and removes temp desktops
- **How to Use** — shows usage instructions
- **Check for Updates** — checks for new releases on GitHub
- **Exit** — restores all windows, then exits

## How It Works

1. **Maximize** — creates a new virtual desktop, moves the window there, switches to it, and maximizes the window. The desktop is named `[MVD] ProcessName` so you can identify it in Task View.
2. **Auto-restore on un-maximize** — if you restore/un-maximize a tracked window, it's automatically sent back to its original desktop and the temp desktop is cleaned up.
3. **Auto-restore on close** — closing a tracked window triggers the same cleanup.
4. **Toggle** — pressing the hotkey on an already-tracked window restores it.
5. **Crash recovery** — if the app is killed or crashes, orphaned desktops are automatically cleaned up on next launch.

## Requirements

- **Windows 10** (build 19041/20H1 or later) or **Windows 11** (any version)
- Self-contained — no .NET installation required

The app uses the [Slions.VirtualDesktop](https://github.com/Jack251970/VirtualDesktop) library, which auto-detects your Windows build and selects the correct COM interface GUIDs and vtable layouts at runtime. If a future Windows update breaks things, the app enters degraded mode and checks for updates automatically.

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
├── VirtualDesktopService.cs    Wrapper over Slions.VirtualDesktop library
├── WindowMonitor.cs            SetWinEventHook for close/un-maximize detection
├── MaximizeButtonHook.cs       WH_MOUSE_LL for Shift+Click on maximize button
├── TrackerPersistence.cs       JSON crash recovery in %LOCALAPPDATA%
└── Interop/
    └── NativeMethods.cs        P/Invoke declarations
```

Virtual desktop COM interop is handled by the [Slions.VirtualDesktop](https://github.com/Jack251970/VirtualDesktop) NuGet package (forked from [Grabacr07/VirtualDesktop](https://github.com/Grabacr07/VirtualDesktop), MIT license). This library solves the "moving APIs" problem by:

1. **Data-driven GUIDs** — COM interface GUIDs for each Windows build are stored in `app.config`, not hardcoded in C# interface declarations.
2. **Runtime compilation** — at first launch, the library uses Roslyn to compile a version-specific interop DLL matching your Windows build, then caches it on disk.
3. **Registry fallback** — if GUIDs aren't in the config for your exact build, it looks them up in the Windows Registry.

This eliminates the need to manually update vendored COM declarations when Windows updates change the undocumented interface GUIDs.

## Design Principles

1. **Reliable over featureful** — every code path handles failure. No crashes, no orphaned desktops, no stuck windows. If something goes wrong, roll back to the state before we touched anything.
2. **Tight** — one project, minimal packages, 8 files. Each file has one job.
3. **Clean** — `IDisposable` on every native resource. No fire-and-forget. All WinEvent callbacks marshal to the UI thread.

## Known Limitations

- **Elevated windows** — cannot move windows running as Administrator from a non-elevated instance.
- **App crash** — if the app crashes, temporary desktops are cleaned up automatically on next launch. They're prefixed with `[MVD]` in Task View for easy manual identification.

## The Virtual Desktop GUID Problem

Microsoft's Virtual Desktop feature has a proper, documented COM interface — `IVirtualDesktopManager` — but it can only tell you *which* desktop a window is on and move a window *you own* between desktops. The actually useful operations — creating desktops, switching desktops, moving *any* window, naming desktops — all live behind **undocumented COM interfaces** like `IVirtualDesktopManagerInternal` and `IVirtualDesktop`.

The problem? **Microsoft changes the interface GUIDs with nearly every major Windows update.** Not the methods. Not the signatures. Just the GUIDs. This means every app that uses virtual desktop automation — this one, [Peach](https://peachapp.net), [FancyWM](https://github.com/FancyWM/fancywm), and dozens of others — breaks silently 2-3 times a year and has to scramble to update hardcoded GUIDs.

### How this app handles it

This app uses the [Slions.VirtualDesktop](https://github.com/Jack251970/VirtualDesktop) library (forked from Grabacr07/VirtualDesktop) which solves the moving GUIDs problem through:

1. **Data-driven GUID database** — GUIDs for every known Windows build are stored in `app.config`, keyed by build number. The library ships with GUIDs for Windows 10 (17134+) through Windows 11 24H2+.
2. **Runtime compilation** — at first launch, Roslyn compiles a version-specific interop DLL with the correct GUIDs for your exact Windows build. This DLL is cached on disk for subsequent launches.
3. **Registry fallback** — if your build isn't in the database, the library can discover GUIDs from the Windows Registry (`HKCR\Interface`).
4. **Community-maintained** — when a new Windows build ships, the community updates the `app.config` and publishes a new NuGet version. This is the same update cadence as before, but the fix is a simple NuGet version bump instead of reverse-engineering vtable layouts.

Previously, this app vendored COM interface declarations from [MScholtes/VirtualDesktop](https://github.com/MScholtes/VirtualDesktop) and maintained two vtable layouts (pre-24H2 and 24H2+) with a runtime smoke test. The library approach is more maintainable and supports a broader range of Windows builds out of the box.

### Dear Microsoft

Please stabilize the Virtual Desktop COM interfaces or provide a proper public API. Every third-party virtual desktop tool in the ecosystem depends on reverse-engineered GUIDs that break with every update. A stable API for creating, switching, naming, and moving windows between virtual desktops would eliminate an entire class of fragility. [PowerToys has asked for this too](https://github.com/microsoft/PowerToys/issues/13993).

## Prior Art & Credits

- **[Jack251970/VirtualDesktop (Slions.VirtualDesktop)](https://github.com/Jack251970/VirtualDesktop)** — the library we use for virtual desktop COM interop. Forked from Grabacr07/VirtualDesktop, it handles version-specific COM GUIDs automatically via runtime compilation. MIT license.
- **[Markus Scholtes (MScholtes/VirtualDesktop)](https://github.com/MScholtes/VirtualDesktop)** — does the hard work of reverse-engineering and publishing updated GUIDs for every Windows build. The previous version of this app vendored his COM interface definitions.
- [Peach](https://peachapp.net) — MS Store app with similar hotkey UX
- PowerToys feature requests [#13993](https://github.com/microsoft/PowerToys/issues/13993), [#21597](https://github.com/microsoft/PowerToys/issues/21597)

## License

MIT

---

*Credit to Kieran Mockford for the original idea. May he rest in peace.*
