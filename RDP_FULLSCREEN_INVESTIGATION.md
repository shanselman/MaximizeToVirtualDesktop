# RDP Fullscreen Investigation WIP

## Status

This branch preserves the experimental work from the RDP fullscreen investigation. The current technical conclusion is that **true MSTSC/MSRDC fullscreen, especially multimon fullscreen, should not be advertised as supported yet**.

The experiments proved that MaximizeToVirtualDesktop can sometimes identify a manageable RDP application-view HWND, create an MVD desktop, move that HWND, switch to the new desktop, and preserve the RDP fullscreen geometry. However, the user experience is still unsafe: once true fullscreen RDP is on the MVD desktop, RDP owns or redirects the normal local escape paths, and the user may not have a reliable way to return without minimizing/exiting fullscreen first.

This is valuable research and should be kept, but the current behavior should be treated as **work in progress**, not product-ready.

## Original problem

The existing Shift+Click and maximize-transition flow was built around ordinary Windows maximize semantics:

1. Watch `EVENT_OBJECT_LOCATIONCHANGE` and `EVENT_SYSTEM_MOVESIZEEND`.
2. Detect a window whose `WINDOWPLACEMENT.showCmd` becomes `SW_MAXIMIZE`.
3. If Shift behavior says "send to virtual desktop", defer until resizing settles.
4. Move the window to a new virtual desktop.
5. Switch to that virtual desktop.
6. Maximize the window and track it.
7. Restore/cleanup when the tracked window is restored, closed, or toggled.

MSTSC/MSRDC true fullscreen does not reliably participate in that model. In true fullscreen, the RDP client often behaves like a borderless monitor-sized shell surface, not a normally maximized top-level window. That means `GetWindowPlacement` can report a normal state even while the user sees RDP as fullscreen.

The first idea was to treat an RDP HWND as "fullscreen" when:

- the owning process is `mstsc` or `msrdc`
- the window is borderless (`WS_CAPTION` and `WS_THICKFRAME` are absent)
- `GetWindowRect` matches either the nearest monitor rect or the full virtual-screen rect, with a small tolerance

That detection was good enough to find RDP fullscreen transitions, but it exposed deeper state-management and UX issues.

## Important implementation changes preserved on this branch

### Shared fullscreen predicate

`WindowMonitor` now has a shared fullscreen-state decision rather than a single `showCmd == SW_MAXIMIZE` check. The experimental state machine distinguishes:

- `None`
- `Maximized`
- `RdpFullscreen`

This was necessary because RDP fullscreen can be fullscreen by geometry even when `WINDOWPLACEMENT.showCmd` is not `SW_MAXIMIZE`.

### RDP geometry helpers

`NativeMethods` now includes additional Win32 interop used by the investigation:

- `GetWindowRect`
- `MonitorFromWindow`
- `GetMonitorInfo`
- `GetSystemMetrics`
- `GetWindowLongPtr` / `GetWindowLong`
- `GetAncestor`
- constants for virtual-screen metrics, monitor selection, styles, and root-owner checks

The `GetWindowLongPtr` path includes the proper Win32 zero-return handling pattern:

1. Clear last error with `SetLastError(0)`.
2. Call `GetWindowLongPtr`.
3. Treat zero as an error only if `Marshal.GetLastWin32Error()` is non-zero.

This matters because a style value of zero is technically valid.

### RDP-origin tracking flag

`TrackingEntry` now records whether a window was tracked via the RDP fullscreen path. This was added because normal tracked windows can be restored when they stop being maximized, but RDP fullscreen emits noisy geometry changes during resize/reconnect/multimon negotiation. For RDP-origin tracked windows, immediate restore-on-location-change is unsafe.

### Manager in-flight guard

`FullScreenManager` now uses a direct in-flight guard around `MaximizeToDesktop` and `Restore`, not just `Toggle`.

This was needed because the WinEvent monitor can call manager methods directly, and RDP can emit many events for the same logical transition. Without the direct guard, repeated callbacks could queue overlapping restore/maximize operations.

The public methods now go through `RunWithInFlight`, while internal cross-calls use core methods to avoid breaking toggle behavior.

### RDP candidate coalescing

The latest experiment changed RDP handling from "process every fullscreen-looking RDP HWND" to "collect a batch of fullscreen-looking RDP HWNDs and pick one manageable candidate."

This was necessary because a single true fullscreen transition produced many top-level RDP-related HWNDs. Most were not valid virtual-desktop application views:

```text
VirtualDesktopService: GetDesktopIdForWindow failed: Element not found. (0x8002802B (TYPE_E_ELEMENTNOTFOUND))
```

The coalescing path now:

1. Adds every RDP fullscreen-looking HWND to `_pendingRdpMaximize`.
2. Waits for a debounce interval.
3. Rechecks that candidates still look like RDP fullscreen.
4. Prefers foreground/root-owner candidates.
5. Calls `FullScreenManager.CanManageWindow` so only an HWND with a desktop ID is selected.
6. Moves only the selected HWND.

### RDP geometry preservation

The latest version avoids calling `ShowWindow(hwnd, SW_MAXIMIZE)` for RDP fullscreen windows after moving them to the MVD desktop. It also avoids applying the saved original `WINDOWPLACEMENT` during RDP restore.

This change was made because the "zoomed-out RDP on another monitor" artifact strongly suggested that after RDP had negotiated fullscreen/multimon geometry, our additional normal-window maximize/placement call was forcing RDP back into a scaled or mismatched local geometry state.

The current branch logs:

```text
FullScreenManager: Preserving RDP fullscreen geometry.
FullScreenManager: Preserving RDP fullscreen placement during restore.
```

## Manual test timeline and findings

### Test 1: initial RDP geometry detection

The first working detector queued multiple RDP HWNDs and moved one:

```text
WindowMonitor: Queued maximize for window 2037862 after resize end.
WindowMonitor: Queued maximize for window 9965530 after resize end.
WindowMonitor: Queued maximize for window 11475060 after resize end.
WindowMonitor: Queued maximize for window 11342982 after resize end.
WindowMonitor: Queued maximize for window 329248 after resize end.
WindowMonitor: Queued maximize for window 525668 after resize end.
```

One candidate failed:

```text
VirtualDesktopService: GetDesktopIdForWindow failed: Element not found. (0x8002802B (TYPE_E_ELEMENTNOTFOUND))
FullScreenManager: Could not determine original desktop, aborting.
```

Another moved successfully:

```text
VirtualDesktopService: Created desktop ac5442c8-11f6-4209-82ef-82449e8087dd
VirtualDesktopService: Moved window 11475060 to desktop ac5442c8-11f6-4209-82ef-82449e8087dd
VirtualDesktopService: Switched to desktop ac5442c8-11f6-4209-82ef-82449e8087dd
FullScreenTracker: Now tracking 11475060 (total: 1)
FullScreenManager: Successfully moved window to desktop ac5442c8-11f6-4209-82ef-82449e8087dd
```

Then location-change noise immediately restored it:

```text
WindowMonitor: Tracked window 11475060 restored via location change.
FullScreenManager: Restoring window 11475060 from temp desktop ac5442c8-11f6-4209-82ef-82449e8087dd
```

This proved that using normal `LOCATIONCHANGE` restore semantics for RDP fullscreen was wrong.

### Test 2: skip location-change restore for RDP-origin tracked windows

After adding the RDP-origin tracking flag and delaying/coalescing restore attempts, the app no longer immediately restored from location-change noise. However, it still processed multiple RDP candidates independently. Five candidates failed `GetDesktopIdForWindow`, and one moved:

```text
WindowMonitor: Queued maximize for window 9833140 after resize end.
WindowMonitor: Queued maximize for window 4654738 after resize end.
WindowMonitor: Queued maximize for window 597728 after resize end.
WindowMonitor: Queued maximize for window 1970024 after resize end.
WindowMonitor: Queued maximize for window 10159182 after resize end.
WindowMonitor: Queued maximize for window 17833180 after resize end.
...
WindowMonitor: Fallback processing for pending maximize window 4654738.
VirtualDesktopService: Created desktop 01ee06d7-6550-4428-938d-92087fe7c285
VirtualDesktopService: Moved window 4654738 to desktop 01ee06d7-6550-4428-938d-92087fe7c285
VirtualDesktopService: Switched to desktop 01ee06d7-6550-4428-938d-92087fe7c285
FullScreenTracker: Now tracking 4654738 (total: 1)
FullScreenManager: Successfully moved window to desktop 01ee06d7-6550-4428-938d-92087fe7c285
```

The user still saw a zoomed/scaled RDP view, likely caused by post-move normal maximize/placement behavior fighting RDP's fullscreen geometry.

### Test 3: coalesced RDP candidate selection and geometry preservation

The latest branch state produced the closest result:

```text
WindowMonitor: Queued RDP fullscreen candidate 24775574.
WindowMonitor: Queued RDP fullscreen candidate 14622756.
WindowMonitor: Queued RDP fullscreen candidate 10362038.
WindowMonitor: Queued RDP fullscreen candidate 14622756.
WindowMonitor: Queued RDP fullscreen candidate 11606424.
WindowMonitor: Queued RDP fullscreen candidate 14025924.
WindowMonitor: Queued RDP fullscreen candidate 58269810.
WindowMonitor: Queued RDP fullscreen candidate 1120334.
VirtualDesktopService: GetDesktopIdForWindow failed: Element not found. (0x8002802B (TYPE_E_ELEMENTNOTFOUND))
WindowMonitor: Processing RDP fullscreen candidate 14622756 from 7 window(s).
VirtualDesktopService: Created desktop ddff526c-32fb-42a2-8d64-78bbc9d23e9d
VirtualDesktopService: Moved window 14622756 to desktop ddff526c-32fb-42a2-8d64-78bbc9d23e9d
VirtualDesktopService: Switched to desktop ddff526c-32fb-42a2-8d64-78bbc9d23e9d
FullScreenManager: Preserving RDP fullscreen geometry.
FullScreenTracker: Now tracking 14622756 (total: 1)
FullScreenManager: Successfully moved window to desktop ddff526c-32fb-42a2-8d64-78bbc9d23e9d
```

The user reported that this mostly worked: RDP ended up on the MVD desktop and could be switched to. The remaining problem was escape/return. Since true fullscreen RDP owns the local input path, RDP did not know that the local app had created a private MVD desktop and provided no natural way to get back.

The user could recover only by minimizing/exiting fullscreen enough to regain local shell control.

## Current conclusion

The code path is technically interesting and may be useful later, but **true RDP fullscreen should probably be blocked rather than supported**.

Reasons:

1. RDP fullscreen is not just a normal maximized window. It can be a set of borderless shell/client surfaces with internal geometry negotiation.
2. A single logical fullscreen transition emits many top-level RDP HWNDs. Most are not manageable virtual-desktop application views.
3. Moving the manageable HWND can work, but RDP does not know it was placed on an app-created local virtual desktop.
4. Once in fullscreen, RDP may capture or redirect normal local escape paths, including shortcuts the user expects to restore or switch desktops.
5. A low-level keyboard escape hook might be possible, but that would be fragile, surprising, and likely to fight RDP/security/input-redirection behavior.
6. Shipping this could strand users on an MVD desktop with a fullscreen remote session and no obvious return path.

The safe product decision is:

- Support normal windowed/maximized RDP if it behaves like a normal window.
- Do not support true MSTSC/MSRDC fullscreen or multimon fullscreen.
- Detect true RDP fullscreen and show a notification instead of moving it.
- Remove or revise documentation claiming MSTSC/MSRDC fullscreen compatibility.

Suggested user-facing notification:

```text
RDP fullscreen isn't supported.
Exit fullscreen or minimize RDP, then use the hotkey.
```

## Future recovery ideas

If this work is revisited later, possible avenues:

1. Investigate whether MSTSC/MSRDC expose a reliable local-client escape API or message that exits fullscreen before moving.
2. Investigate RDP window class names and application views more deeply so only the canonical client host is considered.
3. Test whether moving an RDP session before it enters fullscreen is safer than reacting after it is already fullscreen.
4. Add an explicit local-only emergency hotkey implemented with `WH_KEYBOARD_LL`, but only if it can be made predictable and documented.
5. Consider a block-first approach with an advanced experimental flag for people who accept the escape-path risk.

## Verification performed during the WIP

Repeated local builds succeeded after changes:

```powershell
dotnet build .\MaximizeToVirtualDesktop.slnx
```

The solution has no real test project, so `dotnet test --no-build` effectively only validated that the solution could be processed:

```powershell
dotnet test .\MaximizeToVirtualDesktop.slnx --no-build
```

A temporary no-NuGet reflection harness was created outside the repo under the Copilot artifacts directory to exercise the pure rectangle logic. It verified:

- exact monitor match is fullscreen
- within two-pixel tolerance is fullscreen
- outside tolerance is not equal
- multimon virtual-screen match is fullscreen
- ordinary window is not fullscreen

## Recommendation for the next branch

Start from this WIP only as reference. The likely shippable fix should be much smaller:

1. Keep enough detection to identify true MSTSC/MSRDC fullscreen.
2. Refuse the operation with a notification.
3. Avoid moving RDP fullscreen HWNDs at all.
4. Update README to remove fullscreen/multimon RDP from the compatibility table.

This branch intentionally preserves the deeper experimental implementation so the investigation is not lost.
