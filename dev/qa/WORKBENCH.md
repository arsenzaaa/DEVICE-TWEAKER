# Development workbench

[Developer checks](./README.md) · [Русский](./WORKBENCH.ru.md)

Reproduce a problem, inspect the GUI, verify a change and retain evidence. Host settings remain unchanged. Windows run on a separate invisible desktop; input is sent only to the runner's own controls.

## Short development cycle

From the repository root:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Invoke-DevSession.ps1
```

This runs `Scenarios/device-flow.json` on a hidden desktop. The result path is printed at the end: `local/dev-sessions/dev_*`.

Developer and inspector builds are shared under `local/qa-builds`, keyed by input paths/content and the selected .NET SDK. Unchanged inputs reuse the build after checking source and output inventories and SHA-256 values. Each session has separate logs, captures, sandbox state and review records. The inspector runs from its compiled DLL without rebuilding for each command.

Cache preparation and the scenario run in one bounded private job, including compiler descendants. The launcher allows 30 minutes for preparation in addition to the scenario timeout and two minutes for verification. Memory-pressure monitoring covers this entire period; an existing bounded integration worker reuses its job.

`-BuildRoot` selects a shared cache directory, including across worktrees. Use a short path when a nested checkout makes Windows DLL paths too long; it does not change Windows settings.

`-FreshBuild` creates an independent build for diagnosing a suspected build problem. An incomplete or altered cache entry stops reuse; it is preserved for inspection. Failed entries are not overwritten and old builds are never deleted automatically. Before archiving a build, retain it with the sessions that reference it, or accept that source/build verification will require restoring it. A session's `build-reference.json` identifies the exact directory and manifest hash.

Create a task-specific JSON using [device-flow.json](Scenarios/device-flow.json) or [control-states.json](Scenarios/control-states.json). After editing, rerun the same scenario against the previous successful run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Invoke-DevSession.ps1 -Scenario .\dev\qa\Scenarios\control-states.json -CompareWith 'local/dev-sessions/previous-run-directory'
```

If reproduction initially fails, keep its screenshots and exceptions. Automated comparison requires two successful runs with the same scenario and environment.

`refresh-preserves-edits.json` checks MSI/Limit drafts, modeled IMOD, scroll position and RU/EN rebuilds. `refresh-preserves-network-edits.json` checks NIC ITR input, queue counts, base CPU and the related mask after REFRESH; it does not perform hardware RSS configuration. Roles `nic-itr`, `rss-queues` and `affinity` select actual card controls; property `rss-base` reads the card state's base CPU. `select` sets NumericUpDown through its public property and normal ValueChanged event; it does not test physical arrow-key presses.

`refresh-readback-lifecycle.json` delays only the virtual IRQ source within the shared read path. It checks awaiting, progress until completion, UI release after timeout and generation-based rejection of late responses. Data comes from a preset; Windows WMI/PCI and the driver are not called. `refresh-readback-model` modes exist only in the developer workbench. The normal application's IRQ-read timeout remains 20 seconds.

## Focused resource measurements

`Measure-IdleResources.ps1 -ExePath '<production EXE>' -OutputDirectory 'local/idle-measurement' -Seconds 30` observes the exact production binary after its initial refresh and three seconds of settling. It uses the bounded invisible desktop and immutable read-only mode. `resources.json` records process CPU time, cycles, private bytes, working set, handles and threads at one-second intervals. CPU percentages include both one-core and whole-machine values. It neither sends physical input nor assumes populated Raw Input buffers; a short observation does not exclude a long-term leak.

`Invoke-DevSession.ps1 -Scenario dev/qa/Scenarios/raw-polling-performance.json` measures the actual polling update with two synthetic sample buffers, unchanged data and arriving samples. It also observes ten seconds of the real 250 ms timer. The allocation budget concerns two populated synthetic states; accelerated-loop timings include JIT effects and are not physical input latency. Native Raw Input packet delivery and monitor smoothness are outside this scenario.

## Inspecting results

`Scenarios/hardware-matrix.json` enumerates every system and CPU preset in the test panel. It runs applicable operations and injected faults through the sandbox runner, then checks REFRESH, pending edits and completed category navigation. Each profile has its own JSON report and screenshot. Unsupported operation/fault combinations are recorded as `NOT_APPLICABLE`. CPU-only profiles use the first system preset's devices. This models application decisions and virtual writes, not firmware, physical driver behavior or device timing.

Run this larger matrix with `-Scenario .\dev\qa\Scenarios\hardware-matrix.json -ScenarioTimeoutSeconds 840`. The default scenario timeout remains 120 seconds. Setting the `audit-hardware-matrix` step's `value` to `presentation` checks profile loading, refresh, drafts and navigation without repeating operation/fault cases; its report explicitly identifies that narrower scope.

Use `value: "cpu"` to rerun only CPU profiles, including operations and faults. The runner applies the selected topology and checks its name, logical processor count, core groups, efficiency classes, CCDs and CCXs before running cases. Selecting a preset in the panel alone only fills the topology editor.

- `index.html`: screenshots, control tree, steps and a transition-frame player. This local page has no external resources.
- `comparison/index.html`: before/after/difference. JSON records the changed region and differing pixel count.
- `trees/*.json`: captions, source strings, visibility, availability, focus, bounds, checks and selected values.
- `timelines/`: frames, contact sheets and actual capture intervals. Capture rate is not display FPS.
- `state/`: virtual devices, test-storage settings and full exceptions with their origin.
- `session.json`, `execution.json`, `process-*.json`: steps, hashes, normal process exit and immutable read-only mode.
- `inputs.json`, `build-reference.json`, `isolation/desktop.json`: source hashes, the shared build reference and confirmation of the hidden desktop. Full integration runs retain their own `input/repo`.

`REVIEW_REQUIRED` follows an initial run or changed screenshots. Inspect them. `MATCH` means checkpoint images match, not that the design is approved. The baseline is not overwritten; masks and tolerances are not added automatically. Incompatible scenarios/environments, missing steps, changed PNG/JSON and crash reports stop verification.

## Recording visual review

First inspect PNGs, differences and transition frames, then create a template:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Review-DevSession.ps1 -RunDirectory 'local/dev-sessions/run-directory' -CreateTemplate
```

Fill in the reviewer, overall result and every image in `review-input.json`: `ok`, `intentional-change` or `defect`, with a specific explanation. `PENDING`, omitted frames and hash mismatches are rejected.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Review-DevSession.ps1 -RunDirectory 'local/dev-sessions/run-directory'
```

The result is `REVIEWED` or `REJECTED`, tied to the exact frames. Hash verification establishes report integrity; it cannot automatically establish that someone actually inspected the images.

## Scenario actions

Unknown fields/actions and duplicate IDs are rejected. Every step must finish.

- `capture`: native client-area PrintWindow capture and control tree. No automatic rendering fallback.
- `click`, `select`, `text`, `toggle`: actual control handlers/values.
- `hover`, `leave`, `focus`: events on the runner's own control; no physical cursor.
- `key`: one key without modifiers, WM_KEYDOWN/WM_KEYUP to a specific control; the main window uses its navigation handler. Physical input, IME and global Windows routing are not tested.
- `assert`, `wait`: text, value, availability, focus, checked state or caret position; waits have timeouts.
- `language`, `resize`, `scroll`: RU/EN, window size and list Home/End/PageUp/PageDown.
- `reveal`: scroll to a control. Interacting with a clipped or unavailable control fails; `viewport` explicitly checks visibility.
- `transition`: pre-action capture, frame sequence, change/settling assertions and final capture.
- `save-state`, `restore-state`: virtual CPU/device/settings JSON and loading it into the UI.
- `apply-model`: shared DeviceSettingsWriter writes test JSON; the result is verified by reading the file.
- `model-error`: an actual IOException in fake storage, the shared writer's stack and a real dialog. No Windows/driver calls.
- `audit-editors`: every visible normal text field in the current device set, scrolling, native PNG, RGB comparison against a taller reference, focus and leaving the field. Saves close-ups and card context; unknown field types stop verification.
- `gallery`: the existing test-panel gallery.

Control selectors: `@main`, `@active`, `id:NAME`, `text:CAPTION`. For cards, specify `device` with a unique name/Instance ID substring and a role: `role:title`, `role:msi`, `role:limit`, `role:priority`, `role:policy`, `role:imod`, `role:imod-toggle`, `role:imod-check`, `role:imod-current`, `role:mode`. Multiple matches fail; refine the scenario. Use IDs for translated buttons.

Transitions accept `trigger`, `durationMs`, `intervalMs`, `requireChange`, `requireSettled`. The last two default to true. Missing changes or unfinished transitions fail. Frames still require inspection; passing is not an automatic smoothness guarantee.

At least three real frames are saved. Slow capture can make observation exceed `durationMs`; `timelines` records actual frame times and intervals. Requested intervals are not presented as achieved frame rates.

## Full verification

[device-palette.json](Scenarios/device-palette.json) shows every badge class: NVIDIA/AMD/Intel/unknown GPU, NDIS/NetAdapterCx, USB, audio and five storage variants. It captures EN/RU and language transitions. This is a separate visual scenario without physical devices.

[device-colors.json](Scenarios/device-colors.json) shows actual virtual-device GPU/NET cards beside USB and NVMe in EN/RU. Run and compare through `Invoke-DevSession.ps1`.

[usb-role-order.json](Scenarios/usb-role-order.json) shows all USB roles on one controller: EN/RU headers, language transition and a narrowed window. Run `Invoke-DevSession.ps1 -Scenario .\dev\qa\Scenarios\usb-role-order.json`. This caption-order review scenario runs separately from the general suite.

`Invoke-Qa.ps1` runs the integration scenarios twice and compares checkpoints: `device-flow`, `control-states`, `dtimod-status`, `imod-check`, `imod-editor`, four `text-fields-*`, six `dialog-audit-*`, `backup-menu-flow`, `navigation-audit`, `irq-check-loading` and the three `refresh-reuse*` scenarios. This checks reproducibility of the current build. Gallery, EN/RU, scaling models, log/resource audits, tests and intentional mutations follow. `dtimod-status` uses shared lifecycle code with fake backends and actual error dialogs, without loading the driver.

Registry, power plans, startup and drivers are neither changed nor completely emulated. A new feature needs its own scenario and shared technical-code checks against fake data; coverage does not appear automatically.

`ui-controls` compares glyph pixels in editors and the "Аудио" caption with taller native controls. It covers 0xFA0, Latin/Cyrillic descenders and 100/125/150% size/font models; intentionally clipped controls must differ. This checks glyph completeness, not font quality or actual Windows DPI transitions.

`text-fields-*` checks actual search, MSI Limit, IMOD, role-based IMOD, NIC ITR and RSS fields across Intel, AMD, IMOD and NIC device sets. Each scenario uses EN at 1200×900 and RU after narrowing to 1000×720. The inner editor is compared with a taller field of identical width, text, font and colours: vertical glyph completeness, not capacity for arbitrary long input. Bounds are checked separately against the viewport; context captures support visual layout review.

## Dialogs and menus

`dialog-audit-en/ru` covers save/restore windows, file paths, system checks with known/unavailable data, confirmations, backup location, empty/normal/long lists and success/warnings/partial/failed/stress results. `-120/-144` variants set `layoutDpi` to 120/144: 125/150% models without changing Windows DPI. `audit-dialog` records label height, button capacity and window bounds.

`backup-menu-flow` uses a fake callback through the shared menu to verify deletion: window retention, original-backup protection, an empty list, failure/exception, selection of a remaining file and full paths. No real backup files are deleted. `navigation-audit` opens actual dropdowns with F4, navigates items, closes with Escape and changes categories in EN/RU.

`dialog` invokes only allowed previews of shared windows. `active:id:...` selects a control in the active nested window; ordinary `id:...` still rejects ambiguity. Form Enter/Escape goes through its dialog-key handler; other controls receive their own messages. System focus and physical input are not modeled. A report and `MATCH` do not replace PNG review.

In `transition`, a nonempty `value` checks the target's selected `property` before the action and on every observed frame. `imod-check` uses this for a stable CHECK/ПРОВЕРИТЬ caption throughout rereading. `check-presentation` separately captures hover, focus, IRQ/HW lines and the backup picker in EN/RU. Category buttons scroll to devices; search hides nonmatching cards.

`irq-check-loading` presses the global IRQ button in the top row in EN/RU: GPU/USB, success, failure and retry. Only hardware reads are replaced with delayed responses. Every PCI device, including filtered cards, must be queried once. Each frame checks dark technical backgrounds, form availability, disabled buttons and exactly one working indicator. No driver is loaded.

`interface-wording` shows the main window, IMOD, system details, restore and result dialogs in EN/RU. `dialog-audit-*` also checks pre-restore, pre-release and IMOD/NIC ITR save reasons. JSON reasons and filenames are not translated.

`irq-field-layout` compares IRQ Mode fields and navigation before/after the change in EN/RU, GPU/USB and 1200/1000 widths. `irq-placement` adds refresh hover/focus; `irq-toolbar-dense` checks search/category capacity on a dense device set. The top row has four main actions and the global IRQ Mode read button.

`irq-control-states` captures the IRQ Mode field and global button in normal, hover, focus and reading states in EN/RU, with layoutDpi 96/120/144 supported. `irq-glyph-errors` checks leaving the reading state after failure/retry, Russian status capacity and complete PCI coverage.

`irq-global-label` checks the single global IRQ button: hidden PCI devices, failure and retry. Keyboard activation must query the same complete set once each, without driver loading.

`irq-summary` checks the five-button top row: one global IRQ button outside the cards, no overlaps, caption capacity, EN/RU, success/failure/retry, complete PCI coverage, Space activation and 1200/1000 widths. layoutDpi 120/144 variants check scaling models.

`irq-top-actions` retains comparisons of the former strip and the global READ IRQ MODE/СЧИТАТЬ IRQ button in the top row: EN/RU, reads, failure, retry, focus and 1200/1000 widths. `irq-summary` separately requires exactly five actions, fitting captions and absence of the previous strip.

## Interaction latency and motion

`ui-latency` measures synchronous search, navigation, draft toggles, IMOD details, language and presentation handlers on the DevicePalette model. `state/handlers.json` contains individual timings and allocations. The waits between samples allow painting and are excluded from handler timings. Results describe this device/CPU model and host; they are not hardware latency or display FPS.

`irq-stable-value` retains the last IRQ Mode during a slow read. `irq-repeat-clicks` exercises repeated button clicks and re-entrant commands through the shared operation guard, including failure and retry; only the read backend is synthetic.

`toggle-motion` captures the native toggle component in EN/RU: immediate state/events, 120 ms movement, reversal, disabled/reduced-motion state and hidden/disposed frame subscription cleanup. The initial movement uses the default production policy, including when Windows client-area animation is off. A QA-only override checks disabled local motion. High Contrast disables motion; no Windows preference is changed.

`ui-motion` exercises the actual Power Saving draft control, wheel accumulation/reversal, category retargeting, thumb capture, scroll-range changes, button hover and frame subscription cleanup in EN/RU. Ordinary navigation steps await the visual destination; transition steps retain intermediate frames. Main-list wheel/category/page movement takes 80 ms; button colour feedback takes 100 ms. Direct thumb dragging remains immediate.

`filter-layout-transitions` hides cards while changing language, width and layout DPI, then reveals them. Use it with the existing USB/network presentation and REFRESH preservation scenarios when changing layout reuse.

### Frame pacing and motion lifecycle

`frame-pacing` records native toggle paint intervals at a modeled 300 Hz, two synthetic refreshes, UI heartbeat gaps and layout/card construction timings. The shared frame clock allows at most one pending UI callback; it does not render idle frames. These measurements describe off-screen application painting, not physical display FPS or full hardware scan duration.

`motion-lifecycle` checks the always-on default with real native controls: completion when hidden, no idle frame posts, movement after showing, coalescing during a busy UI and worker cleanup on disposal. It also verifies removal of the old system-menu command. JSON tests cover dropping the obsolete flag on language save while retaining unknown settings.

`dialog-consistency.json` covers driver consent, the USB IMOD choice before and after consent, and driver information on EN/RU at 100/125/150% layout scaling. Dialog audits check text fit and a shared action baseline. Centered messages use centered action groups; technical and selection dialogs keep the scaled 20-pixel right inset. Both retain the 20-pixel bottom inset. `dialog-audit-en.json` and `dialog-audit-ru.json` cover the remaining dialog surfaces.

Ordinary checkpoints wait for active animations in the main window and owned dialogs to finish after painting their current targets. Transition steps keep recording the motion itself. Pixel tolerance and masks are unchanged.

`dialog-result-alignment` covers reset completion, backup results and backup-location choices in EN/RU at 100/125/150% layout scaling.
