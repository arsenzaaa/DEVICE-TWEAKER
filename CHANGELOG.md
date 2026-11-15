# Changelog

[DEVICE TWEAKER](./README.md) / Changelog · [Русский](./CHANGELOG.ru.md)

## 0.1.0

Initial release of **DEVICE TWEAKER REWORK**.

### Device configuration

- Configure MSI Mode, MSI Limit, IRQ Priority and interrupt affinity per device. Saved registry settings and the current IRQ mode are reported separately.
- Use automatic optimization to select CPU cores for graphics cards, input devices and network adapters. Selection accounts for physical cores, SMT siblings, P/E cores, CCDs and available CPPC rankings. Manual adjustment and separate ReservedCpuSets controls are available.
- Select RSS processors and queue counts, configure device power saving and Mouse Throttle, and measure mouse event rates through Raw Input.

### USB IMOD and network moderation

- Set xHCI IMOD for an entire controller, by device role or by interrupter index. Read register values back to check the result.
- Configure and read back NIC ITR on supported Intel and Realtek adapters.
- Save IMOD and NIC ITR profiles for application at sign-in, after resume and on supported device-connection events, without keeping the GUI open. These hardware operations use DTIMOD; [driver requirements and compatibility limits](./docs/guide/11-driver.md) apply.

### Interface and responsiveness

- Switch between English and Russian, search the device list and filter it by device category.
- REFRESH discovers devices again and rereads their settings while reusing compatible cards. Pending edits and scroll position are preserved where applicable; changed or disconnected devices are handled using the new discovery results.
- Device discovery and hardware-state reads run in the background. Mouse event-rate estimates reuse their calculated results until the input samples change, and the animation clock stops when there are no active animations.

### Backup, restore and diagnostics

- Keep an original-settings backup and save additional snapshots. Restore device settings and saved IMOD/NIC ITR profiles from the restore dialog.
- Operation reports distinguish completed steps, skipped actions, warnings and failures. Session logs include the build identity, target devices, requested values and stage results for troubleshooting partial operations.

### Distribution and development

- Portable Windows 10/11 x64 builds are available with .NET included or requiring .NET 10 Desktop Runtime x64.
- The source includes hardware and failure models, regression tests and a read-only GUI workbench. QA entry points serialize workloads and enforce memory limits. Models cover application behavior; physical driver loading, register access and Windows power/PnP behavior require hardware verification.
