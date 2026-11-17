# Verification: traces and logs

[Documentation](../README.md) / [Guide](./README.md) / 10 · [Русский](./10-verification.ru.md)

## After applying settings

Read the operation summary before closing it. It distinguishes completed changes, warnings and partial failures. **DETAILS** identifies the affected setting and stage; **LOGS** and **BACKUPS** open the files needed to inspect or restore the operation.

For MSI, Interrupt Affinity and ReservedCpuSets, restart Windows before checking runtime behavior. Read **IRQ Mode** again after changing MSI. For USB IMOD and NIC ITR, use **CHECK** to read the current registers. Saved-profile users should also check the startup log after a reboot.

## Logs and backups

Application logs are stored in `logs` beside the EXE. Startup logs are in `%ProgramData%\DEVICE TWEAKER\Startup\logs`. When reporting a problem, include the session in which it happened, the device, the action and what you observed. The log records build identity, requested values, completed stages and errors.

A backup stores settings for restoration. A register readback shows the value returned at that moment. To check where interrupt processing ran, collect an ETW trace: the saved mask and readback do not contain an ISR/DPC execution history.

## Compare before and after

1. Keep the original backup and record Windows, device and driver versions.
2. Use the same game scene or workload and the same measurement tool. Choose a useful metric: frame-time distribution, CPU load, packet loss or device latency.
3. Change one setting, restart when required, and repeat comparable runs. Keep regressions as well as improvements.

Compare several runs so ordinary variation is visible. For an affinity change, inspect the relevant driver's ISR/DPC activity alongside application threads and frame times.

## Inspecting a trace

Windows Performance Recorder and Analyzer can help inspect execution time and processor placement. The Windows Performance Toolkit is available through the [Windows ADK](https://learn.microsoft.com/en-us/windows-hardware/get-started/adk-install).

With [xperf](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/xperf-command-line-reference) installed, start and stop a basic DPC/ISR capture from an elevated terminal:

```text
xperf -on latency
xperf -d interrupts.etl
```

Run the workload between the two commands. Inspect the relevant modules, processor numbers and durations in the resulting trace. The exact analysis depends on the driver and Windows build; retain the trace when asking for help.

## If the result differs from expectations

Check the device instance, saved policy and whether a required reboot occurred. For IMOD, reread the map after reconnecting a USB device. For networking, compare active RSS settings as well as interrupt affinity.

If the device becomes unstable, use **RESTORE** with the appropriate backup and check operation again. Include the operation log and the workload conditions in a [bug report](../../.github/ISSUE_TEMPLATE/bug_report.yml).

---

[Previous: Power saving](./09-power.md) · [Guide](./README.md) · [Next: Driver and startup](./11-driver.md)
