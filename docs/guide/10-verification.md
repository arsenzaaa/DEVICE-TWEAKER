# Verifying changes

[Documentation](../README.md) / [Guide](./README.md) / 10 · [Русский](./10-verification.ru.md)

## After applying settings

The result dialog shows which settings were applied. Open **DETAILS** to see which device and step produced a warning or error. **LOGS** and **BACKUPS** open their respective folders.

- After changing MSI, interrupt affinity or ReservedCpuSets, restart Windows before testing. Use **READ IRQ MODE** to check which interrupt mode the device is using.
- For USB IMOD or NIC ITR, click **CHECK** in the device card to read the current registers. If you saved a startup profile, check its log after reboot. Reread the USB map after reconnecting a device.
- For networking, compare the saved settings with the active RSS configuration. The [network chapter](./06-network.md) includes the PowerShell command.

## Logs and backups

Application logs are stored in `logs` beside the EXE. Startup logs are in `%ProgramData%\DEVICE TWEAKER\Startup\logs`. Logs record the build, requested settings, completed steps and errors. Use **RESTORE** to apply settings from a backup.

## Compare before and after

1. Keep the original backup and record the Windows build, device model and driver version.
2. Choose a metric for the setting you are testing, such as frame times for GPU affinity or packet loss and CPU load for RSS. Run the same game scene or workload several times to establish how much the result varies.
3. Change one setting, reboot if needed, and repeat the same runs with the same measurement tool. Compare the spread of results across runs.

## Inspecting a trace

An ETW trace includes the events enabled during recording. For interrupt tuning, capture ISR/DPC events and thread switches so you can compare driver activity with application threads. See *Windows Internals*, 7th ed., Part 2, pp. 499–500, for how ETW sessions collect events.

### Xperf and WPA

Install [Windows Performance Toolkit](https://learn.microsoft.com/en-us/windows-hardware/test/wpt) from the Windows ADK and start a short kernel trace from an elevated terminal:

```text
xperf -on PROC_THREAD+LOADER+INTERRUPT+DPC+CSWITCH
```

Reproduce the workload, then save the trace and generate a [DPC/ISR report](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/dpcisr):

```text
xperf -d interrupts.etl
xperf -i interrupts.etl -o interrupts.txt -a dpcisr
```

In the report, look at the event count and duration for the driver module on each CPU. Compare recordings of the same workload over equal time spans. A shared module such as `Wdf01000.sys` can appear for several devices. To pin down the source, capture call stacks and inspect the driver frames.

Open the ETL in **Windows Performance Analyzer (WPA)**. Start with [DPC/ISR Duration by Module, Function](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/list-of-wpa-graphs) to find the module, then check its placement in **DPC/ISR Duration by CPU**. Use **CPU Usage (Precise)** over the same time range to see which application threads were running.

### Media eXperience Analyzer (MXA)

In MXA, **CPU Scheduler** shows when and where threads ran. To investigate audio dropouts or video stutter, record the relevant events with a [media profile in Windows Performance Recorder (WPR)](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/optimizing-windows-devices-for-multimedia-experiences-exercise-2). Open the capture from the MXA folder with `xa -i "C:\Traces\Media.etl"`, using your trace path.

For audio, zoom in on a dropout in **Audio Glitches** and check driver activity in **CPU Scheduler** over the same interval. Use **Callstacks** to identify the driver if stacks were recorded. Microsoft's [MXA walkthrough](https://learn.microsoft.com/en-us/shows/defrag-tools/151-media-experience-analyzer-part-3-audio-glitch-analysis) demonstrates this with a long DPC in a network driver.

## If the result differs from expectations

Check the device instance and the failed step in **DETAILS**. If the device becomes unstable, restore the original settings and repeat the workload. Include the action, observed result and relevant log excerpt in a [bug report](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml).

---

[Previous: Power saving](./09-power.md) · [Guide](./README.md) · [Next: Driver and startup](./11-driver.md)
