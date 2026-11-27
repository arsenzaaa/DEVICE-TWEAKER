# Verifying results

[Documentation](../README.md) / [Guide](./README.md) / 10 · [Русский](./10-verification.md)

## Confirming applied state vs runtime behavior

Writing a registry value does not prove that a driver adopted it or that ISR/DPC placement improved under load. Treat configuration state, hardware readback and measured runtime traces as three separate layers of evidence:

- **MSI, Interrupt Affinity, ReservedCpuSets:** reboot Windows first, then verify the active interrupt mode with **READ IRQ MODE** and inspect the result log (**DETAILS** / **LOGS**).
- **USB IMOD and NIC ITR:** click **CHECK** in the card to read live MMIO registers back. If you saved a startup profile, verify `%ProgramData%\DEVICE TWEAKER\Startup\logs` after rebooting, and reread the USB map after reconnecting peripherals.
- **RSS:** compare the saved settings in the application log against `Get-NetAdapterRss` ([Chapter 6](./06-network.md)).

Application logs live in `logs` beside the EXE and record the build, requested settings, completed stages and errors. Use **RESTORE** to roll back from any saved backup.

## Controlled before/after testing

1. Keep the initial backup intact and record the Windows build, device model, driver version and **VBS/HVCI** state. Under an active hypervisor (VBS/HVCI), every external hardware interrupt triggers a `VMEXIT` and SynIC delivery (*Windows Internals*, 7th ed., Part 2, ch. 9) — never compare traces captured with different virtualization states.
2. Pick a metric that matches the setting: frame times (1% / 0.1% lows and variance) and GPU ISR/DPC placement for GPU affinity; packet loss, CPU load and receive DPCs for RSS/ITR; event jitter for USB IMOD. Run the exact same scene or traffic pattern multiple times to gauge run-to-run noise.
3. Change one variable at a time, reboot when required, and repeat the same runs with the same tool.

## Inspecting a trace

Windows ETW kernel tracing carries zero overhead while inactive (the kernel's performance group masks remain cleared — *Windows Internals*, 7th ed., Part 2, pp. 499–500). During an active capture, keep the provider flags and trace duration identical across before/after comparisons (typically 10–30 seconds).

### Xperf and WPA

Install [Windows Performance Toolkit](https://learn.microsoft.com/en-us/windows-hardware/test/wpt) from the Windows ADK and start a kernel trace from an elevated terminal:

```text
xperf -on PROC_THREAD+LOADER+INTERRUPT+DPC+CSWITCH
```

If you need to attribute a shared framework module such as `Wdf01000.sys` or `ndis.sys` to a specific device driver, enable sampling call stacks with `PROFILE` and `-stackwalk Profile` (`DPC` and `INTERRUPT` are kernel event flags rather than `-stackwalk` keywords):

```text
xperf -on PROC_THREAD+LOADER+PROFILE+INTERRUPT+DPC+CSWITCH -stackwalk Profile
```

Reproduce the workload, stop the trace and generate a [DPC/ISR summary report](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/dpcisr):

```text
xperf -d interrupts.etl
xperf -i interrupts.etl -o interrupts.txt -a dpcisr
```

Durations in the `-a dpcisr` report are in **microseconds**. Compare call counts and total/maximum durations per driver module on each CPU across equal capture windows. As a practical reference from Microsoft's multimedia guidance, target budgets are ISR ≤ 25 µs and DPC ≤ 100 µs; isolated spikes above 500–1000 µs on a core running game or audio threads deserve closer inspection.

Open `interrupts.etl` in **Windows Performance Analyzer (WPA)**:

1. Locate the driver module in [DPC/ISR Duration by Module, Function](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/list-of-wpa-graphs).
2. Check which logical CPUs actually executed its handlers in **DPC/ISR Duration by CPU**.
3. Overlay **CPU Usage (Precise)** over the same time slice to verify whether the driver's ISRs/DPCs still collide with your application's busy threads.

### Media eXperience Analyzer (MXA)

For audio dropouts or video stutter, capture a trace with a [Windows Performance Recorder (WPR) multimedia profile](https://learn.microsoft.com/en-us/windows-hardware/test/wpt/optimizing-windows-devices-for-multimedia-experiences-exercise-2) and open it from the MXA directory with `xa -i "C:\Traces\Media.etl"`.

Locate a dropout in **Audio Glitches**, inspect concurrent driver activity in **CPU Scheduler** over that exact interval, and open **Callstacks** (when stacks were captured) to pinpoint the offending module — Microsoft's [MXA walkthrough](https://learn.microsoft.com/en-us/shows/defrag-tools/151-media-experience-analyzer-part-3-audio-glitch-analysis) shows a textbook case of a long network-driver DPC causing audio glitches.

## When something goes wrong

Check the exact device instance and failed step under **DETAILS**. If a device becomes unstable, restore the original backup and verify stability under the same load. When opening a [bug report](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml), include the action taken, observed behavior and the relevant log excerpt.

---

[Previous: Power saving](./09-power.md) · [Guide](./README.md) · [Next: Driver and startup](./11-driver.md)
