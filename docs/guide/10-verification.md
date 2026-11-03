# Verification: traces and logs

[Documentation](../README.md) / [Guide](./README.md) / 10 · [Русский](./10-verification.ru.md)

## Three kinds of evidence

| Check | What it establishes |
| :--- | :--- |
| Log and backup | Requested steps, captured values and retained errors. |
| Readback | A current accessible registry, resource, WMI or hardware-register value. |
| Trace and measurements | Where processing ran, its duration and the measured workload outcome. |

These checks complement each other. Matching an IMOD register to an entered value is not a mouse-latency measurement. Saved affinity does not establish where every DPC runs.

## ISR/DPC with ETW

Install Windows Performance Toolkit. A short recording can be made from an elevated console:

```powershell
xperf -on PROC_THREAD+LOADER+INTERRUPT+DPC
# Reproduce a short, predefined workload.
xperf -d trace.etl
xperf -i trace.etl -a dpcisr > report.txt
```

The `dpcisr` action produces ISR/DPC summaries: [Microsoft documentation](https://learn.microsoft.com/en-us/previous-versions/windows/desktop/xperf/the-dpc-isr-action). Open the ETL in Windows Performance Analyzer for individual events and CPUs. Do not assume that the text summary alone proves complete per-core placement.

Compare active modules, execution times, CPUs and application-thread activity. Fixed thread names, handler durations or IRQ counts are not universal reference values across Windows and driver versions.

## A fair comparison

1. Record Windows and driver versions, settings and the original backup.
2. Define a repeatable workload and a measurable outcome: frame-time distribution, CPU use, packet loss or device latency.
3. Change one factor, reboot where required and repeat several comparable runs.
4. Preserve traces and results, including regressions and instability.

Use repeated runs to distinguish a consistent change from normal variation. For unexpected CPU placement, check the device instance, applied policy and driver behaviour against the trace.

## Project QA

The developer workbench exercises device models, write guards and the UI in read-only mode. Screenshots establish only the displayed state and require inspection. MATCH means pixel equality.

Virtual runs do not validate real driver loading, physical MMIO/readback, Windows-setting writes, reboot or device resume. The workbench's 100/125/150% sizes model layout rather than actual Windows DPI transitions. Test windows are constrained to the available display, with their sizes recorded explicitly.

See [development](../DEVELOPMENT.md). Physical hardware requires separate validation with a known baseline and recovery procedure.

---

[Previous: Power saving](./09-power.md) · [Guide](./README.md) · [Next: Driver and startup](./11-driver.md)
