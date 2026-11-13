using System.Diagnostics;
using DeviceTweakerQa;

namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private async Task DevMeasureRawPollingAsync(string id, bool requireBudget)
    {
        if (!QaRun.IsReadOnly || !_testDevicesOnly || _blocks.Any(b => !b.Device.IsTestDevice))
            throw new InvalidOperationException("Polling measurements require isolated synthetic devices.");
        var previous = _rawPollingStates.ToArray();
        bool timerEnabled = _rawPollingTimer?.Enabled == true;
        _rawPollingTimer?.Stop();
        _rawPollingStates.Clear();
        using Process process = Process.GetCurrentProcess();
        List<object> measurements = [];
        double unchangedBytesPerTick = 0;
        try
        {
            for (int device = 0; device < 2; device++)
            {
                RawPollingState state = new() { Role = device == 0 ? "Mouse" : "Keyboard",
                    InstanceId = "qa-polling-" + device, ControllerId = "qa-polling-controller-" + device };
                for (int sample = 0; sample < 256; sample++) state.RecordInterval(device == 0 ? 1d : .125d);
                _rawPollingStates[new IntPtr(-100 - device)] = state;
            }
            // Finish the real confirmation path and JIT before measuring steady work.
            for (int i = 0; i < 32; i++) UpdateRawPollingMeasurements();
            await Task.Delay(500);
            foreach (bool newSamples in new[] { false, true })
            {
                const int iterations = 2000;
                long allocated = GC.GetAllocatedBytesForCurrentThread();
                long started = Stopwatch.GetTimestamp();
                for (int i = 0; i < iterations; i++)
                {
                    if (newSamples)
                        foreach (RawPollingState state in _rawPollingStates.Values)
                            state.RecordInterval(state.Role == "Mouse" ? 1d : .125d);
                    UpdateRawPollingMeasurements();
                }
                double elapsedMs = Stopwatch.GetElapsedTime(started).TotalMilliseconds;
                long bytes = GC.GetAllocatedBytesForCurrentThread() - allocated;
                measurements.Add(new { newSamples, iterations, elapsedMs, allocatedBytes = bytes,
                    bytesPerTick = (double)bytes / iterations });
                if (!newSamples) unchangedBytesPerTick = (double)bytes / iterations;
            }
            // Observe the actual 250 ms UI timer with populated buffers and no new input.
            process.Refresh();
            double cpuBefore = process.TotalProcessorTime.TotalMilliseconds;
            long allocatedBefore = GC.GetTotalAllocatedBytes();
            long privateBefore = process.PrivateMemorySize64;
            long observedAt = Stopwatch.GetTimestamp();
            _rawPollingTimer?.Start();
            await Task.Delay(10000);
            _rawPollingTimer?.Stop();
            process.Refresh();
            DevFiles.Write(DevArtifact("state/" + id + ".json"), new {
                measurements, observedMs = Stopwatch.GetElapsedTime(observedAt).TotalMilliseconds,
                cpuTimeMs = process.TotalProcessorTime.TotalMilliseconds - cpuBefore,
                allocatedBytes = GC.GetTotalAllocatedBytes() - allocatedBefore,
                privateBytesDelta = process.PrivateMemorySize64 - privateBefore,
                logicalProcessors = Environment.ProcessorCount,
                tags = _rawPollingStates.Values.Select(s => new { s.Role, s.LastDisplayedTag, s.LastLiveTag }).ToArray(),
                scope = "Actual MainForm update path, two synthetic 1K/8K streams, bounded sample buffers. Accelerated loop excludes native packet delivery; 10 s observation uses the real UI timer. Not physical HID frequency or a long-term leak test."
            });
            DevRecordState("state/" + id + ".json");
            if (_rawPollingStates.Values.Any(s => s.LastDisplayedTag != (s.Role == "Mouse" ? "1K" : "8K")
                    || s.LastLiveTag != (s.Role == "Mouse" ? "1K" : "8K")))
                throw new InvalidOperationException("Polling measurement changed the expected device rates.");
            if (requireBudget && unchangedBytesPerTick > 4096)
                throw new InvalidOperationException("Unchanged polling buffers allocate more than 4 KiB per update.");
        }
        finally
        {
            _rawPollingTimer?.Stop();
            _rawPollingStates.Clear();
            foreach (var pair in previous) _rawPollingStates.Add(pair.Key, pair.Value);
            foreach (int device in new[] { 0, 1 })
            {
                _rawPollingByController.Remove("qa-polling-controller-" + device);
                _rawLivePollingByController.Remove("qa-polling-controller-" + device);
            }
            if (timerEnabled) _rawPollingTimer?.Start();
        }
    }
}
