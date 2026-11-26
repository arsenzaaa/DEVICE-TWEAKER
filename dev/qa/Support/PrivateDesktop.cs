using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

// Launcher only; it never switches desktops or synthesizes physical input.
public static class DeviceTweakerQaDesktop
{
    private delegate bool EnumWindow(IntPtr hwnd, IntPtr data);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindow callback, IntPtr data);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hwnd, uint command);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hwnd, uint msg, IntPtr w, IntPtr l);
    public static void CloseOwnedWindows(int id) {
        var owned = new System.Collections.Generic.List<IntPtr>();
        var roots = new System.Collections.Generic.List<IntPtr>();
        EnumWindows((hwnd, data) => {
            uint pid; GetWindowThreadProcessId(hwnd, out pid);
            if (pid == id) {
                var kind = new StringBuilder(256); GetClassName(hwnd, kind, 256);
                if (!kind.ToString().StartsWith("WindowsForms10.Window") && kind.ToString() != "#32770") return true;
                if (GetWindow(hwnd, 4) != IntPtr.Zero) owned.Add(hwnd);
                else {
                    var title = new StringBuilder(256); GetWindowText(hwnd, title, 256);
                    if (title.ToString() == "DEVICE TWEAKER") roots.Add(hwnd);
                }
            }
            return true;
        }, IntPtr.Zero);
        // Close modals before their owner. Never enumerate another desktop.
        foreach (IntPtr hwnd in owned) PostMessage(hwnd, 0x10, IntPtr.Zero, IntPtr.Zero);
        foreach (IntPtr hwnd in roots) PostMessage(hwnd, 0x10, IntPtr.Zero, IntPtr.Zero);
    }
    public static IntPtr FindMainWindow(int id) {
        IntPtr found = IntPtr.Zero;
        EnumWindows((hwnd, data) => {
            uint pid; GetWindowThreadProcessId(hwnd, out pid);
            if (pid != id || GetWindow(hwnd, 4) != IntPtr.Zero) return true;
            var title = new StringBuilder(256); var kind = new StringBuilder(256);
            GetWindowText(hwnd, title, 256); GetClassName(hwnd, kind, 256);
            if (title.ToString() != "DEVICE TWEAKER" || !kind.ToString().StartsWith("WindowsForms10.Window")) return true;
            found = hwnd; return false;
        }, IntPtr.Zero);
        return found;
    }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct StartupInfo {
        public int cb; public string reserved, desktop, title;
        public int x, y, width, height, charsX, charsY, fill, flags;
        public short show, reservedBytes; public IntPtr reservedData, stdin, stdout, stderr;
    }
    [StructLayout(LayoutKind.Sequential)]
    struct ProcessInfo { public IntPtr process, thread; public uint pid, tid; }
    [StructLayout(LayoutKind.Sequential)]
    struct BasicLimits {
        public long processTime, jobTime; public uint flags;
        public UIntPtr minWorking, maxWorking; public uint activeLimit;
        public UIntPtr affinity; public uint priority, scheduling;
    }
    [StructLayout(LayoutKind.Sequential)]
    struct IoCounters { public ulong a, b, c, d, e, f; }
    [StructLayout(LayoutKind.Sequential)]
    struct ExtendedLimits {
        public BasicLimits basic; public IoCounters io;
        public UIntPtr processMemory, jobMemory, peakProcess, peakJob;
    }
    [StructLayout(LayoutKind.Sequential)]
    struct JobAccounting {
        public long a, b, c, d; public uint faults, total, active, terminated;
    }
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern IntPtr CreateDesktopW(string name, IntPtr device, IntPtr mode, uint flags, uint access, IntPtr security);
    [DllImport("user32.dll")] static extern bool CloseDesktop(IntPtr handle);
    [DllImport("user32.dll")] static extern IntPtr GetThreadDesktop(uint tid);
    [DllImport("user32.dll")] static extern IntPtr GetProcessWindowStation();
    [DllImport("user32.dll", SetLastError = true)] static extern IntPtr OpenInputDesktop(uint flags, bool inherit, uint access);
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool GetUserObjectInformationW(IntPtr obj, int index, StringBuilder name, int bytes, out int needed);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool CreateProcessW(string app, StringBuilder command, IntPtr pa, IntPtr ta, bool inherit,
        uint flags, IntPtr env, string cwd, ref StartupInfo startup, out ProcessInfo info);
    [DllImport("kernel32.dll", SetLastError = true)] static extern IntPtr CreateJobObjectW(IntPtr attr, string name);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool SetInformationJobObject(IntPtr job, int kind, ref ExtendedLimits limits, int bytes);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool TerminateJobObject(IntPtr job, uint code);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool QueryInformationJobObject(IntPtr job, int kind, out JobAccounting data, int bytes, IntPtr returned);
    [DllImport("kernel32.dll", SetLastError = true)] static extern uint ResumeThread(IntPtr thread);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool SetPriorityClass(IntPtr process, uint priority);
    [DllImport("kernel32.dll")] static extern uint WaitForSingleObject(IntPtr handle, uint ms);
    [DllImport("kernel32.dll", SetLastError = true)] static extern bool GetExitCodeProcess(IntPtr process, out uint code);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);
    [DllImport("kernel32.dll")] static extern bool TerminateProcess(IntPtr handle, uint code);
    static Exception Error(string call) { return new Win32Exception(Marshal.GetLastWin32Error(), call); }
    static string Name(IntPtr handle) {
        var text = new StringBuilder(512); int needed;
        if (handle == IntPtr.Zero || !GetUserObjectInformationW(handle, 2, text, text.Capacity * 2, out needed))
            throw Error("GetUserObjectInformation");
        return text.ToString();
    }
    public static string CurrentDesktop() { return Name(GetThreadDesktop(GetCurrentThreadId())); }
    public static string InputDesktop() {
        IntPtr handle = OpenInputDesktop(0, false, 1);
        if (handle == IntPtr.Zero) return "unavailable"; // e.g. Windows lock/UAC secure desktop
        try { return Name(handle); } finally { CloseDesktop(handle); }
    }
    // Shared by independent launchers, including preparation before a GUI build.
    // A worker inside the existing private job must not acquire a second lease.
    public sealed class WorkloadLease : IDisposable {
        System.Threading.Mutex mutex;
        internal WorkloadLease(System.Threading.Mutex value) { mutex = value; }
        public void Dispose() {
            if (mutex == null) return;
            try { mutex.ReleaseMutex(); } finally { mutex.Dispose(); mutex = null; }
        }
    }
    public static WorkloadLease AcquireWorkload() {
        // Nested probes run within the already bounded private job. Contract
        // tests may temporarily clear the environment, so check the desktop.
        string current = CurrentDesktop();
        if (current.StartsWith("DeviceTweakerQa_", StringComparison.Ordinal)
            && !String.Equals(current, InputDesktop(), StringComparison.OrdinalIgnoreCase)) {
            ExtendedLimits inherited;
            if (QueryInformationJobObject(IntPtr.Zero, 9, out inherited, Marshal.SizeOf(typeof(ExtendedLimits)), IntPtr.Zero)
                && (inherited.basic.flags & 0x2200) == 0x2200
                && inherited.jobMemory.ToUInt64() > 0 && inherited.jobMemory.ToUInt64() <= (3UL << 30)) return null;
        }
        var mutex = new System.Threading.Mutex(false, @"Local\DEVICE_TWEAKER_QA_WORKLOAD");
        bool acquired = false;
        try {
            try { acquired = mutex.WaitOne(0); }
            catch (System.Threading.AbandonedMutexException) { acquired = true; }
            if (!acquired) throw new InvalidOperationException("Another QA workload is already running. Run QA and builds sequentially.");
            MemoryStatus memory = ReadMemoryStatus();
            ulong reserve = Math.Min(1UL << 30, memory.totalPhysical / 8);
            if (!HasMemoryHeadroom(memory.availablePhysical, memory.availableCommit, reserve + (256UL << 20)))
                throw new InvalidOperationException("Not enough free memory for QA. No build or worker was started.");
            return new WorkloadLease(mutex);
        } catch { if (acquired) mutex.ReleaseMutex(); mutex.Dispose(); throw; }
    }
    [StructLayout(LayoutKind.Sequential)]
    struct MemoryStatus {
        public uint length, load;
        public ulong totalPhysical, availablePhysical, totalCommit, availableCommit;
        public ulong totalVirtual, availableVirtual, extendedVirtual;
    }
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool GlobalMemoryStatusEx(ref MemoryStatus data);
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool QueryInformationJobObject(IntPtr job, int kind, out ExtendedLimits data, int bytes, IntPtr returned);
    static MemoryStatus ReadMemoryStatus() {
        var value = new MemoryStatus { length = (uint)Marshal.SizeOf(typeof(MemoryStatus)) };
        if (!GlobalMemoryStatusEx(ref value)) throw Error("GlobalMemoryStatusEx");
        return value;
    }
    public static bool HasMemoryHeadroom(ulong physical, ulong commit, ulong reserve) {
        return physical >= reserve && commit >= reserve;
    }
    static void WriteResources(string path, Result result, string status) {
        if (String.IsNullOrEmpty(path)) return;
        // Numeric counters and fixed statuses only. Persist during the run, so a
        // reboot does not erase the latest observed memory pressure.
        string json = String.Format(System.Globalization.CultureInfo.InvariantCulture,
            "{{\"status\":\"{0}\",\"jobMemoryLimitBytes\":{1},\"reserveBytes\":{2},\"peakJobCommitBytes\":{3},\"minAvailablePhysicalBytes\":{4},\"minAvailableCommitBytes\":{5},\"samples\":{6},\"cleanupConfirmed\":{7},\"injectedPressure\":{8}}}",
            status, result.jobMemoryLimitBytes, result.reserveBytes, result.peakJobCommitBytes,
            result.minAvailablePhysicalBytes, result.minAvailableCommitBytes, result.memorySamples,
            result.cleanupConfirmed ? "true" : "false", result.injectedPressure ? "true" : "false");
        System.IO.File.WriteAllText(path, json, new UTF8Encoding(false));
    }
    public sealed class Result {
        public string desktop, inputBefore, inputAfter;
        public uint processId, exitCode;
        public int inputChecks;
        public bool cleanupConfirmed, injectedPressure;
        public ulong jobMemoryLimitBytes, reserveBytes, peakJobCommitBytes;
        public ulong minAvailablePhysicalBytes, minAvailableCommitBytes;
        public int memorySamples;
    }
    public static Result Run(string desktopName, string host, string command, string cwd, int timeoutSeconds) {
        return Run(desktopName, host, command, cwd, timeoutSeconds, null);
    }
    public static Result Run(string desktopName, string host, string command, string cwd, int timeoutSeconds, string resourceReport) {
        return RunCore(desktopName, host, command, cwd, timeoutSeconds, resourceReport, false);
    }
    // Exercises the real pressure shutdown/cleanup without exhausting host RAM.
    public static Result RunPressureProbe(string desktopName, string host, string command, string cwd, int timeoutSeconds, string resourceReport) {
        return RunCore(desktopName, host, command, cwd, timeoutSeconds, resourceReport, true);
    }
    static Result RunCore(string desktopName, string host, string command, string cwd, int timeoutSeconds, string resourceReport, bool pressureProbe) {
        using (AcquireWorkload()) {
            return RunBounded(desktopName, host, command, cwd, timeoutSeconds, resourceReport, pressureProbe);
        }
    }
    static Result RunBounded(string desktopName, string host, string command, string cwd, int timeoutSeconds, string resourceReport, bool pressureProbe) {
        MemoryStatus memory = ReadMemoryStatus();
        ulong reserve = Math.Min(1UL << 30, memory.totalPhysical / 8);
        if (!HasMemoryHeadroom(memory.availablePhysical, memory.availableCommit, reserve + (256UL << 20)))
            throw new InvalidOperationException("Not enough free memory for QA. No worker was started.");
        var result = new Result { desktop = desktopName, inputBefore = InputDesktop(), reserveBytes = reserve,
            jobMemoryLimitBytes = Math.Min(3UL << 30, Math.Min(memory.totalPhysical / 4, memory.availablePhysical - reserve)),
            minAvailablePhysicalBytes = memory.availablePhysical, minAvailableCommitBytes = memory.availableCommit,
            injectedPressure = pressureProbe };
        WriteResources(resourceReport, result, "RUNNING");
        // 0xFF deliberately excludes DESKTOP_SWITCHDESKTOP (0x100).
        IntPtr desktop = CreateDesktopW(desktopName, IntPtr.Zero, IntPtr.Zero, 0, 0xFF, IntPtr.Zero);
        if (desktop == IntPtr.Zero) throw Error("CreateDesktop");
        IntPtr job = IntPtr.Zero; ProcessInfo info = new ProcessInfo(); bool assigned = false, completed = false;
        try {
            job = CreateJobObjectW(IntPtr.Zero, null);
            if (job == IntPtr.Zero) throw Error("CreateJobObject");
            var limits = new ExtendedLimits();
            limits.basic.flags = 0x2000 | 0x200 | 0x20; // KILL_ON_JOB_CLOSE | JOB_MEMORY | PRIORITY_CLASS
            limits.basic.priority = 0x4000; // BELOW_NORMAL for every process in the job
            limits.jobMemory = new UIntPtr(result.jobMemoryLimitBytes);
            if (!SetInformationJobObject(job, 9, ref limits, Marshal.SizeOf(limits))) throw Error("SetInformationJobObject");
            var startup = new StartupInfo { cb = Marshal.SizeOf(typeof(StartupInfo)),
                desktop = Name(GetProcessWindowStation()) + "\\" + desktopName, flags = 0x80 };
            // No console + suspended. Children inherit the worker desktop and environment.
            if (!CreateProcessW(host, new StringBuilder(command), IntPtr.Zero, IntPtr.Zero, false,
                0x08000004, IntPtr.Zero, cwd, ref startup, out info)) throw Error("CreateProcess");
            result.processId = info.pid;
            if (!AssignProcessToJobObject(job, info.process)) throw Error("AssignProcessToJobObject");
            assigned = true;
            if (!SetPriorityClass(info.process, 0x4000)) throw Error("SetPriorityClass"); // BELOW_NORMAL
            if (ResumeThread(info.thread) == uint.MaxValue) throw Error("ResumeThread");
            var clock = Stopwatch.StartNew();
            int pressureSamples = 0;
            while (true) {
                string input = InputDesktop(); result.inputChecks++;
                if (String.Equals(input, desktopName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("QA desktop became the input desktop.");
                uint wait = WaitForSingleObject(info.process, 250);
                if (wait != 0 && wait != 258) throw Error("WaitForSingleObject");
                memory = ReadMemoryStatus();
                if (pressureProbe) memory.availablePhysical = 0;
                result.minAvailablePhysicalBytes = Math.Min(result.minAvailablePhysicalBytes, memory.availablePhysical);
                result.minAvailableCommitBytes = Math.Min(result.minAvailableCommitBytes, memory.availableCommit);
                ExtendedLimits usage;
                if (!QueryInformationJobObject(job, 9, out usage, Marshal.SizeOf(typeof(ExtendedLimits)), IntPtr.Zero))
                    throw Error("QueryInformationJobObject memory");
                result.peakJobCommitBytes = Math.Max(result.peakJobCommitBytes, usage.peakJob.ToUInt64());
                result.memorySamples++;
                pressureSamples = HasMemoryHeadroom(memory.availablePhysical, memory.availableCommit, reserve) ? 0 : pressureSamples + 1;
                if (result.memorySamples % 4 == 1) WriteResources(resourceReport, result, "RUNNING");
                if (wait == 0) break;
                if (pressureSamples >= 8) throw new InvalidOperationException("QA stopped to preserve host memory headroom.");
                if (clock.Elapsed.TotalSeconds > timeoutSeconds) throw new TimeoutException("Private QA worker timed out.");
            }
            uint exit;
            if (!GetExitCodeProcess(info.process, out exit)) throw Error("GetExitCodeProcess");
            result.exitCode = exit;
            completed = true;
            return result;
        } finally {
            try {
                // Closing a kill-on-close job starts termination but does not wait.
                // Confirm cleanup on timeout and exceptions as well as normal exit.
                if (assigned) {
                    if (!TerminateJobObject(job, 1)) throw Error("TerminateJobObject");
                    var cleanup = Stopwatch.StartNew();
                    do {
                        JobAccounting counts;
                        if (!QueryInformationJobObject(job, 1, out counts, Marshal.SizeOf(typeof(JobAccounting)), IntPtr.Zero))
                            throw Error("QueryInformationJobObject");
                        if (counts.active == 0) { result.cleanupConfirmed = true; break; }
                        System.Threading.Thread.Sleep(25);
                    } while (cleanup.Elapsed.TotalSeconds < 5);
                    if (!result.cleanupConfirmed) throw new InvalidOperationException("QA descendants did not stop.");
                } else if (info.process != IntPtr.Zero) {
                    if (!TerminateProcess(info.process, 1)) throw Error("TerminateProcess");
                    uint wait = WaitForSingleObject(info.process, 5000);
                    if (wait == 258) throw new InvalidOperationException("Unassigned QA worker did not stop.");
                    if (wait != 0) throw Error("WaitForSingleObject");
                }
                result.inputAfter = InputDesktop();
            } finally {
                if (job != IntPtr.Zero) CloseHandle(job);
                if (info.thread != IntPtr.Zero) CloseHandle(info.thread);
                if (info.process != IntPtr.Zero) CloseHandle(info.process);
                CloseDesktop(desktop);
                WriteResources(resourceReport, result, completed && result.exitCode == 0 ? "PASS" : "FAIL");
            }
        }
    }
}
