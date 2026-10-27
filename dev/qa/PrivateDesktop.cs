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
    public sealed class Result {
        public string desktop, inputBefore, inputAfter;
        public uint processId, exitCode;
        public int inputChecks;
        public bool cleanupConfirmed;
    }
    public static Result Run(string desktopName, string host, string command, string cwd, int timeoutSeconds) {
        var result = new Result { desktop = desktopName, inputBefore = InputDesktop() };
        // 0xFF deliberately excludes DESKTOP_SWITCHDESKTOP (0x100).
        IntPtr desktop = CreateDesktopW(desktopName, IntPtr.Zero, IntPtr.Zero, 0, 0xFF, IntPtr.Zero);
        if (desktop == IntPtr.Zero) throw Error("CreateDesktop");
        IntPtr job = IntPtr.Zero; ProcessInfo info = new ProcessInfo(); bool assigned = false;
        try {
            job = CreateJobObjectW(IntPtr.Zero, null);
            if (job == IntPtr.Zero) throw Error("CreateJobObject");
            var limits = new ExtendedLimits(); limits.basic.flags = 0x2000; // KILL_ON_JOB_CLOSE
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
            while (true) {
                string input = InputDesktop(); result.inputChecks++;
                if (String.Equals(input, desktopName, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("QA desktop became the input desktop.");
                uint wait = WaitForSingleObject(info.process, 250);
                if (wait == 0) break;
                if (wait != 258) throw Error("WaitForSingleObject");
                if (clock.Elapsed.TotalSeconds > timeoutSeconds) throw new TimeoutException("Private QA worker timed out.");
            }
            uint exit;
            if (!GetExitCodeProcess(info.process, out exit)) throw Error("GetExitCodeProcess");
            result.exitCode = exit;
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
            result.inputAfter = InputDesktop();
            return result;
        } finally {
            if (info.process != IntPtr.Zero && !assigned) TerminateProcess(info.process, 1);
            if (job != IntPtr.Zero) CloseHandle(job);
            if (info.thread != IntPtr.Zero) CloseHandle(info.thread);
            if (info.process != IntPtr.Zero) CloseHandle(info.process);
            CloseDesktop(desktop);
        }
    }
}
