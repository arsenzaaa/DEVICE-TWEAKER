// DTIMOD loading policy, protocol check and the loader journal.
#if NET8_0_OR_GREATER
#nullable disable
#endif
namespace DeviceTweakerCS
{
    public enum DriverReadyResult { Unavailable, Existing, Kdu, Service }

    public static class DriverReadyPolicy
    {
        // Backends own handles and native access. This sequence is also exercised on fake backends.
        public static DriverReadyResult Ensure(bool allowLoad, System.Func<bool> openCompatible,
            System.Action assertIdle, System.Func<bool> loadKdu, System.Func<bool> loadService)
        {
            if (!allowLoad) return openCompatible() ? DriverReadyResult.Existing : DriverReadyResult.Unavailable;
            assertIdle();
            if (openCompatible()) return DriverReadyResult.Existing;
            bool mapped = loadKdu();
            assertIdle(); // A timeout/unknown outcome must stop SCM fallback, even if a device appeared.
            if (openCompatible()) return mapped ? DriverReadyResult.Kdu : DriverReadyResult.Existing;
            if (loadService() && openCompatible()) return DriverReadyResult.Service;
            return DriverReadyResult.Unavailable;
        }
    }

    public static class DriverRuntime
    {
        public static readonly bool SupportsInfoQuery = false;
        public const uint ProtocolVersion = 1;
        public const uint RequiredCapabilities = 7; // physical read/write, PCI query, resident lifetime
        public const int InfoSize = 32;
        public const string RestartMessage = "DTIMOD is unavailable.\nRestart Windows, then press CHECK.";

        public static bool ValidateInfo(byte[] reply, int returned, out string description)
        {
            description = RestartMessage;
            if (reply == null || returned != InfoSize || reply.Length < InfoSize) return false;
            uint size = System.BitConverter.ToUInt32(reply, 0);
            uint protocol = System.BitConverter.ToUInt32(reply, 4);
            ulong build = System.BitConverter.ToUInt64(reply, 8);
            uint capabilities = System.BitConverter.ToUInt32(reply, 16);
            uint mode = System.BitConverter.ToUInt32(reply, 20);
            if (size != InfoSize || protocol != ProtocolVersion || build == 0
                || (capabilities & RequiredCapabilities) != RequiredCapabilities || (mode != 1 && mode != 2)
                || System.BitConverter.ToUInt64(reply, 24) != 0) return false;
            description = "build=" + build + " protocol=" + protocol + " mode=" + (mode == 2 ? "kdu" : "service") + " lifetime=resident";
            return true;
        }

        public static bool QueryInfo(System.IntPtr handle, out string description)
        {
            DemandNativeAccess();
            if (!SupportsInfoQuery)
            { description = "device=DeviceTweakerImod2 query-info=unsupported loaded-build=unverified"; return true; }
            byte[] reply = new byte[InfoSize];
            int returned;
            if (!DeviceIoControl(handle, (0x8010u << 16) | (1u << 14) | (0x815u << 2),
                System.IntPtr.Zero, 0, reply, reply.Length, out returned, System.IntPtr.Zero))
            { description = RestartMessage + " (query error " + System.Runtime.InteropServices.Marshal.GetLastWin32Error() + ")"; return false; }
            return ValidateInfo(reply, returned, out description);
        }

        internal static void DemandNativeAccess()
        {
#if NET8_0_OR_GREATER
            QaRun.DemandWrite("DTIMOD RUNTIME COORDINATOR");
#endif
        }

        [System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool DeviceIoControl(System.IntPtr handle, uint code, System.IntPtr input, int inputLength,
            byte[] output, int outputLength, out int returned, System.IntPtr overlapped);
    }

    // The journal survives an app exit; the boot ID prevents a stale timeout surviving a reboot.
    // All methods are called while holding Global\DEVICE_TWEAKER_DTIMOD_LOAD; Startup\ApplyIMOD.ps1 takes the same lock.
    public sealed class DriverOperationLease : System.IDisposable
    {
        private readonly System.Threading.Mutex mutex;
        private bool held;
        private DriverOperationLease(string name, int timeout)
        {
            mutex = new System.Threading.Mutex(false, name);
            try
            {
                try { held = mutex.WaitOne(timeout); }
                catch (System.Threading.AbandonedMutexException) { held = true; }
                if (!held) throw new System.TimeoutException("Another DTIMOD operation is still preparing the driver. Try again after it finishes.");
            }
            catch { mutex.Dispose(); throw; }
        }

        public static DriverOperationLease Acquire(string name, int timeout) { return new DriverOperationLease(name, timeout); }
        public void Dispose()
        {
            if (!held) return;
            held = false;
            try { mutex.ReleaseMutex(); } finally { mutex.Dispose(); }
        }
    }

    public sealed class DriverLoadGate : System.IDisposable
    {
        private readonly DriverOperationLease lease;
        private DriverLoadGate()
        {
            DriverRuntime.DemandNativeAccess();
            lease = DriverOperationLease.Acquire(@"Global\DEVICE_TWEAKER_DTIMOD_LOAD", 70000);
        }

        public static DriverLoadGate Enter() { return new DriverLoadGate(); }

        public static string CheckJournal(string record, string bootId)
        {
            if (string.IsNullOrEmpty(record)) return null;
            string[] parts = record.Split('|');
            System.Guid recordedBoot;
            if (parts.Length != 2 || !System.Guid.TryParse(parts[0], out recordedBoot) || parts[1] != "pending") return DriverRuntime.RestartMessage;
            if (!System.String.Equals(parts[0], bootId, System.StringComparison.OrdinalIgnoreCase)) return null;
            return "A previous DTIMOD loader has not completed normally. Restart Windows before loading again.";
        }

        private static string JournalPath
        {
            get { return System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.CommonApplicationData), "DEVICE TWEAKER", "IMOD", "loader.pending"); }
        }

        public static void AssertLoaderIdle()
        {
            DriverRuntime.DemandNativeAccess();
            if (!System.IO.File.Exists(JournalPath)) return;
            string error = CheckJournal(System.IO.File.ReadAllText(JournalPath), BootId());
            if (error != null) throw new System.InvalidOperationException(error);
            System.IO.File.Delete(JournalPath);
        }

        public static void BeginLoader()
        {
            DriverRuntime.DemandNativeAccess();
            AssertLoaderIdle();
            string bootId = BootId();
            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(JournalPath));
            // A crash even before Process.Start leaves a conservative, explicit restart requirement.
            using (System.IO.FileStream stream = new System.IO.FileStream(JournalPath, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.Read))
            {
                byte[] bytes = System.Text.Encoding.UTF8.GetBytes(bootId + "|pending");
                stream.Write(bytes, 0, bytes.Length);
                stream.Flush(true);
            }
        }

        public static void CompleteLoader()
        {
            DriverRuntime.DemandNativeAccess();
            // Only a normally observed terminal process/service start may clear this record.
            if (System.IO.File.Exists(JournalPath)) System.IO.File.Delete(JournalPath);
        }

        private static string BootId()
        {
            BootEnvironment info;
            int length;
            int status = NtQuerySystemInformation(90, out info, System.Runtime.InteropServices.Marshal.SizeOf(typeof(BootEnvironment)), out length);
            if (status < 0 || info.Identifier == System.Guid.Empty)
                throw new System.InvalidOperationException("Could not identify the Windows boot session; DTIMOD loading was not started.");
            return info.Identifier.ToString("D");
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct BootEnvironment { public System.Guid Identifier; public uint FirmwareType; public ulong Flags; }
        [System.Runtime.InteropServices.DllImport("ntdll.dll")]
        private static extern int NtQuerySystemInformation(int informationClass, out BootEnvironment information, int length, out int returned);

        public void Dispose()
        {
            lease.Dispose();
        }
    }
}
