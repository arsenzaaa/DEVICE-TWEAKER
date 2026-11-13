// DTIMOD loading policy and the loader journal shared with startup.
#if NET8_0_OR_GREATER
#nullable disable
#endif
namespace DeviceTweakerCS
{
    public enum DriverReadyResult { Unavailable, Existing, Kdu, Service }

    public static class DriverReadyPolicy
    {
        // Used by ImodDriverContext; callbacks retain native handles and diagnostic details.
        public static DriverReadyResult Ensure(System.Func<bool> openExisting,
            System.Func<bool> loadKdu, System.Action assertIdle, System.Func<bool> loadService)
        {
            if (openExisting()) return DriverReadyResult.Existing;
            if (loadKdu()) return DriverReadyResult.Kdu;
            assertIdle(); // An unknown KDU outcome must stop SCM fallback.
            return loadService() ? DriverReadyResult.Service : DriverReadyResult.Unavailable;
        }
    }

    public static class DriverRuntime
    {
        public const string RestartMessage = "DTIMOD is unavailable.\nRestart Windows, then press CHECK.";

        public static string DescribeLoadedDriver()
        {
            DemandNativeAccess();
            // The bundled driver has no QUERY_INFO operation. Opening its device
            // does not establish the identity of an already loaded binary.
            return "device=DeviceTweakerImod2 query-info=unsupported loaded-build=unverified";
        }

        internal static void DemandNativeAccess()
        {
#if NET8_0_OR_GREATER
            QaRun.DemandWrite("DTIMOD RUNTIME COORDINATOR");
#endif
        }
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
        { return DeviceTweakerLoaderJournal.CheckJournal(record, bootId); }

        public static void AssertLoaderIdle()
        { DeviceTweakerLoaderJournal.AssertLoaderIdle(); }

        public static void BeginLoader()
        { DeviceTweakerLoaderJournal.BeginLoader(); }

        public static void CompleteLoader()
        { DeviceTweakerLoaderJournal.CompleteLoader(); }

        public void Dispose()
        {
            lease.Dispose();
        }
    }
}
