// Refuses to work next to the DTIMOD build that used the old device name.
#if NET8_0_OR_GREATER
#nullable disable
#endif
namespace DeviceTweakerCS
{
    public sealed class LegacyDriverService
    {
        public bool Exists;
        public bool Stopped;
        public bool Owned;
    }

    public interface IDriverRenameBackend
    {
        bool DeviceAvailable(string path);
        LegacyDriverService ReadLegacyService();
        void DeleteStoppedLegacyService();
    }

    public static class DriverRename
    {
        public const string ServiceName = "DeviceTweakerImod2";
        public const string DevicePath = @"\\.\DeviceTweakerImod2";
        public const string LegacyServiceName = "DTIMOD";
        public const string LegacyDevicePath = @"\\.\DTIMOD";
        public const string RestartMessage = "An older DTIMOD driver is loaded. Restart Windows, then press CHECK again.";

        public static void Prepare(IDriverRenameBackend backend, bool readOnly)
        {
            if (readOnly) throw new System.InvalidOperationException("QA read-only session: driver migration skipped");
            if (backend.DeviceAvailable(LegacyDevicePath)) throw new System.InvalidOperationException(RestartMessage);

        }

        public static string Complete(IDriverRenameBackend backend, bool readOnly)
        {
            if (readOnly) throw new System.InvalidOperationException("QA read-only session: driver migration skipped");
            if (!backend.DeviceAvailable(DevicePath)) throw new System.InvalidOperationException("DTIMOD is not available; legacy service was retained.");
            return "Compatibility profile: service migration disabled.";
        }

        public static IDriverRenameBackend Native(string expectedDriverPath)
        {
            return new NativeBackend(expectedDriverPath);
        }

        private sealed class NativeBackend : IDriverRenameBackend
        {
            private readonly string expectedPath;
            public NativeBackend(string path) { expectedPath = System.IO.Path.GetFullPath(path); }

            public bool DeviceAvailable(string path)
            {
                DemandNativeAccess();
                using (Microsoft.Win32.SafeHandles.SafeFileHandle handle = CreateFile(path, 0x80000000, 3, System.IntPtr.Zero, 3, 0, System.IntPtr.Zero))
                {
                    if (!handle.IsInvalid) return true;
                    int error = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
                    if (error == 2 || error == 3) return false;
                    throw new System.ComponentModel.Win32Exception(error, "Could not check DTIMOD migration state.");
                }
            }

            public LegacyDriverService ReadLegacyService()
            {
                DemandNativeAccess();
                System.IntPtr scm = OpenSCManager(null, null, 1);
                if (scm == System.IntPtr.Zero) throw LastError();
                try
                {
                    System.IntPtr service = OpenService(scm, LegacyServiceName, 4);
                    if (service == System.IntPtr.Zero)
                    {
                        int error = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
                        if (error == 1060) return new LegacyDriverService();
                        throw new System.ComponentModel.Win32Exception(error);
                    }
                    try
                    {
                        ServiceStatus status;
                        if (!QueryServiceStatus(service, out status)) throw LastError();
                        bool owned = false;
                        using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\" + LegacyServiceName))
                        {
                            if (key != null)
                            {
                                string image = System.Environment.ExpandEnvironmentVariables(System.Convert.ToString(key.GetValue("ImagePath")) ?? "").Trim('"');
                                if (image.StartsWith(@"\??\", System.StringComparison.Ordinal)) image = image.Substring(4);
                                if (image.StartsWith(@"\SystemRoot\", System.StringComparison.OrdinalIgnoreCase))
                                    image = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Windows), image.Substring(12));
                                owned = System.Convert.ToInt32(key.GetValue("Type", 0)) == 1
                                    && System.String.Equals(System.IO.Path.GetFullPath(image), expectedPath, System.StringComparison.OrdinalIgnoreCase);
                            }
                        }
                        return new LegacyDriverService { Exists = true, Stopped = status.CurrentState == 1, Owned = owned };
                    }
                    finally { CloseServiceHandle(service); }
                }
                finally { CloseServiceHandle(scm); }
            }

            public void DeleteStoppedLegacyService()
            {
                DemandNativeAccess();
                LegacyDriverService owner = ReadLegacyService();
                if (!owner.Exists) return;
                if (!owner.Owned || !owner.Stopped) throw new System.InvalidOperationException("Legacy service changed during migration; retained.");
                System.IntPtr scm = OpenSCManager(null, null, 1);
                if (scm == System.IntPtr.Zero) throw LastError();
                try
                {
                    System.IntPtr service = OpenService(scm, LegacyServiceName, 0x10000 | 4);
                    if (service == System.IntPtr.Zero)
                    {
                        if (System.Runtime.InteropServices.Marshal.GetLastWin32Error() == 1060) return;
                        throw LastError();
                    }
                    try
                    {
                        ServiceStatus state;
                        if (!QueryServiceStatus(service, out state)) throw LastError();
                        if (state.CurrentState != 1) throw new System.InvalidOperationException("Legacy DTIMOD started during migration; retained.");
                        if (!DeleteService(service) && System.Runtime.InteropServices.Marshal.GetLastWin32Error() != 1072) throw LastError();
                    }
                    finally { CloseServiceHandle(service); }
                }
                finally { CloseServiceHandle(scm); }
            }

            private static System.ComponentModel.Win32Exception LastError()
            { return new System.ComponentModel.Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error()); }

            private static void DemandNativeAccess()
            {
#if NET8_0_OR_GREATER
                if (QaRun.IsReadOnly) throw new System.InvalidOperationException("QA read-only session: native driver migration blocked");
#endif
            }

            [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
            private struct ServiceStatus
            { public uint Type, CurrentState, Controls, Win32Exit, ServiceExit, Checkpoint, WaitHint; }
            [System.Runtime.InteropServices.DllImport("kernel32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
            private static extern Microsoft.Win32.SafeHandles.SafeFileHandle CreateFile(string path, uint access, uint share, System.IntPtr security, uint creation, uint flags, System.IntPtr template);
            [System.Runtime.InteropServices.DllImport("advapi32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
            private static extern System.IntPtr OpenSCManager(string machine, string database, uint access);
            [System.Runtime.InteropServices.DllImport("advapi32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode, SetLastError = true)]
            private static extern System.IntPtr OpenService(System.IntPtr scm, string name, uint access);
            [System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError = true)]
            private static extern bool QueryServiceStatus(System.IntPtr service, out ServiceStatus state);
            [System.Runtime.InteropServices.DllImport("advapi32.dll", SetLastError = true)]
            private static extern bool DeleteService(System.IntPtr service);
            [System.Runtime.InteropServices.DllImport("advapi32.dll")]
            private static extern bool CloseServiceHandle(System.IntPtr handle);
        }
    }
}
