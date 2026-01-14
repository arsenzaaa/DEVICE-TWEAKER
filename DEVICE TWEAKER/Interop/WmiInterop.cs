using System.Management;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

internal static class WmiInterop
{
    public static List<WmiPnPDevice> GetPnPDevices()
    {
        List<WmiPnPDevice> results = [];
        try
        {
            using ManagementObjectSearcher searcher = new(
                "root\\CIMV2",
                "SELECT PNPDeviceID, Name, PNPClass, Service, Status, ConfigManagerErrorCode FROM Win32_PnPEntity");

            foreach (ManagementObject mo in searcher.Get())
            {
                string? id = mo["PNPDeviceID"] as string;
                if (string.IsNullOrWhiteSpace(id))
                {
                    continue;
                }

                results.Add(new WmiPnPDevice(
                    InstanceId: id,
                    Name: mo["Name"] as string,
                    Class: mo["PNPClass"] as string,
                    Service: mo["Service"] as string,
                    Status: mo["Status"] as string,
                    ConfigManagerErrorCode: mo["ConfigManagerErrorCode"] as int?));
            }
        }
        catch
        {
            return results;
        }

        return results;
    }

    public static WmiPnPDevice? TryGetPnPDeviceById(string instanceId)
    {
        if (string.IsNullOrWhiteSpace(instanceId))
        {
            return null;
        }

        try
        {
            string escaped = EscapeWqlString(instanceId);
            using ManagementObjectSearcher searcher = new(
                "root\\CIMV2",
                $"SELECT PNPDeviceID, Name, PNPClass, Service, Status, ConfigManagerErrorCode FROM Win32_PnPEntity WHERE PNPDeviceID='{escaped}'");

            foreach (ManagementObject mo in searcher.Get())
            {
                string? id = mo["PNPDeviceID"] as string;
                if (string.IsNullOrWhiteSpace(id))
                {
                    continue;
                }

                return new WmiPnPDevice(
                    InstanceId: id,
                    Name: mo["Name"] as string,
                    Class: mo["PNPClass"] as string,
                    Service: mo["Service"] as string,
                    Status: mo["Status"] as string,
                    ConfigManagerErrorCode: mo["ConfigManagerErrorCode"] as int?);
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    public static List<(string ControllerId, string DependentId)> GetUsbControllerDevicePairs(Func<string, string> normalizeInstanceId)
    {
        List<(string ControllerId, string DependentId)> results = [];
        try
        {
            using ManagementObjectSearcher searcher = new(
                "root\\CIMV2",
                "SELECT Antecedent, Dependent FROM Win32_USBControllerDevice");

            foreach (ManagementObject mo in searcher.Get())
            {
                string? antecedent = mo["Antecedent"] as string;
                string? dependent = mo["Dependent"] as string;
                if (string.IsNullOrWhiteSpace(antecedent) || string.IsNullOrWhiteSpace(dependent))
                {
                    continue;
                }

                string? controllerId = ExtractDeviceId(antecedent);
                string? dependentId = ExtractDeviceId(dependent);
                if (string.IsNullOrWhiteSpace(controllerId) || string.IsNullOrWhiteSpace(dependentId))
                {
                    continue;
                }

                controllerId = normalizeInstanceId(controllerId);
                dependentId = normalizeInstanceId(dependentId);

                if (string.IsNullOrWhiteSpace(controllerId) || string.IsNullOrWhiteSpace(dependentId))
                {
                    continue;
                }

                results.Add((controllerId, dependentId));
            }
        }
        catch
        {
            return results;
        }

        return results;
    }

    public static List<WmiPhysicalDisk> GetPhysicalDisks()
    {
        List<WmiPhysicalDisk> results = [];
        try
        {
            using ManagementObjectSearcher searcher = new(
                "root\\Microsoft\\Windows\\Storage",
                "SELECT FriendlyName, BusType, MediaType FROM MSFT_PhysicalDisk");

            foreach (ManagementObject mo in searcher.Get())
            {
                string name = mo["FriendlyName"] as string ?? mo["FriendlyName"]?.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = "Unknown";
                }

                ushort busType = ToUShort(mo["BusType"]);
                ushort mediaType = ToUShort(mo["MediaType"]);
                results.Add(new WmiPhysicalDisk(name, busType, mediaType));
            }
        }
        catch
        {
            return results;
        }

        return results;
    }

    private static string EscapeWqlString(string value)
    {
        return value.Replace("\\", "\\\\").Replace("'", "''");
    }

    private static string? ExtractDeviceId(string wmiRef)
    {
        Match match = Regex.Match(wmiRef, "DeviceID=\"([^\"]+)\"", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (!match.Success)
        {
            return null;
        }

        string raw = match.Groups[1].Value;
        return raw.Replace("\\\\", "\\");
    }

    private static ushort ToUShort(object? value)
    {
        if (value is null)
        {
            return 0;
        }

        try
        {
            return Convert.ToUInt16(value);
        }
        catch
        {
            return 0;
        }
    }
}
