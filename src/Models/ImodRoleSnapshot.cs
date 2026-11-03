using System.Collections.ObjectModel;

namespace DeviceTweakerCS;

internal sealed record ImodRoleDevice(string InstanceId, string UsbRoles, string AudioEndpoints);

// A late native read must never enumerate the live card list or polling map.
internal sealed class ImodRoleSnapshot
{
    internal IReadOnlyList<ImodRoleDevice> Devices { get; }
    internal IReadOnlyDictionary<string, string> Overrides { get; }

    internal ImodRoleSnapshot(IEnumerable<ImodRoleDevice> devices, IReadOnlyDictionary<string, string> overrides)
    {
        Devices = Array.AsReadOnly(devices.ToArray());
        Overrides = new ReadOnlyDictionary<string, string>(new Dictionary<string, string>(overrides, StringComparer.OrdinalIgnoreCase));
    }
}
