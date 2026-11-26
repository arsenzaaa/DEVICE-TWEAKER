using System.Text;

namespace DeviceTweakerCS;

// Which USB device roles (Mouse, Keyboard, Audio, Gamepad, Webcam) use which xHCI
// interrupter. Works on captured inputs only: the role snapshot of the device cards,
// the USB endpoint list and the interrupter topology read from the controller.
internal static class ImodRoleMap
{
    internal static readonly string[] Priority = DeviceTweakerImodRuntime.AdaptiveRolePriority;

    internal sealed record EndpointMapCounts(int AddressMatches, int RootPortMatches, int Classified, int FilteredGamepads);

    // Exact map: an endpoint's device address, else its root port, selects the
    // interrupters the controller assigned. Generic HID is a Gamepad only when the
    // controller's cards report a gamepad.
    internal static Dictionary<uint, HashSet<string>> MapEndpoints(
        IEnumerable<UsbEndpointInfo> endpoints,
        DeviceTweakerImodRuntime.XhciInterrupterTopology topology,
        string controllerDeviceId,
        IReadOnlySet<string> controllerRoles,
        out EndpointMapCounts counts)
    {
        Dictionary<uint, HashSet<string>> rolesByInterrupter = [];
        int addressMatches = 0;
        int rootPortMatches = 0;
        int classifiedEndpoints = 0;
        int filteredGamepadEndpoints = 0;
        foreach (UsbEndpointInfo endpoint in endpoints)
        {
            if (!HostControllerPathMatchesDeviceId(endpoint.HostControllerPath, controllerDeviceId)
                || !TryClassifyUsbRole(endpoint, out string role))
            {
                continue;
            }

            if (role.Equals("Gamepad", StringComparison.OrdinalIgnoreCase)
                && !controllerRoles.Contains("Gamepad"))
            {
                filteredGamepadEndpoints++;
                continue;
            }

            classifiedEndpoints++;
            bool mapped = false;
            if (endpoint.DeviceAddress > 0
                && topology.ByDeviceAddress.TryGetValue((uint)endpoint.DeviceAddress, out List<uint>? addressIntrs))
            {
                AddRoleToInterrupters(rolesByInterrupter, addressIntrs, role);
                addressMatches++;
                mapped = true;
            }

            if (!mapped
                && TryGetUsbRootPort(endpoint.TopologyPath, out uint rootPort)
                && topology.ByRootPort.TryGetValue(rootPort, out List<uint>? rootIntrs))
            {
                AddRoleToInterrupters(rolesByInterrupter, rootIntrs, role);
                rootPortMatches++;
            }
        }

        counts = new EndpointMapCounts(addressMatches, rootPortMatches, classifiedEndpoints, filteredGamepadEndpoints);
        return rolesByInterrupter;
    }

    internal static HashSet<string> ControllerRoleFallback(ImodRoleSnapshot roleSnapshot, string controllerDeviceId)
    {
        HashSet<string> roles = new(StringComparer.OrdinalIgnoreCase);
        foreach (ImodRoleDevice device in roleSnapshot.Devices)
        {
            if (!HostControllerPathMatchesDeviceId(device.InstanceId, controllerDeviceId) &&
                !HostControllerPathMatchesDeviceId(controllerDeviceId, device.InstanceId))
            {
                continue;
            }

            AddRolesFromText(EffectiveUsbRoles(device.UsbRoles, device.InstanceId, controllerDeviceId, roleSnapshot.Overrides), roles);
            AddRolesFromText(device.AudioEndpoints, roles);
        }

        return roles;
    }

    internal static Dictionary<string, List<string>> DisplayLabelsByRole(ImodRoleSnapshot roleSnapshot, string controllerDeviceId)
    {
        Dictionary<string, List<string>> labelsByRole = new(StringComparer.OrdinalIgnoreCase);
        foreach (ImodRoleDevice device in roleSnapshot.Devices)
        {
            if (!HostControllerPathMatchesDeviceId(device.InstanceId, controllerDeviceId) &&
                !HostControllerPathMatchesDeviceId(controllerDeviceId, device.InstanceId))
            {
                continue;
            }

            AddDisplayLabels(EffectiveUsbRoles(device.UsbRoles, device.InstanceId, controllerDeviceId, roleSnapshot.Overrides), labelsByRole);
            AddDisplayLabels(device.AudioEndpoints, labelsByRole);
        }

        return labelsByRole;
    }

    // Every device label of the controller, in role priority order, each once.
    internal static List<string> ControllerDeviceDisplayLabels(ImodRoleSnapshot roleSnapshot, string controllerDeviceId)
    {
        Dictionary<string, List<string>> labelsByRole = DisplayLabelsByRole(roleSnapshot, controllerDeviceId);
        List<string> labels = [];
        HashSet<string> emitted = new(StringComparer.OrdinalIgnoreCase);

        foreach (string role in Priority)
        {
            if (!labelsByRole.TryGetValue(role, out List<string>? roleLabels))
            {
                continue;
            }

            foreach (string label in roleLabels)
            {
                if (emitted.Add(label))
                {
                    labels.Add(label);
                }
            }
        }

        foreach (KeyValuePair<string, List<string>> pair in labelsByRole.OrderBy(static kvp => kvp.Key, StringComparer.OrdinalIgnoreCase))
        {
            foreach (string label in pair.Value)
            {
                if (emitted.Add(label))
                {
                    labels.Add(label);
                }
            }
        }

        return labels;
    }

    // Device labels for the roles of one interrupter; a role without a label is shown by name.
    internal static IEnumerable<string> DisplayRoles(
        IEnumerable<string> roles,
        IReadOnlyDictionary<string, List<string>> labelsByRole)
    {
        HashSet<string> emitted = new(StringComparer.OrdinalIgnoreCase);
        foreach (string role in OrderRoles(roles))
        {
            if (labelsByRole.TryGetValue(role, out List<string>? labels) && labels.Count > 0)
            {
                foreach (string label in labels)
                {
                    if (emitted.Add(label))
                    {
                        yield return label;
                    }
                }
            }
            else if (emitted.Add(role))
            {
                yield return role;
            }
        }
    }

    // A polling override for the controller wins over one for the device.
    internal static string EffectiveUsbRoles(string usbRoles, string instanceId, string controllerDeviceId,
        IReadOnlyDictionary<string, string> overrides)
    {
        string controllerKey = DeviceIds.Normalize(controllerDeviceId);
        if (!string.IsNullOrWhiteSpace(controllerKey)
            && overrides.TryGetValue(controllerKey, out string? controllerRoles)
            && !string.IsNullOrWhiteSpace(controllerRoles))
        {
            return MergePollingOverride(usbRoles, controllerRoles);
        }

        string deviceKey = DeviceIds.Normalize(instanceId);
        if (!string.IsNullOrWhiteSpace(deviceKey)
            && overrides.TryGetValue(deviceKey, out string? deviceRoles)
            && !string.IsNullOrWhiteSpace(deviceRoles))
        {
            return MergePollingOverride(usbRoles, deviceRoles);
        }

        return usbRoles;
    }

    internal static void AddDisplayLabels(string? text, Dictionary<string, List<string>> labelsByRole)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        foreach (string rawPart in text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            string label = rawPart.Trim();
            if (label.Length == 0 || !TryGetRoleBase(label, out string role))
            {
                continue;
            }

            if (!labelsByRole.TryGetValue(role, out List<string>? labels))
            {
                labels = [];
                labelsByRole[role] = labels;
            }

            if (!labels.Contains(label, StringComparer.OrdinalIgnoreCase))
            {
                labels.Add(label);
            }
        }
    }

    // The role of a device label, in English or Russian.
    internal static bool TryGetRoleBase(string text, out string role)
    {
        role = string.Empty;
        if (text.Contains("Mouse", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Мышь", StringComparison.OrdinalIgnoreCase))
        {
            role = "Mouse";
            return true;
        }

        if (text.Contains("Keyboard", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Клавиатура", StringComparison.OrdinalIgnoreCase))
        {
            role = "Keyboard";
            return true;
        }

        if (text.Contains("Audio", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Аудио", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Microphone", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Микрофон", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Speaker", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Headphone", StringComparison.OrdinalIgnoreCase))
        {
            role = "Audio";
            return true;
        }

        if (text.Contains("Gamepad", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Геймпад", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Joystick", StringComparison.OrdinalIgnoreCase))
        {
            role = "Gamepad";
            return true;
        }

        if (text.Contains("Camera", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Webcam", StringComparison.OrdinalIgnoreCase))
        {
            role = "Webcam";
            return true;
        }

        return false;
    }

    internal static IEnumerable<string> OrderRoles(IEnumerable<string> roles)
    {
        return roles
            .OrderBy(static role => PriorityIndex(role))
            .ThenBy(static role => role, StringComparer.OrdinalIgnoreCase);
    }

    internal static bool HostControllerPathMatchesDeviceId(string hostControllerPath, string controllerDeviceId)
    {
        string pathKey = CompactDeviceMatchText(hostControllerPath.Replace('#', '\\'));
        string idKey = CompactDeviceMatchText(controllerDeviceId);
        return pathKey.Length > 0 && idKey.Length > 0 && pathKey.Contains(idKey, StringComparison.OrdinalIgnoreCase);
    }

    // Root port: the first "PortN" path segment, else the first ".PN" marker.
    internal static bool TryGetUsbRootPort(string topologyPath, out uint rootPort)
    {
        rootPort = 0;
        if (string.IsNullOrWhiteSpace(topologyPath))
        {
            return false;
        }

        foreach (string segment in topologyPath.Split(['/', '\\'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (!segment.StartsWith("Port", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string suffix = segment[4..];
            int end = 0;
            while (end < suffix.Length && char.IsDigit(suffix[end]))
            {
                end++;
            }

            if (end > 0 && uint.TryParse(suffix[..end], out rootPort) && rootPort > 0)
            {
                return true;
            }
        }

        int index = 0;
        while (index >= 0 && index < topologyPath.Length)
        {
            index = topologyPath.IndexOf(".P", index, StringComparison.OrdinalIgnoreCase);
            if (index < 0)
            {
                break;
            }

            int start = index + 2;
            int end = start;
            while (end < topologyPath.Length && char.IsDigit(topologyPath[end]))
            {
                end++;
            }

            if (end > start && uint.TryParse(topologyPath[start..end], out rootPort) && rootPort > 0)
            {
                return true;
            }

            index = start;
        }

        return false;
    }

    internal static bool TryClassifyUsbRole(UsbEndpointInfo endpoint, out string role)
    {
        role = string.Empty;
        string cls = endpoint.InterfaceClass.Trim();
        string sub = endpoint.InterfaceSubClass.Trim();
        string proto = endpoint.InterfaceProtocol.Trim();

        if (IsUsbDescriptorValue(cls, "03") || cls.Contains("HID", StringComparison.OrdinalIgnoreCase))
        {
            if (IsUsbDescriptorValue(proto, "02") || proto.Contains("Mouse", StringComparison.OrdinalIgnoreCase))
            {
                role = "Mouse";
                return true;
            }

            if (IsUsbDescriptorValue(proto, "01") || proto.Contains("Keyboard", StringComparison.OrdinalIgnoreCase))
            {
                role = "Keyboard";
                return true;
            }

            if (!IsUsbDescriptorValue(sub, "01"))
            {
                role = "Gamepad";
                return true;
            }
        }

        if (IsUsbDescriptorValue(cls, "01") || cls.Contains("Audio", StringComparison.OrdinalIgnoreCase))
        {
            role = "Audio";
            return true;
        }

        if (IsUsbDescriptorValue(cls, "0E") || cls.Contains("Video", StringComparison.OrdinalIgnoreCase))
        {
            role = "Webcam";
            return true;
        }

        return false;
    }

    private static string MergePollingOverride(string currentRoles, string overrideRoles)
    {
        if (string.IsNullOrWhiteSpace(currentRoles) || string.IsNullOrWhiteSpace(overrideRoles))
        {
            return currentRoles;
        }

        Dictionary<string, string> pollingOverrides = new(StringComparer.OrdinalIgnoreCase);
        foreach (string rawEntry in overrideRoles.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            string entry = rawEntry.Trim();
            foreach (string role in new[] { "Mouse", "Keyboard" })
            {
                if (!RawPollingService.IsUsbRoleText(entry, role))
                {
                    continue;
                }

                string tag = entry[role.Length..].Trim();
                if (tag.Length > 0)
                {
                    pollingOverrides[role] = tag;
                }
            }
        }

        return pollingOverrides.Count > 0
            ? RawPollingService.ApplyUsbRolePollingOverrides(currentRoles, pollingOverrides)
            : currentRoles;
    }

    private static void AddRoleToInterrupters(
        Dictionary<uint, HashSet<string>> rolesByInterrupter,
        IEnumerable<uint> interrupters,
        string role)
    {
        foreach (uint intr in interrupters)
        {
            if (!rolesByInterrupter.TryGetValue(intr, out HashSet<string>? roles))
            {
                roles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                rolesByInterrupter[intr] = roles;
            }

            roles.Add(role);
        }
    }

    private static void AddRolesFromText(string? text, HashSet<string> roles)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        if (text.Contains("Mouse", StringComparison.OrdinalIgnoreCase))
        {
            roles.Add("Mouse");
        }

        if (text.Contains("Keyboard", StringComparison.OrdinalIgnoreCase))
        {
            roles.Add("Keyboard");
        }

        if (text.Contains("Audio", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Microphone", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Speaker", StringComparison.OrdinalIgnoreCase))
        {
            roles.Add("Audio");
        }

        if (text.Contains("Gamepad", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Joystick", StringComparison.OrdinalIgnoreCase))
        {
            roles.Add("Gamepad");
        }

        if (text.Contains("Camera", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("Webcam", StringComparison.OrdinalIgnoreCase))
        {
            roles.Add("Webcam");
        }
    }

    private static int PriorityIndex(string role)
    {
        for (int i = 0; i < Priority.Length; i++)
        {
            if (Priority[i].Equals(role, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return Priority.Length;
    }

    private static string CompactDeviceMatchText(string text)
    {
        StringBuilder sb = new(text.Length);
        foreach (char ch in text)
        {
            if (char.IsLetterOrDigit(ch))
            {
                sb.Append(char.ToUpperInvariant(ch));
            }
        }

        return sb.ToString();
    }

    private static bool IsUsbDescriptorValue(string text, string expectedHex)
    {
        string compact = new string(text.Where(Uri.IsHexDigit).ToArray()).TrimStart('0');
        string expected = expectedHex.TrimStart('0');
        if (compact.Length == 0)
        {
            compact = "0";
        }
        if (expected.Length == 0)
        {
            expected = "0";
        }

        return string.Equals(compact, expected, StringComparison.OrdinalIgnoreCase);
    }
}
