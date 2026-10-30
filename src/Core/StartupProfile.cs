using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DeviceTweakerCS;

// profile.json read by Startup\ApplyIMOD.ps1. Register values are hex strings so
// the file stays readable; the script converts them with ConvertFrom-ImodNumber.
internal sealed class StartupProfile
{
    public const int CurrentSchemaVersion = 1;

    public int SchemaVersion { get; set; } = CurrentSchemaVersion;
    public string CreatedBy { get; set; } = string.Empty;
    public StartupFile Driver { get; set; } = new();
    public StartupFile Runtime { get; set; } = new();
    public StartupLoader Loader { get; set; } = new();
    public StartupLogging Logging { get; set; } = new();
    public StartupUsb Usb { get; set; } = new();
    public List<StartupNicItr> NicItr { get; set; } = [];

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = true,
    };

    internal string ToJson() => JsonSerializer.Serialize(this, JsonOptions) + "\n";

    internal static StartupProfile FromJson(string json)
    {
        StartupProfile profile = JsonSerializer.Deserialize<StartupProfile>(json, JsonOptions)
            ?? throw new InvalidDataException("profile.json is empty.");
        if (profile.SchemaVersion != CurrentSchemaVersion)
        {
            throw new InvalidDataException($"Unsupported profile.json schemaVersion {profile.SchemaVersion}.");
        }

        return profile;
    }

    internal static string Hex(ulong value) => "0x" + value.ToString("X", CultureInfo.InvariantCulture);

    internal static ulong ParseNumber(string? text)
    {
        string value = (text ?? string.Empty).Trim();
        return value.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
            ? ulong.Parse(value.AsSpan(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture)
            : ulong.Parse(value, NumberStyles.Integer, CultureInfo.InvariantCulture);
    }
}

internal sealed class StartupFile
{
    public string? File { get; set; }
    public string? Path { get; set; }
    public string Sha256 { get; set; } = string.Empty;
}

internal sealed class StartupLoader
{
    public StartupFile Kdu { get; set; } = new();
    public StartupFile Database { get; set; } = new();
}

internal sealed class StartupLogging
{
    public bool Enabled { get; set; } = true;
    public bool Verbose { get; set; }
}

internal sealed class StartupUsb
{
    public bool Apply { get; set; }
    public string Interval { get; set; } = "0xC8";
    public string HcsparamsOffset { get; set; } = "0x4";
    public string Rtsoff { get; set; } = "0x18";
    public List<StartupUsbController> Controllers { get; set; } = [];
}

internal sealed class StartupUsbController
{
    public string Match { get; set; } = string.Empty;
    public bool? Enabled { get; set; }
    public string? Interval { get; set; }
    public List<string>? Intervals { get; set; }
    public bool? AdaptiveRoleBinding { get; set; }
    public string? RoleIntervals { get; set; }
    public string? HcsparamsOffset { get; set; }
    public string? Rtsoff { get; set; }
}

internal sealed class StartupNicItr
{
    public string Match { get; set; } = string.Empty;
    public string Family { get; set; } = string.Empty;
    public string BaseOffset { get; set; } = "0x0";
    public string Stride { get; set; } = "0x0";
    public int Queues { get; set; }
    public int Width { get; set; }
    public string Mask { get; set; } = "0x0";
    public string OrBits { get; set; } = "0x0";
    public List<string> Values { get; set; } = [];
}
