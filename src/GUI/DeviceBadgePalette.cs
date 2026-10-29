using System.Drawing;

namespace DeviceTweakerCS;

// Category accents belong here; operation feedback has its own status colours.
// Known GPU vendors retain their identifying colours in GetGpuKindBadge.
internal static class DeviceBadgePalette
{
    internal static readonly Color Network = Color.FromArgb(98, 174, 245);
    internal static readonly Color Usb = Color.FromArgb(72, 217, 208);
    internal static readonly Color Storage = Color.FromArgb(190, 155, 235);
    internal static readonly Color Audio = Color.FromArgb(255, 155, 125);
    internal static readonly Color Unknown = Color.FromArgb(190, 190, 200);
}
