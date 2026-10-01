namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private readonly Color _bgForm = Color.FromArgb(5, 5, 7);
    private readonly Color _bgPanel = Color.FromArgb(12, 12, 15);
    private readonly Color _bgGroup = Color.FromArgb(8, 8, 10);
    private readonly Color _fgMain = Color.FromArgb(240, 240, 240);
    private readonly Color _accent = Color.FromArgb(235, 235, 235);
    private readonly Color _accentDark = Color.FromArgb(190, 190, 190);
    private readonly Color _border = Color.FromArgb(150, 150, 150);
    private readonly Color _statusDanger = Color.FromArgb(255, 80, 60);
    private readonly Color _mutedText = Color.FromArgb(200, 200, 200);
    private readonly Color _mutedWarn = Color.FromArgb(255, 80, 60);

    private readonly Color[] _cpuPalette =
    [
        Color.FromArgb(120, 190, 255),
        Color.FromArgb(255, 90, 90),
        Color.FromArgb(120, 255, 160),
        Color.FromArgb(255, 130, 180),
    ];

    private readonly Color _cpuTextP = Color.FromArgb(255, 230, 120);
    private readonly Color _cpuTextE = Color.FromArgb(120, 200, 255);
    private readonly Color _cpuTextSmt = Color.FromArgb(255, 180, 60);

    private readonly Font _baseFont = new("Consolas", 9);
    private readonly Font _dialogFont = new("Consolas", 10.5f);
    private readonly Font _titleFont = new("Consolas", 11, FontStyle.Bold);
    private readonly Font _blockTitleFont = new("Consolas", 10, FontStyle.Bold);
    private readonly Font _brandFont = new("Consolas", 27, FontStyle.Bold);
    private readonly Font _headerFont = new("Consolas", 9.5f);
    private readonly Font _htFont = new("Consolas", 13, FontStyle.Bold);
    private readonly Font _buttonFont = new("Consolas", 11, FontStyle.Bold);
    private readonly Font _blockFont = new("Consolas", 8.5f);

    private ToolTip _copyToolTip = null!;
    private Icon? _appIcon;
    private bool _darkModeInitialized;

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _copyToolTip?.Dispose();
            _appIcon?.Dispose();
            _baseFont.Dispose();
            _dialogFont.Dispose();
            _titleFont.Dispose();
            _blockTitleFont.Dispose();
            _brandFont.Dispose();
            _headerFont.Dispose();
            _htFont.Dispose();
            _buttonFont.Dispose();
            _blockFont.Dispose();
        }

        base.Dispose(disposing);
    }
}
