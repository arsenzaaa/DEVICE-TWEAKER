namespace DeviceTweakerCS;

public sealed partial class MainForm
{
    private const float UiBaseDpi = 96f;
    private float _uiScale = 1f;

    private void UpdateUiScale()
    {
        float scale = 1f;
        try
        {
            int dpi = 0;
            if (IsHandleCreated)
            {
                dpi = NativeUser32.GetDpiForWindow(Handle);
            }

            if (dpi <= 0)
            {
                dpi = NativeUser32.GetDpiForSystem();
            }

            if (dpi <= 0)
            {
                dpi = (int)UiBaseDpi;
            }

            scale = dpi / UiBaseDpi;
        }
        catch
        {
            scale = 1f;
        }

        if (scale <= 0f)
        {
            scale = 1f;
        }

        _uiScale = scale;
    }

    private int UiScale(int value)
    {
        return (int)Math.Round(value * _uiScale);
    }

    private float UiScale(float value)
    {
        return value * _uiScale;
    }

    private Size UiScale(int width, int height)
    {
        return new Size(UiScale(width), UiScale(height));
    }
}
