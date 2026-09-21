using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace DeviceTweakerCS;

internal enum UiLanguageCode
{
    English,
    Russian,
}

internal static partial class UiLanguage
{
    private const string LanguageEnvironmentVariable = "DEVICE_TWEAKER_LANGUAGE";
    private const string SettingsApplicationDirectory = "DEVICE TWEAKER";
    private const string SettingsFileName = "settings.json";
    private const string PreviousSettingsVendorDirectory = "arsenza";
    private const string PreviousSettingsApplicationDirectory = "DeviceTweaker";
    private const string LegacyLanguageFileName = "ui-language.txt";
    private static readonly object Sync = new();
    private static bool _initialized;
    private static UiLanguageCode _current = UiLanguageCode.English;

    public static event EventHandler? Changed;

    public static UiLanguageCode Current
    {
        get
        {
            EnsureInitialized();
            return _current;
        }
    }

    public static bool IsRussian => Current == UiLanguageCode.Russian;

    public static void Initialize()
    {
        lock (Sync)
        {
            if (_initialized)
            {
                return;
            }

            _initialized = true;
            _current = ResolveInitialLanguage();
            ApplyCulture(_current);
        }
    }

    public static void Set(UiLanguageCode language, bool persist = true)
    {
        EnsureInitialized();
        bool changed;
        lock (Sync)
        {
            changed = _current != language;
            _current = language;
            ApplyCulture(language);
            if (persist && string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(LanguageEnvironmentVariable)))
            {
                TryPersist(language);
            }
        }

        if (changed)
        {
            Changed?.Invoke(null, EventArgs.Empty);
        }
    }

    public static string Text(string? english)
    {
        string source = english ?? string.Empty;
        if (!IsRussian || source.Length == 0)
        {
            return source;
        }

        return TranslateRussian(source);
    }

    internal static bool ValidateRussianTranslationContract(out string error)
    {
        string[] samples =
        [
            "devices: Mouse 1K -> intr0 0x0 0 ns",
            "interrupters 0-1: intr0 0x0 0 ns | intr1 0xC8 50 us",
            "detail: I0 Mouse 1K/Keyboard 8K/Audio/Microphone=0x0 (0 ns) | I1=0xC8 (50 us)",
            "raw: I0:0x0 | I1:0xC8",
            "time: I0:0 ns | I1:50 us",
            "map: exact xHCI map: addressMatches=15, rootPortMatches=0, endpointTargets=4, slotTargets=3, classified=15, filteredGenericHid=10; addrMap=I1:I0",
            "devices: no exact USB role map",
            "interrupters: -",
        ];
        string[] forbidden =
        [
            "devices:", "detail:", "raw:", "time:", "map:",
            "Mouse", "Keyboard", "Microphone", "addressMatches=", "rootPortMatches=",
            "Audio", "endpointTargets=", "slotTargets=", "classified=", "filteredGenericHid=",
        ];

        foreach (string sample in samples)
        {
            string translated = TranslateRussian(sample);
            string? untranslated = forbidden.FirstOrDefault(token =>
                translated.Contains(token, StringComparison.OrdinalIgnoreCase));
            if (untranslated is not null)
            {
                error = $"token={untranslated} source={sample} translated={translated}";
                return false;
            }
        }

        (string SourceToken, string RequiredToken)[] terminologyRules =
        [
            ("CPU Affinity", "CPU Affinity"),
            ("affinity mask", "Affinity Mask"),
            ("interrupt-affinity policy", "interrupt-affinity policy"),
            ("interrupter", "interrupter"),
            ("MSI Mode", "MSI Mode"),
            ("MSI Limit", "MSI Limit"),
            ("IRQ Priority", "IRQ Priority"),
            ("NDIS Mode", "NDIS Mode"),
            ("RSS Queues", "RSS Queues"),
            ("RSS receive queues", "RSS receive queues"),
            ("Power Saving", "Power Saving"),
            ("Mouse Throttle:", "Mouse Throttle"),
            ("Raw mouse throttle", "Raw mouse throttle"),
            ("Raw Input", "Raw Input"),
            ("IMOD Value", "IMOD Value"),
            ("IMOD persistence", "IMOD persistence"),
            ("NIC ITR", "NIC ITR"),
            ("Selective Suspend", "Selective Suspend"),
            ("Reserved CPU Sets", "Reserved CPU Sets"),
            ("Hybrid CPU", "Hybrid CPU"),
            ("Dual-CCD", "Dual-CCD"),
            ("E-core", "E-core"),
            ("BOTH", "BOTH"),
        ];

        foreach (KeyValuePair<string, string> entry in Russian)
        {
            foreach ((string sourceToken, string requiredToken) in terminologyRules)
            {
                string sourcePattern = $@"(?<![A-Za-z0-9]){Regex.Escape(sourceToken)}(?![A-Za-z0-9])";
                if (Regex.IsMatch(entry.Key, sourcePattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                    && !entry.Value.Contains(requiredToken, StringComparison.OrdinalIgnoreCase))
                {
                    error = $"terminology={requiredToken} source={entry.Key} translated={entry.Value}";
                    return false;
                }
            }
        }

        (string Source, string RequiredToken)[] generatedTerminologySamples =
        [
            ("Affinity Mask: 0x4", "Affinity Mask"),
            ("Affinity (RSS mask): 0x4", "Affinity (RSS mask)"),
            ("IRQ count: 1 (MSI: Enabled)", "IRQ count"),
            ("interrupters 0-1: intr0 0x0 0 ns", "interrupters"),
            ("devices: fallback controller roles assigned to active interrupters", "interrupters"),
            ("Note: UI and affinity masks are capped at 64 LPs.", "affinity masks"),
            ("CPU 0: P-core, CCD 0, Group 0, Core 0, Local 0. CPPC: rating 120, preferred rank #1", "CPU 0: P-ядро, CCD 0, Группа 0, Ядро 0, Локальный индекс 0. CPPC: рейтинг 120, приоритетный ранг #1"),
        ];
        foreach ((string source, string requiredToken) in generatedTerminologySamples)
        {
            string translated = TranslateRussian(source);
            if (!translated.Contains(requiredToken, StringComparison.Ordinal))
            {
                error = $"terminology={requiredToken} source={source} translated={translated}";
                return false;
            }
        }

        (string SourceToken, string RequiredToken)[] typographyRules =
        [
            ("Affinity Mask", "Affinity Mask"),
            ("Affinity mask", "Affinity mask"),
            ("affinity mask", "affinity mask"),
            ("affinity masks", "affinity masks"),
            ("interrupt-affinity policy", "interrupt-affinity policy"),
            ("Interrupters", "Interrupters"),
            ("interrupters", "interrupters"),
            ("Interrupter", "Interrupter"),
            ("interrupter", "interrupter"),
            ("IRQ Count", "IRQ Count"),
            ("IRQ count", "IRQ count"),
            ("Core groups", "Core groups"),
            ("Core group", "Core group"),
            ("CCD groups", "CCD groups"),
            ("CCD group", "CCD group"),
            ("CCX groups", "CCX groups"),
            ("CCX group", "CCX group"),
            ("LP assignments", "LP assignments"),
            ("RSS Base", "RSS Base"),
            ("RSS base", "RSS base"),
            ("RSS Receive Queues", "RSS Receive Queues"),
            ("RSS receive queues", "RSS receive queues"),
            ("Power Saving", "Power Saving"),
            ("power saving", "power saving"),
            ("Raw Input", "Raw Input"),
            ("raw input", "raw input"),
            ("raw mouse input", "raw mouse input"),
            ("IMOD Mode", "IMOD Mode"),
            ("IMOD mode", "IMOD mode"),
            ("IMOD Persistence", "IMOD Persistence"),
            ("IMOD persistence", "IMOD persistence"),
            ("Selective Suspend", "Selective Suspend"),
            ("selective suspend", "selective suspend"),
            ("Reserved CPU Sets", "Reserved CPU Sets"),
            ("Reserved CPU sets", "Reserved CPU sets"),
            ("reserved CPU sets", "reserved CPU sets"),
            ("USB Host Controller", "USB Host Controller"),
            ("USB host controller", "USB host controller"),
            ("Device Mapping", "Device Mapping"),
            ("Device mapping", "Device mapping"),
            ("Role Mode", "Role Mode"),
            ("Role mode", "Role mode"),
            ("kernel CI", "kernel CI"),
        ];
        foreach (KeyValuePair<string, string> entry in Russian)
        {
            foreach ((string sourceToken, string requiredToken) in typographyRules)
            {
                string sourcePattern = $@"(?<![A-Za-z0-9]){Regex.Escape(sourceToken)}(?![A-Za-z0-9])";
                string requiredPattern = $@"(?<![A-Za-z0-9]){Regex.Escape(requiredToken)}(?![A-Za-z0-9])";
                if (Regex.IsMatch(entry.Key, sourcePattern, RegexOptions.CultureInvariant)
                    && Regex.IsMatch(entry.Value, requiredPattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant)
                    && !entry.Value.Contains(requiredToken, StringComparison.Ordinal))
                {
                    error = $"typography={requiredToken} source={entry.Key} translated={entry.Value}";
                    return false;
                }
            }
        }

        error = string.Empty;
        return true;
    }

    private static string TranslateRussian(string source)
    {
        if (source.Length == 0)
        {
            return source;
        }

        if (Russian.TryGetValue(source, out string? exact))
        {
            return exact;
        }

        string normalized = source.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
        if (normalized.Contains('\n'))
        {
            string translated = string.Join("\n", normalized.Split('\n').Select(TranslateLine));
            return source.Contains("\r\n", StringComparison.Ordinal)
                ? translated.Replace("\n", "\r\n", StringComparison.Ordinal)
                : translated;
        }

        return TranslateLine(source);
    }

    public static string Text(string english, params object?[] args)
        => string.Format(CultureInfo.CurrentCulture, Text(english), args);

    public static string RoleText(string value)
    {
        if (!IsRussian || string.IsNullOrWhiteSpace(value))
        {
            return value;
        }

        (string English, string Russian)[] roles =
        [
            ("Mouse scanning", "Мышь"),
            ("Mouse", "Мышь"),
            ("Keyboard", "Клавиатура"),
            ("Microphone", "Микрофон"),
            ("Audio", "Аудио"),
            ("Gamepad", "Геймпад"),
        ];
        foreach ((string english, string russian) in roles)
        {
            if (value.Equals(english, StringComparison.OrdinalIgnoreCase))
            {
                return russian;
            }

            if (value.StartsWith(english + " ", StringComparison.OrdinalIgnoreCase))
            {
                return russian + value[english.Length..];
            }
        }

        return value;
    }

    private static string TranslateLine(string source)
    {
        if (Russian.TryGetValue(source, out string? exact))
        {
            return exact;
        }

        Match match = AffinityMaskRegex().Match(source);
        if (match.Success)
        {
            return $"Affinity Mask: {TranslateInlineStatus(match.Groups[1].Value)}";
        }

        match = RssAffinityRegex().Match(source);
        if (match.Success)
        {
            return $"Affinity (RSS mask): {match.Groups[1].Value}";
        }

        match = IrqCountRegex().Match(source);
        if (match.Success)
        {
            return $"IRQ count: {TranslateInlineStatus(match.Groups[1].Value)}";
        }

        match = CurrentRegex().Match(source);
        if (match.Success)
        {
            return $"текущее: {TranslateInlineStatus(match.Groups[1].Value)}";
        }

        match = DefaultRegex().Match(source);
        if (match.Success)
        {
            return $"стандарт: {TranslateInlineStatus(match.Groups[1].Value)}";
        }

        match = DevicesRegex().Match(source);
        if (match.Success)
        {
            return $"устройства: {TranslateImodStructuredText(match.Groups[1].Value)}";
        }

        match = InterruptersRegex().Match(source);
        if (match.Success)
        {
            return $"interrupters{match.Groups[1].Value}: {TranslateImodStructuredText(match.Groups[2].Value)}";
        }

        match = DetailRegex().Match(source);
        if (match.Success)
        {
            return $"подробности: {TranslateImodStructuredText(match.Groups[1].Value)}";
        }

        match = RawRegex().Match(source);
        if (match.Success)
        {
            return $"исходные значения: {TranslateImodStructuredText(match.Groups[1].Value)}";
        }

        match = MapRegex().Match(source);
        if (match.Success)
        {
            return $"карта: {TranslateImodStructuredText(match.Groups[1].Value)}";
        }

        match = ProcessedRegex().Match(source);
        if (match.Success)
        {
            return $"обработано: {match.Groups[1].Value}";
        }

        match = ProgressCountRegex().Match(source);
        if (match.Success)
        {
            string prefix = match.Groups[1].Value switch
            {
                "Building device list" => "Формирование списка устройств",
                "Applying device settings" => "Применение настроек устройств",
                "Previewing device settings" => "Проверка настроек устройств",
                "Resetting device settings" => "Сброс настроек устройств",
                "Dry-run reset" => "Тестовый сброс",
                _ => match.Groups[1].Value,
            };
            return $"{prefix} ({match.Groups[2].Value}/{match.Groups[3].Value}){match.Groups[4].Value}";
        }

        match = PresetCountRegex().Match(source);
        if (match.Success)
        {
            string prefix = match.Groups[1].Value switch
            {
                "CPU preset" => "Пресет CPU",
                "System preset" => "Пресет системы",
                "Device preset" => "Пресет устройства",
                _ => match.Groups[1].Value,
            };
            return $"{prefix} ({match.Groups[2].Value}):";
        }

        match = CurrentTestDevicesRegex().Match(source);
        if (match.Success)
        {
            return $"Тестовых устройств: {match.Groups[1].Value}";
        }

        match = RealDeviceVisibilityRegex().Match(source);
        if (match.Success)
        {
            return $"Реальные устройства: видно={match.Groups[1].Value}, скрыто={match.Groups[2].Value}";
        }

        match = AffinityCapRegex().Match(source);
        if (match.Success)
        {
            return $"Примечание: UI и affinity masks ограничены {match.Groups[1].Value} LP.";
        }

        match = TimeRegex().Match(source);
        if (match.Success)
        {
            return $"время: {TranslateInlineStatus(match.Groups[1].Value)}";
        }

        match = PreviewRegex().Match(source);
        if (match.Success)
        {
            return $"предпросмотр: {TranslateInlineStatus(match.Groups[1].Value)}";
        }

        match = QueuesRegex().Match(source);
        if (match.Success)
        {
            return $"очереди: {TranslateInlineStatus(match.Groups[1].Value)}";
        }

        match = SubtitleRegex().Match(source);
        if (match.Success)
        {
            return $"альфа-версия {match.Groups[1].Value} — разработчик {match.Groups[2].Value}";
        }

        match = CpuValueRegex().Match(source);
        if (match.Success)
        {
            return $"Значение: ReservedCpuSets = {match.Groups[1].Value} | CPU: {match.Groups[2].Value}";
        }

        match = ApplyChangesDirtyRegex().Match(source);
        if (match.Success)
        {
            return $"Применить настройки (Ctrl+S) — изменено: {match.Groups[1].Value}";
        }

        match = CpuTooltipRegex().Match(source);
        if (match.Success)
        {
            return TranslateCpuTooltip(match);
        }

        string result = source;
        foreach ((string english, string russian) in InlineReplacements)
        {
            result = result.Replace(english, russian, StringComparison.Ordinal);
        }

        return TranslateBracketRoles(result);
    }

    private static string TranslateBracketRoles(string value)
    {
        int open = value.LastIndexOf('[');
        int close = value.LastIndexOf(']');
        if (open < 0 || close <= open)
        {
            return value;
        }

        string content = value[(open + 1)..close];
        string translated = string.Join(", ", content.Split(',', StringSplitOptions.TrimEntries).Select(part =>
        {
            string[] pipeParts = part.Split('|', StringSplitOptions.TrimEntries);
            return string.Join(" | ", pipeParts.Select(RoleText));
        }));
        return value[..(open + 1)] + translated + value[close..];
    }

    private static string TranslateInlineStatus(string value)
    {
        string trimmed = value.Trim();
        if (Russian.TryGetValue(trimmed, out string? translated))
        {
            return translated;
        }

        Match enterValues = EnterValuesRegex().Match(trimmed);
        if (enterValues.Success)
        {
            return $"введите одно значение или {enterValues.Groups[1].Value} значений";
        }

        foreach ((string english, string russian) in InlineReplacements)
        {
            trimmed = trimmed.Replace(english, russian, StringComparison.Ordinal);
        }

        return trimmed;
    }

    private static string TranslateImodStructuredText(string value)
    {
        string translated = TranslateInlineStatus(value);
        (string English, string Russian)[] replacements =
        [
            ("Mouse scanning", "Мышь"),
            ("Microphone", "Микрофон"),
            ("Keyboard", "Клавиатура"),
            ("Gamepad", "Геймпад"),
            ("Audio", "Аудио"),
            ("Mouse", "Мышь"),
            ("no exact USB role map", "точная карта ролей USB отсутствует"),
            ("no adaptive role map", "адаптивная карта ролей отсутствует"),
            ("no IMOD values were read", "значения IMOD не считаны"),
            ("fallback controller roles assigned to active interrupters", "резервное сопоставление ролей контроллера выполнено для активных interrupters"),
            ("exact xHCI map", "точная карта xHCI"),
            ("addressMatches=", "совпадения адресов="),
            ("rootPortMatches=", "совпадения корневых портов="),
            ("endpointTargets=", "цели конечных точек="),
            ("slotTargets=", "цели слотов="),
            ("classified=", "классифицировано="),
            ("filteredGenericHid=", "отфильтровано generic HID="),
            ("addrMap=", "карта адресов="),
            ("rootPortMap=", "карта корневых портов="),
            ("endpointTargetMap=", "карта конечных точек="),
            ("slotTargetMap=", "карта слотов="),
            ("press CHECK", "нажмите ПРОВЕРИТЬ"),
            ("mixed", "разные значения"),
            (" more", " ещё"),
        ];

        foreach ((string english, string russian) in replacements)
        {
            translated = translated.Replace(english, russian, StringComparison.OrdinalIgnoreCase);
        }

        return translated;
    }

    private static void EnsureInitialized()
    {
        if (!_initialized)
        {
            Initialize();
        }
    }

    private static UiLanguageCode ResolveInitialLanguage()
    {
        string? requested = Environment.GetEnvironmentVariable(LanguageEnvironmentVariable);
        if (TryParse(requested, out UiLanguageCode environmentLanguage))
        {
            return environmentLanguage;
        }

        try
        {
            string settingsPath = GetSettingsFilePath();
            if (TryReadSettingsLanguage(settingsPath, out UiLanguageCode savedLanguage, out bool needsRewrite))
            {
                if (needsRewrite)
                {
                    TryPersist(savedLanguage);
                }

                return savedLanguage;
            }

            string previousSettingsPath = GetPreviousSettingsFilePath();
            if (TryReadSettingsLanguage(previousSettingsPath, out UiLanguageCode previousLanguage, out _))
            {
                if (TryPersist(previousLanguage))
                {
                    TryDeleteMigratedSettings(previousSettingsPath);
                }

                return previousLanguage;
            }

            string legacyPath = GetLegacyLanguageFilePath();
            if (File.Exists(legacyPath)
                && TryParse(File.ReadAllText(legacyPath).Trim(), out UiLanguageCode legacyLanguage))
            {
                if (TryPersist(legacyLanguage))
                {
                    TryDeleteLegacyLanguageFile(legacyPath);
                }

                return legacyLanguage;
            }
        }
        catch
        {
            // A language preference must never prevent the application from starting.
        }

        return UiLanguageCode.English;
    }

    private static bool TryParse(string? value, out UiLanguageCode language)
    {
        if (string.Equals(value, "ru", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "ru-RU", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "russian", StringComparison.OrdinalIgnoreCase))
        {
            language = UiLanguageCode.Russian;
            return true;
        }

        if (string.Equals(value, "en", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "en-US", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "english", StringComparison.OrdinalIgnoreCase))
        {
            language = UiLanguageCode.English;
            return true;
        }

        language = UiLanguageCode.English;
        return false;
    }

    private static void ApplyCulture(UiLanguageCode language)
    {
        CultureInfo culture = CultureInfo.GetCultureInfo(language == UiLanguageCode.Russian ? "ru-RU" : "en-US");
        CultureInfo.CurrentUICulture = culture;
        CultureInfo.DefaultThreadCurrentUICulture = culture;
    }

    private static bool TryPersist(UiLanguageCode language)
    {
        string? tempPath = null;
        try
        {
            string path = GetSettingsFilePath();
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            tempPath = path + ".tmp-" + Guid.NewGuid().ToString("N");
            UiSettings settings = new()
            {
                Language = language == UiLanguageCode.Russian ? "ru" : "en",
            };
            JsonSerializerOptions options = new()
            {
                PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
                WriteIndented = true,
            };
            string json = JsonSerializer.Serialize(settings, options) + Environment.NewLine;
            File.WriteAllText(tempPath, json, new UTF8Encoding(false));
            File.Move(tempPath, path, overwrite: true);
            return true;
        }
        catch
        {
            // The selected language still applies for the current session.
            return false;
        }
        finally
        {
            if (tempPath is not null)
            {
                try
                {
                    File.Delete(tempPath);
                }
                catch
                {
                    // Best-effort cleanup only.
                }
            }
        }
    }

    private static bool TryReadSettingsLanguage(
        string path,
        out UiLanguageCode language,
        out bool needsRewrite)
    {
        language = UiLanguageCode.English;
        needsRewrite = false;
        if (!File.Exists(path))
        {
            return false;
        }

        string json = File.ReadAllText(path);
        UiSettings? settings = JsonSerializer.Deserialize<UiSettings>(
            json,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        if (settings is null || !TryParse(settings.Language, out language))
        {
            return false;
        }

        using JsonDocument document = JsonDocument.Parse(json);
        needsRewrite = document.RootElement.TryGetProperty("schemaVersion", out _);
        return true;
    }

    private static void TryDeleteLegacyLanguageFile(string path)
    {
        try
        {
            File.Delete(path);
            string? directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory)
                && Directory.Exists(directory)
                && !Directory.EnumerateFileSystemEntries(directory).Any())
            {
                Directory.Delete(directory);
            }
        }
        catch
        {
            // A migrated preference remains valid even if legacy cleanup is blocked.
        }
    }

    private static void TryDeleteMigratedSettings(string path)
    {
        try
        {
            File.Delete(path);
            string? applicationDirectory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(applicationDirectory)
                && Directory.Exists(applicationDirectory)
                && !Directory.EnumerateFileSystemEntries(applicationDirectory).Any())
            {
                Directory.Delete(applicationDirectory);
            }

            string? vendorDirectory = Path.GetDirectoryName(applicationDirectory);
            if (!string.IsNullOrWhiteSpace(vendorDirectory)
                && Directory.Exists(vendorDirectory)
                && !Directory.EnumerateFileSystemEntries(vendorDirectory).Any())
            {
                Directory.Delete(vendorDirectory);
            }
        }
        catch
        {
            // The migrated preference remains valid even if old folder cleanup is blocked.
        }
    }

    private static string GetSettingsFilePath()
        => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            SettingsApplicationDirectory,
            SettingsFileName);

    private static string GetPreviousSettingsFilePath()
        => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            PreviousSettingsVendorDirectory,
            PreviousSettingsApplicationDirectory,
            SettingsFileName);

    private static string GetLegacyLanguageFilePath()
        => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            SettingsApplicationDirectory,
            LegacyLanguageFileName);

    private sealed class UiSettings
    {
        public string Language { get; set; } = "en";
    }

    private static readonly IReadOnlyDictionary<string, string> Russian =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["DEVICE TWEAKER"] = "DEVICE TWEAKER",
            ["APPLY"] = "ПРИМЕНИТЬ",
            ["AUTO-OPTIMIZATION"] = "АВТООПТИМИЗАЦИЯ",
            ["REFRESH"] = "ОБНОВИТЬ",
            ["RESTORE"] = "ВОССТАНОВЛЕНИЕ",
            ["RESET WINDOWS DEFAULT"] = "СБРОС WINDOWS",
            ["DELETE SELECTED"] = "УДАЛИТЬ",
            ["RESTORE LAST"] = "ПОСЛЕДНИЙ БЭКАП",
            ["RESTORE SELECTED"] = "ВЫБРАННЫЙ БЭКАП",
            ["CANCEL"] = "ОТМЕНА",
            ["YES"] = "ДА",
            ["NO"] = "НЕТ",
            ["OK"] = "ОК",
            ["DETAILS"] = "ПОДРОБНОСТИ",
            ["HIDE DETAILS"] = "СКРЫТЬ ПОДРОБНОСТИ",
            ["OPEN LOG"] = "ОТКРЫТЬ ЛОГ",
            ["CLOSE"] = "ЗАКРЫТЬ",
            ["SKIP"] = "ПРОПУСТИТЬ",
            ["CONTINUE"] = "ПРОДОЛЖИТЬ",
            ["RETRY"] = "ПОВТОРИТЬ",
            ["SET"] = "ЗАДАТЬ",
            ["SAVE"] = "СОХРАНИТЬ",
            ["DELETE"] = "УДАЛИТЬ",
            ["CHECK"] = "ПРОВЕРИТЬ",
            ["COPY"] = "КОПИРОВАТЬ",
            ["COPIED"] = "СКОПИРОВАНО",
            ["COMPLETED"] = "ЗАВЕРШЕНО",
            ["CHANGES APPLIED"] = "ИЗМЕНЕНИЯ ПРИМЕНЕНЫ",
            ["APPLIED WITH WARNINGS"] = "ПРИМЕНЕНО С ПРЕДУПРЕЖДЕНИЯМИ",
            ["NOT APPLIED"] = "НЕ ПРИМЕНЕНО",
            ["PARTIALLY APPLIED"] = "ПРИМЕНЕНО ЧАСТИЧНО",
            ["SKIPPED"] = "ПРОПУЩЕНО",
            ["FAILED"] = "ОШИБКА",
            ["SUCCESS"] = "УСПЕШНО",
            ["WARNING"] = "ПРЕДУПРЕЖДЕНИЕ",
            ["ERROR"] = "ОШИБКА",
            ["DIAGNOSTICS & SYSTEM LOG"] = "ДИАГНОСТИКА И СИСТЕМНЫЙ ЛОГ",
            ["RESTORE SYSTEM SETTINGS"] = "ВОССТАНОВЛЕНИЕ НАСТРОЕК",
            ["CONFIRMATION"] = "ПОДТВЕРЖДЕНИЕ",
            ["INFORMATION"] = "ИНФОРМАЦИЯ",
            ["Enabled"] = "Включено",
            ["Disabled"] = "Выключено",
            ["On"] = "Вкл",
            ["Off"] = "Выкл",
            ["True"] = "Да",
            ["False"] = "Нет",
            ["Unknown"] = "Неизвестно",
            ["Unavailable"] = "Недоступно",
            ["unavailable"] = "недоступно",
            ["reading..."] = "чтение...",
            ["loading driver..."] = "загрузка драйвера...",
            ["invalid input"] = "некорректное значение",
            ["not set"] = "не задано",
            ["unlocked"] = "без ограничения",
            ["MachineDefault"] = "MachineDefault",
            ["Undefined"] = "Не задан",
            ["AllCloseProcessors"] = "AllCloseProcessors",
            ["OneCloseProcessor"] = "OneCloseProcessor",
            ["AllProcessorsInMachine"] = "AllProcessorsInMachine",
            ["SpecifiedProcessors"] = "SpecifiedProcessors",
            ["Windows Default"] = "По умолчанию Windows",
            ["0x0 (Windows Default)"] = "0x0 (По умолчанию Windows)",
            ["NVMe/SATA storage uses Windows Multi-Queue steering. Pinned core affinity is intentionally disabled to ensure maximum SSD speed and low latency."] = "Дисковые контроллеры NVMe/SATA управляются стеком Windows Multi-Queue. Фиксация на ядрах отключена для сохранения максимальной скорости и низкой задержки.",
            ["Display/HDMI audio shares PCIe bus with the GPU and uses Windows default interrupt steering (0x0)."] = "Звук HDMI/DisplayPort делит шину PCIe с видеокартой и использует стандартную маршрутизацию прерываний Windows (0x0).",
            ["SpreadMessagesAcrossAllProcessors"] = "SpreadMessagesAcrossAllProcessors",
            ["High"] = "Высокий",
            ["Normal"] = "Обычный",
            ["Low"] = "Низкий",
            ["Devices"] = "Устройства",
            ["Interrupters"] = "Interrupters",
            ["BOTH"] = "BOTH",
            ["RSS"] = "RSS",
            ["IRQ"] = "IRQ",
            ["CPU Affinity"] = "CPU Affinity",
            ["ALL"] = "ВСЕ",
            ["MOUSE"] = "МЫШЬ",
            ["KEYBOARD"] = "КЛАВИАТУРА",
            ["GAMEPAD"] = "ГЕЙМПАД",
            ["NETWORK"] = "СЕТЬ",
            ["AUDIO"] = "ЗВУК",
            ["STORAGE"] = "НАКОПИТЕЛИ",
            ["NONE"] = "СНЯТЬ ВСЁ",
            ["P-CORES"] = "P-ЯДРА",
            ["MODIFIED"] = "ИЗМЕНЕНО",
            ["[ ALL ]"] = "[ ВСЕ ]",
            ["[ MOUSE ]"] = "[ МЫШЬ ]",
            ["[ KEYBOARD ]"] = "[ КЛАВИАТУРА ]",
            ["[ GAMEPAD ]"] = "[ ГЕЙМПАД ]",
            ["[ USB ]"] = "[ USB ]",
            ["[ GPU ]"] = "[ GPU ]",
            ["[ NETWORK ]"] = "[ СЕТЬ ]",
            ["[ AUDIO ]"] = "[ ЗВУК ]",
            ["[ STORAGE ]"] = "[ НАКОПИТЕЛИ ]",
            ["[ NONE ]"] = "[ СНЯТЬ ВСЁ ]",
            ["[ P-CORES ]"] = "[ P-ЯДРА ]",
            ["[ CCD0 ]"] = "[ CCD0 ]",
            ["[ CCD1 ]"] = "[ CCD1 ]",
            ["[ MODIFIED ]"] = "[ ИЗМЕНЕНО ]",
            ["[ COPY REGISTRY ]"] = "[ СКОПИРОВАТЬ РЕЕСТР ]",
            ["Filter devices... (Ctrl+F)"] = "Фильтр устройств... (Ctrl+F)",
            ["NO MATCHING DEVICES"] = "УСТРОЙСТВА НЕ НАЙДЕНЫ",
            ["No devices match the current filter criteria."] = "Нет устройств, соответствующих заданному фильтру.",
            ["COPY REGISTRY"] = "СКОПИРОВАТЬ РЕЕСТР",
            ["Click to copy full registry path to clipboard"] = "Нажмите, чтобы скопировать путь в реестре в буфер обмена",
            ["Registry path copied to clipboard."] = "Путь в реестре скопирован в буфер обмена.",
            ["Select all logical cores"] = "Выбрать все логические ядра",
            ["Clear all cores"] = "Снять выбор со всех ядер",
            ["Select physical P-cores only"] = "Выбрать только физические P-ядра",
            ["Select CCD0 cores only"] = "Выбрать только ядра CCD0",
            ["MSI Mode:"] = "MSI Mode:",
            ["MSI Limit:"] = "MSI Limit:",
            ["IRQ Priority:"] = "IRQ Priority:",
            ["Policy:"] = "Policy:",
            ["Policy (RSS base)"] = "Policy (RSS base)",
            ["NDIS Mode:"] = "NDIS Mode:",
            ["RSS Queues:"] = "RSS Queues:",
            ["Power Saving:"] = "Power Saving:",
            ["Mouse Throttle:"] = "Mouse Throttle:",
            ["Mode:"] = "Режим:",
            ["IMOD Value:"] = "IMOD Value:",
            ["NIC ITR:"] = "NIC ITR:",
            ["Reserved CPU Sets"] = "Reserved CPU Sets",
            ["Value: ReservedCpuSets = not set"] = "Значение: ReservedCpuSets = не задано",
            ["Hyper-Threading"] = "Hyper-Threading",
            ["Hybrid CPU"] = "Hybrid CPU",
            ["Dual-CCD"] = "Dual-CCD",
            ["Sandbox"] = "Песочница",
            ["TEST ADMIN"] = "ТЕСТ-ПАНЕЛЬ",
            ["Test CPU Topology"] = "Тестовая топология CPU",
            ["CPU TOPOLOGY"] = "ТОПОЛОГИЯ CPU",
            ["DEVICES"] = "УСТРОЙСТВА",
            ["SCENARIO LAB"] = "СЦЕНАРИИ",
            ["RESULT DIALOGS"] = "ОКНА РЕЗУЛЬТАТОВ",
            ["Current section"] = "Текущий раздел",
            ["Jump to section"] = "Перейти к разделу",
            ["Test CPU mode: ACTIVE"] = "Тестовый режим CPU: ВКЛЮЧЁН",
            ["Test CPU mode: OFF"] = "Тестовый режим CPU: ВЫКЛЮЧЕН",
            ["CPU name:"] = "Название CPU:",
            ["CPU preset:"] = "Пресет CPU:",
            ["System preset:"] = "Пресет системы:",
            ["Total logical processors:"] = "Логических процессоров:",
            ["SMT status:"] = "Состояние SMT:",
            ["Hyper-Threading status:"] = "Состояние Hyper-Threading:",
            ["CPPC ratings:"] = "Рейтинги CPPC:",
            ["Group counts:"] = "Количество групп:",
            ["Core groups:"] = "Core groups:",
            ["CCD groups:"] = "CCD groups:",
            ["CCX groups:"] = "CCX groups:",
            ["LP assignments (manual):"] = "LP assignments (вручную):",
            ["Core group"] = "Core group",
            ["CCD group"] = "CCD group",
            ["CCX group"] = "CCX group",
            ["E-core"] = "E-core",
            ["E-Core"] = "E-Core",
            ["How to use: set group counts, then assign each LP to Core/CCD/CCX groups. Tick E-core where needed."] = "Укажите количество групп, затем назначьте логические процессоры (LP) в группы Core/CCD/CCX. При необходимости отметьте E-core.",
            ["Tip: full PC presets are above. This section is for adding/removing individual fake devices and temporarily hiding real devices."] = "Подсказка: готовые пресеты ПК находятся выше. В этом разделе можно добавлять и удалять тестовые устройства, а также временно скрывать реальные.",
            ["Test Devices"] = "Тестовые устройства",
            ["Test options"] = "Параметры тестирования",
            ["Add fake device"] = "Добавление тестового устройства",
            ["Enable test devices"] = "Включить тестовые устройства",
            ["Show test devices only"] = "Показывать только тестовые устройства",
            ["Sandbox dry-run (no registry writes)"] = "Песочница без записи в реестр",
            ["Device preset:"] = "Пресет устройства:",
            ["Name:"] = "Название:",
            ["Kind:"] = "Тип:",
            ["USB roles:"] = "Роли USB:",
            ["Storage tag:"] = "Метка накопителя:",
            ["IRQ count:"] = "IRQ count:",
            ["MSI status:"] = "Состояние MSI:",
            ["USB Selective Suspend:"] = "USB Selective Suspend:",
            ["NIC power saving:"] = "NIC power saving:",
            ["USB has devices"] = "На USB есть устройства",
            ["Integrated GPU (iGPU)"] = "Integrated GPU (iGPU)",
            ["Options:"] = "Параметры:",
            ["ADD PRESET"] = "ДОБАВИТЬ ПРЕСЕТ",
            ["ADD FAKE DEVICE"] = "ДОБАВИТЬ УСТРОЙСТВО",
            ["UPDATE SELECTED"] = "ОБНОВИТЬ ВЫБРАННОЕ",
            ["REMOVE SELECTED"] = "УДАЛИТЬ ВЫБРАННОЕ",
            ["CLEAR ALL"] = "ОЧИСТИТЬ ВСЁ",
            ["Real device visibility"] = "Видимость реальных устройств",
            ["Visible real devices"] = "Видимые реальные устройства",
            ["Hidden real devices"] = "Скрытые реальные устройства",
            ["HIDE SELECTED"] = "СКРЫТЬ ВЫБРАННОЕ",
            ["UNHIDE SELECTED"] = "ПОКАЗАТЬ ВЫБРАННОЕ",
            ["CLEAR HIDDEN"] = "ОЧИСТИТЬ СКРЫТЫЕ",
            ["Scenario Lab"] = "Лаборатория сценариев",
            ["Scenario name:"] = "Название сценария:",
            ["Operation:"] = "Операция:",
            ["Failure injection:"] = "Имитация ошибки:",
            ["Initial state device:"] = "Устройство начального состояния:",
            ["MSI state:"] = "Состояние MSI:",
            ["CPU Affinity state:"] = "Состояние CPU Affinity:",
            ["RSS / power:"] = "RSS / питание:",
            ["RSS state:"] = "Состояние RSS:",
            ["Driver state:"] = "Состояние драйвера:",
            ["Include USB IMOD"] = "Включить USB IMOD",
            ["Limit (absent/0..2048):"] = "MSI Limit (нет/0..2048):",
            ["Priority:"] = "IRQ Priority:",
            ["Mask:"] = "Affinity Mask:",
            ["Base (absent/int):"] = "RSS Base (нет/число):",
            ["Queues:"] = "RSS Queues:",
            ["Power:"] = "Power Saving:",
            ["APPLY INITIAL STATE"] = "ПРИМЕНИТЬ НАЧАЛЬНОЕ СОСТОЯНИЕ",
            ["APPLY TEST"] = "ПРИМЕНИТЬ ТЕСТ",
            ["APPLY CPU TOPOLOGY"] = "ПРИМЕНИТЬ ТОПОЛОГИЮ CPU",
            ["RUN SCENARIO"] = "ЗАПУСТИТЬ СЦЕНАРИЙ",
            ["SAVE JSON"] = "СОХРАНИТЬ JSON",
            ["LOAD JSON"] = "ЗАГРУЗИТЬ JSON",
            ["RUN CORE MATRIX"] = "ЗАПУСТИТЬ ОСНОВНУЮ МАТРИЦУ",
            ["Result Dialog Gallery"] = "Галерея окон результатов",
            ["RESULT: SUCCESS"] = "РЕЗУЛЬТАТ: УСПЕХ",
            ["RESULT: PARTIAL"] = "РЕЗУЛЬТАТ: ЧАСТИЧНО",
            ["RESULT: FAILED"] = "РЕЗУЛЬТАТ: ОШИБКА",
            ["RESULT: STRESS"] = "РЕЗУЛЬТАТ: СТРЕСС-ТЕСТ",
            ["PROMPT: IMOD"] = "ДИАЛОГ: IMOD",
            ["PROMPT: BACKUP"] = "ДИАЛОГ: БЭКАП",
            ["PROMPT: RESTORE"] = "ДИАЛОГ: ВОССТАНОВЛЕНИЕ",
            ["PROMPT: INFO"] = "ДИАЛОГ: ИНФО",
            ["TEST MODE INFO"] = "ИНФО ТЕСТОВОГО РЕЖИМА",
            ["DEVICE TWEAKER is running in TEST ADMIN mode.\n\nAll hardware adjustments and registry modifications are simulated and completely safe."] =
                "DEVICE TWEAKER запущен в тестовом режиме администратора.\n\nВсе аппаратные настройки и изменения реестра симулируются и полностью безопасны.",
            ["RESET TO REAL"] = "ВЕРНУТЬ РЕАЛЬНЫЕ ДАННЫЕ",
            ["DISABLE TEST MODE"] = "ОТКЛЮЧИТЬ ТЕСТОВЫЙ РЕЖИМ",
            ["Required while simulated devices are active."] = "Обязательно при активной эмуляции устройств.",
            ["LOAD"] = "ЗАГРУЗИТЬ",
            ["LOAD FULL"] = "ЗАГРУЗИТЬ ВСЁ",
            ["Auto"] = "Авто",
            ["absent"] = "отсутствует",
            ["unset"] = "не задано",
            ["preserve"] = "сохранить",
            ["Manual (current)"] = "Вручную (текущее)",
            ["Manual / current"] = "Вручную (текущее)",
            ["Manual (custom)"] = "Вручную (своё)",
            ["Manual / custom"] = "Вручную (своё)",
            ["Preview the production result dialogs without running APPLY or AUTO-OPTIMIZATION. Partial also exercises the Vulnerable Driver Blocklist confirm when the blocklist is enabled."] =
                "Предпросмотр рабочих диалогов результатов без вызова ПРИМЕНИТЬ или АВТООПТИМИЗАЦИИ. Вариант ЧАСТИЧНО также открывает подтверждение для отключения Vulnerable Driver Blocklist, если список блокировки включён.",
            ["Release regression"] = "Регрессия релиза",
            ["Apply"] = "Применение",
            ["SafeReset"] = "SafeReset",
            ["Backup"] = "Бэкап",
            ["Restore"] = "Восстановление",
            ["None"] = "Нет",
            ["CpuTopologyUnavailable"] = "CpuTopologyUnavailable",
            ["CpuMapUnavailable"] = "CpuMapUnavailable",
            ["BackupAccessDenied"] = "BackupAccessDenied",
            ["RegistryWriteDenied"] = "RegistryWriteDenied",
            ["UsbPowerWriteFailed"] = "UsbPowerWriteFailed",
            ["WmiTimeout"] = "WmiTimeout",
            ["KduMissing"] = "KduMissing",
            ["KduTimeout"] = "KduTimeout",
            ["KduExitCode"] = "KduExitCode",
            ["KduDeviceUnavailable577"] = "KduDeviceUnavailable577",
            ["ImodReadFailed"] = "ImodReadFailed",
            ["ImodWriteFailed"] = "ImodWriteFailed",
            ["ImodVerificationMismatch"] = "ImodVerificationMismatch",
            ["NicItrReadFailed"] = "NicItrReadFailed",
            ["NicItrWriteFailed"] = "NicItrWriteFailed",
            ["CorruptBackup"] = "CorruptBackup",
            ["RestoreWriteFailed"] = "RestoreWriteFailed",
            ["RollbackFailed"] = "RollbackFailed",
            ["RefreshFailed"] = "RefreshFailed",
            ["Speakers"] = "Динамики",
            ["Optional. Format: 0=120, 1=110, 2=100 or just 120,110,100. Empty = CPPC off in test mode."] = "Необязательно. Формат: 0=120, 1=110, 2=100 или просто 120,110,100. Пусто — CPPC выключен в тестовом режиме.",
            ["Optional fake hardware ID. Example: PCI\\VEN_10EC&DEV_8125\\TEST for NIC ITR profile preview."] = "Необязательный тестовый ID оборудования. Пример: PCI\\VEN_10EC&DEV_8125\\TEST для предпросмотра профиля NIC ITR.",
            ["Optional fake IRQ count shown in device blocks. Empty/Auto = role-based default."] = "Необязательное тестовое количество IRQ в карточках устройств. Пусто/Авто — значение по роли.",
            ["Fake SelectiveSuspendEnabled for USB test devices (Selective Suspend picker)."] = "Тестовое значение SelectiveSuspendEnabled для USB-устройств (поле Selective Suspend).",
            ["IRQ count must be Auto or a number from 0 to 64."] = "IRQ count должен быть Auto или числом от 0 до 64.",
            ["Select a test device first."] = "Сначала выберите тестовое устройство.",
            ["MSI Limit must be absent/0 or 1..2048."] = "MSI Limit должен отсутствовать, быть равен 0 или находиться в диапазоне 1–2048.",
            ["Affinity mask must be hexadecimal, for example 0x3."] = "Affinity mask должна быть шестнадцатеричной, например 0x3.",
            ["Empty group is not allowed."] = "Пустая группа недопустима.",
            ["Scenario could not be loaded."] = "Не удалось загрузить сценарий.",
            ["CPPC ratings are invalid."] = "Рейтинги CPPC указаны неверно.",
            ["Sandbox matrix passed."] = "Матрица песочницы пройдена.",
            ["One or more sandbox scenarios failed."] = "Один или несколько сценариев песочницы завершились ошибкой.",
            ["SANDBOX MATRIX"] = "МАТРИЦА ПЕСОЧНИЦЫ",
            ["Current"] = "Текущее",
            ["Default"] = "По умолчанию",
            ["DEVICE"] = "УСТРОЙСТВО",
            ["VALUE"] = "ЗНАЧЕНИЕ",
            ["DELAY"] = "ЗАДЕРЖКА",
            ["QUEUE"] = "ОЧЕРЕДЬ",
            ["Operation"] = "Операция",
            ["Automatic backup"] = "Автоматический бэкап",
            ["AUTOMATIC BACKUP"] = "АВТОМАТИЧЕСКИЙ БЭКАП",
            ["BACKUP"] = "БЭКАП",
            ["BACKUP RESTORE"] = "ВОССТАНОВЛЕНИЕ БЭКАПА",
            ["DEVICE SETTINGS"] = "НАСТРОЙКИ УСТРОЙСТВ",
            ["AUTO PLAN"] = "ПЛАН АВТООПТИМИЗАЦИИ",
            ["USB POWER"] = "ПИТАНИЕ USB",
            ["USB IMOD"] = "USB IMOD",
            ["REFRESH UI"] = "ОБНОВЛЕНИЕ ИНТЕРФЕЙСА",
            ["Reserved CPU sets"] = "Reserved CPU sets",
            ["Raw mouse throttle"] = "Raw mouse throttle",
            ["IMOD persistence"] = "IMOD persistence",
            ["USB Selective Suspend power plan"] = "USB Selective Suspend в плане питания",
            ["Applied"] = "Применено",
            ["Built"] = "Сформирован",
            ["Selected settings were applied and saved."] = "Выбранные настройки применены и сохранены.",
            ["Please reboot your PC to finish applying the changes."] = "Перезагрузите компьютер, чтобы завершить применение изменений.",
            ["Applied steps were saved."] = "Применённые шаги сохранены.",
            ["Applied steps were saved. One optional step was not completed."] = "Выполненные изменения сохранены. Один необязательный этап не завершён.",
            ["DTIMOD was unavailable. USB IMOD was not changed."] = "DTIMOD недоступен. Значение USB IMOD не изменено.",
            ["Windows blocked DTIMOD via Vulnerable Driver Blocklist. USB IMOD was not changed."] =
                "Windows заблокировал DTIMOD через Vulnerable Driver Blocklist. USB IMOD не изменён.",
            ["Windows blocked DTIMOD via Vulnerable Driver Blocklist."] =
                "Windows заблокировал DTIMOD через Vulnerable Driver Blocklist.",
            ["Windows Defender blocked DTIMOD. USB IMOD was not changed."] =
                "Windows Defender заблокировал DTIMOD. USB IMOD не изменён.",
            ["Windows Defender blocked DTIMOD."] = "Windows Defender заблокировал DTIMOD.",
            ["DTIMOD.sys blocked by Vulnerable Driver Blocklist"] =
                "DTIMOD.sys заблокирован Vulnerable Driver Blocklist",
            ["DTIMOD.sys loader blocked by Vulnerable Driver Blocklist"] =
                "загрузчик DTIMOD.sys заблокирован Vulnerable Driver Blocklist",
            ["DTIMOD.sys blocked by Windows Defender"] = "DTIMOD.sys заблокирован Windows Defender",
            ["OPEN BACKUP"] = "ОТКРЫТЬ БЭКАП",
            ["BACKUP"] = "БЭКАП",
            ["BACKUPS"] = "БЭКАПЫ",
            ["LOGS"] = "ЛОГИ",
            ["DISABLE"] = "ОТКЛЮЧИТЬ",
            ["KEEP"] = "ОСТАВИТЬ",
            ["VULNERABLE DRIVER BLOCKLIST"] = "VULNERABLE DRIVER BLOCKLIST",
            ["USB IMOD was blocked by Vulnerable Driver Blocklist.\n\nDisable Vulnerable Driver Blocklist so DTIMOD can load after reboot?\n\nWarning: this can break Faceit, The Finals, and similar anti-cheats.\nAfter reboot, run CHECK or AUTO-OPTIMIZATION again to apply IMOD."] =
                "USB IMOD заблокирован Vulnerable Driver Blocklist.\n\nОтключить Vulnerable Driver Blocklist, чтобы DTIMOD загрузился после перезагрузки?\n\nВнимание: это может сломать Faceit, The Finals и похожие античиты.\nПосле перезагрузки снова нажмите ПРОВЕРИТЬ или АВТООПТИМИЗАЦИЮ, чтобы применить IMOD.",
            ["Failed to disable Vulnerable Driver Blocklist.\nSee the session log for details."] =
                "Не удалось отключить Vulnerable Driver Blocklist.\nПодробности сохранены в логе.",
            ["Vulnerable Driver Blocklist was disabled.\n\nPlease reboot your PC, then run CHECK or AUTO-OPTIMIZATION again to apply USB IMOD."] =
                "Vulnerable Driver Blocklist отключён.\n\nПерезагрузите компьютер, затем снова нажмите ПРОВЕРИТЬ или АВТООПТИМИЗАЦИЮ, чтобы применить USB IMOD.",
            ["DTIMOD can be blocked by Vulnerable Driver Blocklist, Windows driver signature protection, antivirus, or anti-cheats."] =
                "DTIMOD может быть заблокирован Vulnerable Driver Blocklist, защитой подписи драйверов Windows, антивирусом или античитами.",
            ["USB IMOD tuning is available for detected XHCI controller(s).\n\nDTIMOD can be blocked by Vulnerable Driver Blocklist, Windows driver signature protection, antivirus, or anti-cheats.\n\nApply it during AUTO-OPTIMIZATION?"] =
                "Для обнаруженных XHCI-контроллеров доступна настройка USB IMOD.\n\nDTIMOD может быть заблокирован Vulnerable Driver Blocklist, защитой подписи драйверов Windows, антивирусом или античитами.\n\nПрименить его во время АВТООПТИМИЗАЦИИ?",
            ["USB IMOD tuning is available for detected XHCI controller(s).\n\nDTIMOD driver access can be blocked by Windows security features, antivirus software, or anti-cheats.\n\nApply it during AUTO-OPTIMIZATION?"] =
                "Для обнаруженных XHCI-контроллеров доступна настройка USB IMOD.\n\nДоступ к драйверу DTIMOD может быть заблокирован защитой Windows, антивирусом или античитами.\n\nПрименить его во время АВТООПТИМИЗАЦИИ?",
            ["APPLIED WITH WARNINGS"] = "ПРИМЕНЕНО С ПРЕДУПРЕЖДЕНИЯМИ",
            ["WARNING"] = "ПРЕДУПРЕЖДЕНИЕ",
            ["HIDE DETAILS"] = "СКРЫТЬ ПОДРОБНОСТИ",
            ["Some changes may already have been applied."] = "Некоторые изменения уже могли быть применены.",
            ["Open DETAILS before rebooting or applying more changes."] = "Перед перезагрузкой или новыми изменениями откройте ПОДРОБНОСТИ.",
            ["No settings were changed."] = "Настройки не изменены.",
            ["The operation was cancelled."] = "Операция отменена.",
            ["The selected backup is invalid."] = "Выбранный бэкап повреждён или имеет неверный формат.",
            ["The backup file could not be created."] = "Не удалось создать файл бэкапа.",
            ["backup could not be created; no settings were changed"] = "не удалось создать бэкап; настройки не изменены",
            ["backup could not be created; changes were not applied"] = "не удалось создать бэкап; изменения не применены",
            ["Settings were restored. Please reboot your PC."] = "Настройки восстановлены. Перезагрузите компьютер.",
            ["Some restore steps may already have run."] = "Некоторые этапы восстановления уже могли выполниться.",
            ["Restore did not finish. Open DETAILS before rebooting."] = "Восстановление не завершено. Перед перезагрузкой откройте ПОДРОБНОСТИ.",
            ["RESET WINDOWS DEFAULT was cancelled because the rollback backup failed."] = "Сброс Windows отменён: не удалось создать бэкап для отката.",
            ["Supported tweaks were reset."] = "Поддерживаемые твики сброшены.",
            ["Please reboot your PC to restore runtime hardware defaults."] = "Перезагрузите компьютер, чтобы восстановить стандартные параметры оборудования.",
            ["RESET WINDOWS DEFAULT finished with errors. Some settings may still be active."] = "Сброс Windows завершён с ошибками. Некоторые настройки могут оставаться активными.",
            ["Choose how to restore DEVICE TWEAKER settings."] = "Выберите способ восстановления настроек DEVICE TWEAKER.",
            ["RESTORE LAST uses the newest snapshot. RESTORE SELECTED uses the highlighted snapshot."] = "ПОСЛЕДНИЙ восстанавливает свежий снимок, ВЫБРАННЫЙ — выделенный снимок.",
            ["RESET WINDOWS DEFAULT restores supported settings to Windows defaults."] = "СБРОС WINDOWS возвращает поддерживаемые настройки к значениям Windows по умолчанию.",
            ["ORIGINAL STATE is protected and never pruned."] = "ИСХОДНОЕ СОСТОЯНИЕ защищено и не удаляется автоматически.",
            ["No backup snapshot was found."] = "Снимки бэкапов не найдены.",
            ["ORIGINAL STATE"] = "ИСХОДНОЕ СОСТОЯНИЕ",
            ["pre-apply"] = "до применения",
            ["pre-auto"] = "до автооптимизации",
            ["pre-reset"] = "до сброса",
            ["pre-reset-tweaks"] = "до сброса твиков",
            ["NVIDIA/AMD video driver not detected."] = "Драйвер видеокарты NVIDIA/AMD не обнаружен.",
            ["Install the GPU driver and press REFRESH."] = "Установите драйвер видеокарты и нажмите ОБНОВИТЬ.",
            ["This tool must be run as Administrator (it writes to HKLM registry)."] = "Программу необходимо запустить от имени администратора: она изменяет реестр HKLM.",
            ["Right-click the EXE and choose 'Run as administrator'."] = "Нажмите правой кнопкой по EXE и выберите «Запуск от имени администратора».",
            ["DEVICE TWEAKER encountered a critical error and must close."] = "В DEVICE TWEAKER произошла критическая ошибка. Программа будет закрыта.",
            ["A crash report was saved in the logs folder. No further changes will be applied."] = "Отчёт о сбое сохранён в папке logs. Новые изменения применяться не будут.",
            ["DEVICE TWEAKER — CRITICAL ERROR"] = "DEVICE TWEAKER — КРИТИЧЕСКАЯ ОШИБКА",
            ["Reads HKLM:\\System\\CurrentControlSet\\Control\\Session Manager\\Kernel\\ReservedCpuSets"] = "Читает HKLM:\\System\\CurrentControlSet\\Control\\Session Manager\\Kernel\\ReservedCpuSets",
            ["Registry: HKLM\\System\\CurrentControlSet\\Control\\Session Manager\\kernel"] = "Реестр: HKLM\\System\\CurrentControlSet\\Control\\Session Manager\\kernel",
            ["current: varies by interrupter"] = "текущее: зависит от interrupter",
            ["current: unavailable"] = "текущее: недоступно",
            ["current: reading..."] = "текущее: чтение...",
            ["current: loading driver..."] = "текущее: загрузка драйвера...",
            ["devices: driver unavailable"] = "устройства: драйвер недоступен",
            ["devices: reading..."] = "устройства: чтение...",
            ["devices: unavailable"] = "устройства: недоступно",
            ["devices: invalid IMOD input"] = "устройства: некорректное значение IMOD",
            ["detected USB device roles"] = "обнаруженные роли USB-устройств",
            ["current: unavailable (Windows 11 22H2+)"] = "текущее: недоступно (Windows 11 22H2+)",
            ["0 = unlocked"] = "0 = без ограничения",
            ["(0 = unlocked)"] = "(0 = без ограничения)",
            ["varies by interrupter"] = "зависит от interrupter",
            ["Mouse:"] = "Мышь:",
            ["Keyboard:"] = "Клавиатура:",
            ["Audio:"] = "Аудио:",
            ["Microphone:"] = "Микрофон:",
            ["Gamepad:"] = "Геймпад:",
            ["No HID roles"] = "роли HID не обнаружены",
            ["unsupported"] = "не поддерживается",
            ["read failed"] = "ошибка чтения",
            ["write failed"] = "ошибка записи",
            ["driver check skipped (test)"] = "проверка драйвера пропущена (тест)",
            ["driver load failed"] = "драйвер не загрузился",
            ["driver unavailable"] = "драйвер недоступен",
            ["test preview"] = "тестовый предпросмотр",
            ["backup failed"] = "ошибка бэкапа",
            ["applying..."] = "применение...",
            ["saved for startup"] = "сохранено для автозапуска",
            ["save failed"] = "ошибка сохранения",
            ["off"] = "выключено",
            ["on"] = "включено",
            ["driver blocked"] = "драйвер заблокирован",
            ["kernel CI blocked"] = "заблокировано kernel CI",
            ["signature blocked"] = "подпись драйвера заблокирована",
            ["driver not loaded — press CHECK"] = "драйвер не загружен — нажмите ПРОВЕРИТЬ",
            ["admin required"] = "требуются права администратора",
            ["No additional details."] = "Дополнительные сведения отсутствуют.",
            ["single value for controller"] = "одно значение для контроллера",
            ["values by interrupter index"] = "значения по индексам interrupter",
            ["input: hex / list / roles"] = "ввод: hex / список / роли",
            ["time: reading..."] = "время: чтение...",
            ["time: unavailable"] = "время: недоступно",
            ["EXE FOLDER"] = "ПАПКА EXE",
            ["APPDATA"] = "APPDATA",
            ["AUTO BACKUP"] = "БЭКАП АВТООПТИМИЗАЦИИ",
            ["backup"] = "бэкап",
            ["Where should DEVICE TWEAKER save the pre-auto backup?"] = "Где сохранить бэкап перед автооптимизацией?",
            ["EXE FOLDER = portable backup next to the app."] = "ПАПКА EXE — переносимый бэкап рядом с программой.",
            ["APPDATA = user profile backup that survives app folder cleanup."] = "APPDATA — бэкап в профиле пользователя, не зависящий от папки программы.",
            ["SKIP = do not create an additional pre-auto backup."] = "ПРОПУСТИТЬ — не создавать дополнительный бэкап перед автооптимизацией.",
            ["The protected ORIGINAL STATE snapshot is always retained."] = "Защищённый снимок ИСХОДНОЕ СОСТОЯНИЕ сохраняется всегда.",
            ["Close (X) = cancel AUTO-OPTIMIZATION."] = "Закрыть (X) — отменить АВТООПТИМИЗАЦИЮ.",
            ["TEST DEVICE (no registry writes)"] = "ТЕСТОВОЕ УСТРОЙСТВО (без записи в реестр)",
            ["Protection: Wi-Fi settings preserved (no changes)"] = "Защита: настройки Wi-Fi не изменяются",
            ["Type: Storage controller"] = "Тип: контроллер накопителя",
            ["Type: NVMe storage controller"] = "Тип: контроллер NVMe-накопителя",
            ["Type: SATA/AHCI controller"] = "Тип: контроллер SATA/AHCI",
            ["Refresh devices (F5 / Ctrl+R)"] = "Обновить устройства (F5 / Ctrl+R)",
            ["Apply changes (Ctrl+S)"] = "Применить настройки (Ctrl+S)",
            ["Auto-optimization (Ctrl+O)"] = "Автооптимизация (Ctrl+O)",
            ["Restore settings (Ctrl+Z)"] = "Восстановить настройки (Ctrl+Z)",
            ["Message Signaled Interrupts (MSI/MSI-X). Replaces legacy pin-based line IRQs with direct in-band PCIe memory writes to the local APIC. Eliminates interrupt sharing, lowers latency to sub-microsecond levels, and prevents DPC spikes in games."] =
                "Режим Message Signaled Interrupts (MSI/MSI-X). Заменяет устаревшие строчные прерывания (line-based IRQ) на прямую запись в память контроллера прерываний (APIC) через шину PCIe. Устраняет разделение IRQ между устройствами, снижает задержку прерываний до субмикросекунд и предотвращает DPC-статтеры в играх.",
            ["MessageNumberLimit (registry: MessageSignaledInterruptProperties). Controls the maximum number of MSI-X interrupt vectors the device driver can allocate. 0 = unlimited (hardware default). Recommended: 0 for GPUs and modern NICs."] =
                "Параметр MessageNumberLimit (в реестре: MessageSignaledInterruptProperties). Задаёт максимальное количество векторов прерываний MSI-X, доступных драйверу устройства. 0 = без ограничений (аппаратное значение по умолчанию). Рекомендуется: 0 для видеокарт и современных сетевых карт.",
            ["DevicePriority (registry: Affinity Policy). Controls Windows kernel interrupt servicing priority relative to other hardware devices. Setting High ensures that critical gaming inputs (mouse, keyboard) and GPU interrupts are processed ahead of secondary devices during heavy CPU load."] =
                "Параметр DevicePriority (в реестре: Affinity Policy). Определяет приоритет обработки прерываний устройства ядром Windows относительно другого оборудования. Значение «Высокий» (High) гарантирует, что прерывания мыши, клавиатуры и видеокарты обрабатываются в первую очередь даже при высокой нагрузке на систему.",
            ["DevicePolicy (registry: Affinity Policy). Defines how the Windows HAL routes device interrupts across processors:\n• SpecifiedProcessors: strictly binds interrupts to the selected Affinity Mask\n• MachineDefault: default Windows steering via BIOS/ACPI tables\n• AllCloseProcessors / OneCloseProcessor: routes to near NUMA node cores\n• SpreadMessagesAcrossAllProcessors: distributes MSI-X messages across all cores."] =
                "Параметр DevicePolicy (в реестре: Affinity Policy). Определяет политику маршрутизации прерываний устройства ядром Windows HAL:\n• SpecifiedProcessors: строгая привязка прерываний к выбранной Affinity Mask\n• MachineDefault: стандартное распределение Windows на основе таблиц BIOS/ACPI\n• AllCloseProcessors / OneCloseProcessor: маршрутизация по ядрам текущего NUMA-узла\n• SpreadMessagesAcrossAllProcessors: распределение сообщений MSI-X по всем ядрам.",
            ["Interrupt Affinity Mask (AssignmentSetOverride). Strictly routes hardware interrupt service routines (ISRs) and deferred procedure calls (DPCs) to the selected CPU logical cores."] =
                "Маска привязки прерываний Affinity Mask (AssignmentSetOverride). Направляет обработчики аппаратных прерываний (ISR) и отложенные вызовы (DPC) строго на выбранные логические процессоры (ядра).",
            ["Shows the number of allocated interrupt vectors and current interrupt delivery mode. MSI/MSI-X indicates a modern dedicated interrupt vector. Line-based indicates legacy INTx sharing with other PCI devices."] =
                "Отображает количество выделенных векторов прерываний и текущий режим доставки. MSI/MSI-X указывает на современный выделенный вектор без коллизий. Строчный режим (Line-based) указывает на устаревший общий IRQ INTx, делящий линию с другими устройствами.",
            ["Net type: NetAdapterCx"] = "Тип сети: NetAdapterCx",
            ["Copied"] = "Скопировано",
            ["RSS assigns receive queues to CPUs. IRQ writes interrupt-affinity policy. BOTH writes both settings when the adapter and driver support RSS."] = "RSS распределяет receive queues по CPU. IRQ задаёт interrupt-affinity policy. BOTH применяет обе настройки, если сетевой адаптер и драйвер поддерживают RSS.",
            ["Number of RSS receive queues to configure. The adapter and driver determine the supported range. The plan starts at the selected base CPU."] = "Количество настраиваемых RSS receive queues. Допустимый диапазон определяют сетевой адаптер и драйвер. Распределение начинается с выбранного RSS base CPU.",
            ["Load DTIMOD.sys (IMOD driver) and re-read current NIC ITR values."] = "Загрузить DTIMOD.sys (IMOD driver) и повторно прочитать текущие значения NIC ITR.",
            ["Supported IMOD input:"] = "Поддерживаемый формат IMOD:",
            ["XHCI applies one value to all interrupters on this USB host controller. Device mapping assigns detected devices to their interrupter. Interrupter mode applies values by interrupter index."] = "XHCI применяет одно значение ко всем interrupters этого USB host controller. Device mapping назначает обнаруженные устройства соответствующим interrupters. Interrupter mode применяет значения по индексу interrupter.",
            ["Shows how the selected IMOD mode interprets the IMOD Value field."] = "Показывает, как выбранный IMOD mode интерпретирует поле IMOD Value.",
            ["Apply the current IMOD configuration and read back hardware values."] = "Применить текущую конфигурацию IMOD и прочитать фактические значения оборудования.",
            ["Load DTIMOD.sys (IMOD driver) and re-read current hardware IMOD values."] = "Загрузить DTIMOD.sys (драйвер IMOD) и повторно прочитать текущие значения IMOD оборудования.",
            ["Role mode applies values to the detected interrupter for each USB role."] = "Role mode применяет значения к обнаруженному interrupter каждой USB-роли.",
            ["Click to copy the default IMOD interval into the input field."] = "Нажмите, чтобы скопировать стандартный интервал IMOD в поле ввода.",
            ["Wi-Fi settings are preserved by DEVICE TWEAKER."] = "DEVICE TWEAKER не изменяет настройки Wi-Fi.",
            ["Same idea as Device Manager → Power Management → 'Allow the computer to turn off this device to save power'."] = "Аналог параметра «Разрешить отключение этого устройства для экономии энергии» в диспетчере устройств.",
            ["For USB this disables Selective Suspend on the controller + root hubs and the power-plan USB SS setting."] = "Для USB отключает Selective Suspend у контроллера и корневых USB-концентраторов, а также USB SS в текущем плане электропитания.",
            ["Unchecked = Disabled. Applied with APPLY / AUTO. A reboot may be required."] = "Флажок снят — функция выключена. Применяется через ПРИМЕНИТЬ или АВТООПТИМИЗАЦИЮ. Может потребоваться перезагрузка.",
            ["Device Manager → Power Management → 'Allow the computer to turn off this device to save power'."] = "Параметр «Разрешить отключение этого устройства для экономии энергии» в диспетчере устройств.",
            ["Unchecked sets PnPCapabilities bit 0x08 (do-not-turn-off) and clears MSPower_DeviceEnable on the wired NIC."] = "Снятый флажок устанавливает бит PnPCapabilities 0x08 и отключает MSPower_DeviceEnable у проводного сетевого адаптера.",
            ["Applied with APPLY / AUTO. A reboot may be required."] = "Применяется через ПРИМЕНИТЬ или АВТООПТИМИЗАЦИЮ. Может потребоваться перезагрузка.",
            ["Failed to write ReservedCpuSets."] = "Не удалось записать ReservedCpuSets.",
            ["See the log for details."] = "Подробности сохранены в логе.",
            ["DIAGNOSTIC LOGGING IS UNAVAILABLE"] = "ДИАГНОСТИЧЕСКИЙ ЛОГ НЕДОСТУПЕН",
            ["DEVICE TWEAKER can continue, but this session may not contain enough information for troubleshooting."] = "DEVICE TWEAKER может продолжить работу, но данных этой сессии может быть недостаточно для диагностики.",
            ["The driver check did not complete."] = "Проверка драйвера не завершена.",
            ["DTIMOD driver was unavailable."] = "Драйвер DTIMOD недоступен.",
            ["NIC ITR DRIVER"] = "ДРАЙВЕР NIC ITR",
            ["Backup saved."] = "Бэкап сохранён.",
            ["No backup was created."] = "Бэкап не создан.",
            ["No backup files were found in the EXE folder or APPDATA."] = "Файлы бэкапов не найдены ни в папке EXE, ни в APPDATA.",
            ["Backup restored."] = "Бэкап восстановлен.",
            ["Please reboot your PC to finish applying restored settings."] = "Перезагрузите компьютер, чтобы завершить применение восстановленных настроек.",
            ["Restore did not finish. Any completed rollback steps were preserved."] = "Восстановление не завершено. Выполненные этапы отката сохранены.",
            ["The backup could not be created. No changes were made."] = "Не удалось создать бэкап. Настройки не изменены.",
            ["NIC ITR was cancelled before any changes were made."] = "Настройка NIC ITR отменена до внесения изменений.",
            ["RAW INPUT THROTTLE"] = "ОГРАНИЧЕНИЕ RAW INPUT",
            ["The setting was not changed."] = "Настройка не изменена.",
            ["The setting could not be updated."] = "Не удалось обновить настройку.",
            ["APPLY was cancelled because the automatic backup failed."] = "Применение отменено: не удалось создать автоматический бэкап.",
            ["APPLY preview completed."] = "Предпросмотр применения завершён.",
            ["Sandbox dry-run is ON (no registry changes)."] = "Тестовый режим включён: реестр не изменялся.",
            ["Applied steps were saved. One or more steps were not completed."] = "Выполненные изменения сохранены. Один или несколько этапов не завершены.",
            ["USB IMOD tuning is available for detected XHCI controller(s)."] = "Для обнаруженных XHCI-контроллеров доступна настройка USB IMOD.",
            ["DTIMOD driver access can be blocked by Windows security features, antivirus software, or anti-cheats."] = "Доступ к драйверу DTIMOD могут блокировать средства безопасности Windows, антивирус или античит.",
            ["Apply it during AUTO-OPTIMIZATION?"] = "Применить USB IMOD во время АВТООПТИМИЗАЦИИ?",
            ["USB IMOD TUNING"] = "НАСТРОЙКА USB IMOD",
            ["AUTO-OPTIMIZATION was cancelled because the automatic backup failed."] = "Автооптимизация отменена: не удалось создать автоматический бэкап.",
            ["Original-state backup"] = "Бэкап исходного состояния",
            ["the initial recovery snapshot could not be created; changes were not applied"] = "не удалось создать исходный снимок восстановления; изменения не применены",
            ["AUTO-OPTIMIZATION was cancelled because the original-state backup failed."] = "Автооптимизация отменена: не удалось сохранить исходное состояние.",
            ["IMOD apply was cancelled because the automatic backup failed."] = "Применение IMOD отменено: не удалось создать автоматический бэкап.",
            ["No eligible USB IMOD target for SET."] = "Подходящий USB-контроллер для применения IMOD не найден.",
            ["IMOD applied."] = "IMOD применён.",
            ["USB IMOD completed with warnings or failed steps."] = "USB IMOD завершён с предупреждениями или ошибками.",
            ["IMOD delete"] = "Удаление IMOD",
            ["IMOD reset to defaults."] = "IMOD сброшен к стандартным значениям.",
            ["Startup script removed. Reboot your PC to unload DTIMOD.sys from memory if it was loaded."] = "Скрипт автозапуска удалён. Перезагрузите компьютер, чтобы выгрузить DTIMOD.sys из памяти, если драйвер был загружен.",
            ["IMOD delete finished with errors. Some persistence files may still remain. If the driver was loaded, reboot to unload DTIMOD.sys from memory."] = "Удаление IMOD завершено с ошибками. Некоторые файлы автозапуска могли сохраниться. Если драйвер был загружен, перезагрузите компьютер для его выгрузки.",
            ["Dry-run reset complete (UI only)."] = "Тестовый сброс завершён только в интерфейсе.",
            ["No registry / IMOD startup files were changed."] = "Реестр и файлы автозапуска IMOD не изменялись.",
            ["Disable sandbox dry-run to perform a real RESET WINDOWS DEFAULT."] = "Отключите тестовый режим, чтобы выполнить настоящий СБРОС WINDOWS.",
            ["RESET WINDOWS DEFAULT preview finished with errors."] = "Предпросмотр сброса Windows завершён с ошибками.",
            ["RESET WINDOWS DEFAULT PREVIEW"] = "ПРЕДПРОСМОТР СБРОСА WINDOWS",
            ["(affinity masks not supported on SSD/HDD)"] = "(affinity masks не поддерживаются для накопителей)",
            ["Scanning devices..."] = "Сканирование устройств...",
            ["Clearing device list..."] = "Очистка списка устройств...",
            ["Enumerating devices..."] = "Поиск устройств...",
            ["Building reserved CPU sets..."] = "Формирование reserved CPU sets...",
            ["Laying out devices..."] = "Размещение устройств...",
            ["Updating IMOD display..."] = "Обновление данных IMOD...",
            ["Updating IMOD / IRQ display..."] = "Обновление данных IMOD / IRQ...",
            ["Updating IRQ counts..."] = "Обновление количества IRQ...",
            ["Refreshing devices..."] = "Обновление устройств...",
            ["Applying changes..."] = "Применение изменений...",
            ["Previewing APPLY..."] = "Проверка перед применением...",
            ["Applying USB selective suspend..."] = "Применение USB selective suspend...",
            ["Skipping USB selective suspend..."] = "Пропуск USB selective suspend...",
            ["Applying USB IMOD..."] = "Применение USB IMOD...",
            ["Skipping USB IMOD..."] = "Пропуск USB IMOD...",
            ["Running AUTO-OPTIMIZATION..."] = "Выполняется автооптимизация...",
            ["Planning AUTO-OPTIMIZATION..."] = "Подготовка автооптимизации...",
            ["Windows limits how often background apps receive raw mouse input."] = "Windows ограничивает частоту получения raw mouse input фоновыми приложениями.",
            ["This reduces the message load created by throttling and coalescing input."] = "Это снижает нагрузку от сообщений за счёт ограничения частоты и объединения событий ввода.",
            ["Mice with a polling rate from 1 to 8 kHz can create a high message load."] = "Мыши с частотой опроса от 1 до 8 кГц могут создавать высокую нагрузку сообщениями.",
            ["Foreground apps keep the full input rate."] = "Активные приложения сохраняют полную частоту ввода.",
            ["Background raw input listeners receive a lower message rate."] = "Фоновые обработчики raw input получают сообщения с пониженной частотой.",
            ["The setting writes RawMouseThrottleDuration to HKCU\\Control Panel\\Mouse."] = "Настройка записывает RawMouseThrottleDuration в HKCU\\Control Panel\\Mouse.",
            ["The accepted DWORD range is from 1 to 20 milliseconds."] = "Допустимый диапазон DWORD — от 1 до 20 миллисекунд.",
            ["A larger value produces a lower rate for background listeners."] = "Чем больше значение, тем ниже частота для фоновых обработчиков.",
            ["20 milliseconds is approximately 50 Hz. 8 milliseconds is approximately 125 Hz. 1 millisecond is approximately 1 kHz."] = "20 миллисекунд — примерно 50 Гц, 8 миллисекунд — примерно 125 Гц, 1 миллисекунда — примерно 1 кГц.",
            ["The setting first appeared in the Windows 11 22H2 preview build 22621.1928 with KB5027303."] = "Настройка впервые появилась в предварительной сборке Windows 11 22H2 22621.1928 с KB5027303.",
            ["It was included in the July 2023 update KB5028185."] = "Она вошла в обновление KB5028185 за июль 2023 года.",
            ["This setting is not available on Windows 10 or an older unpatched Windows 11 22H2 build."] = "Настройка недоступна в Windows 10 и старых сборках Windows 11 22H2 без необходимых обновлений.",
            ["This estimate is based on the controller PCI ID."] = "Оценка основана на PCI ID контроллера.",
            ["It does not represent measured latency or the negotiated device link."] = "Она не является измеренной задержкой и не показывает согласованную скорость подключения устройства.",
            ["USB hubs connected downstream are not included."] = "Подключённые далее по цепочке USB-концентраторы не учитываются.",
        };

    private static readonly (string English, string Russian)[] InlineReplacements =
    [
        ("Power Saving=Enabled", "Power Saving=Включено"),
        ("Power Saving=Disabled", "Power Saving=Выключено"),
        ("Hybrid CPU", "Hybrid CPU"),
        ("Dual-CCD", "Dual-CCD"),
        ("Class:", "Класс:"),
        ("Registry:", "Реестр:"),
        ("Polling:", "Опрос:"),
        ("Net type:", "Тип сети:"),
        ("Type:", "Тип:"),
        ("Audio endpoints:", "Аудиовыходы:"),
        ("Path:", "Путь:"),
        ("Topology:", "Топология:"),
        ("Estimated topology:", "Предполагаемая топология:"),
        ("Controller:", "Контроллер:"),
        ("Platform:", "Платформа:"),
        ("PCI ID:", "PCI ID:"),
        ("Controller capability:", "Возможности контроллера:"),
        ("Power Saving:", "Power Saving:"),
        ("USB selective suspend:", "USB Selective Suspend:"),
        ("Raw input throttle:", "Raw Input Throttle:"),
        ("(locked)", "(зафиксировано)"),
        ("test profile ", "тестовый профиль "),
        ("preview ", "предпросмотр "),
        ("all queues", "все очереди"),
        ("Delay ", "Задержка "),
        (" active", " активна"),
        ("invalid", "некорректно"),
        ("Unknown", "Неизвестно"),
        ("Protection:", "Защита:"),
        ("Backup files deleted:", "Удалено файлов бэкапов:"),
        ("Preparing RESET WINDOWS DEFAULT...", "Подготовка сброса Windows..."),
        ("Previewing RESET WINDOWS DEFAULT...", "Предпросмотр сброса Windows..."),
        ("Running RESET WINDOWS DEFAULT...", "Выполняется сброс Windows..."),
        ("Resetting device settings", "Сброс настроек устройств"),
        ("Resetting reserved CPU sets...", "Сброс Reserved CPU Sets..."),
        ("Removing IMOD persistence...", "Удаление IMOD Persistence..."),
        ("Updating IRQ counts...", "Обновление IRQ..."),
        ("Ready", "Готово"),
        ("MSI: Enabled", "MSI: Включено"),
        ("MSI: Disabled", "MSI: Выключено"),
        ("MSI: Unknown", "MSI: Неизвестно"),
        ("Current: ", "Текущее: "),
        ("Default: ", "По умолчанию: "),
        ("not available", "недоступно"),
        ("not detected", "не обнаружено"),
        ("no devices", "нет устройств"),
        ("no changes", "нет изменений"),
        ("processed", "обработано"),
        ("reading...", "чтение..."),
        ("Windows Default", "По умолчанию Windows"),
        ("NVMe storage controller", "контроллер NVMe-накопителя"),
        ("SATA/AHCI controller", "контроллер SATA/AHCI"),
        ("Storage controller", "контроллер накопителя"),
        ("Mouse:", "Мышь:"),
        ("Keyboard:", "Клавиатура:"),
    ];

    [GeneratedRegex("^Affinity Mask: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex AffinityMaskRegex();

    [GeneratedRegex("^Affinity \\(RSS mask\\): (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex RssAffinityRegex();

    [GeneratedRegex("^IRQ Count: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex IrqCountRegex();

    [GeneratedRegex("^current: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex CurrentRegex();

    [GeneratedRegex("^default: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex DefaultRegex();

    [GeneratedRegex("^devices: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex DevicesRegex();

    [GeneratedRegex("^interrupters( [^:]+)?: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex InterruptersRegex();

    [GeneratedRegex("^detail: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex DetailRegex();

    [GeneratedRegex("^raw: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex RawRegex();

    [GeneratedRegex("^map: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex MapRegex();

    [GeneratedRegex("^(\\d+) processed$", RegexOptions.CultureInvariant)]
    private static partial Regex ProcessedRegex();

    [GeneratedRegex("^(Building device list|Applying device settings|Previewing device settings|Resetting device settings|Dry-run reset) \\((\\d+)/(\\d+)\\)(\\.\\.\\.)?$", RegexOptions.CultureInvariant)]
    private static partial Regex ProgressCountRegex();

    [GeneratedRegex("^(CPU preset|System preset|Device preset) \\((\\d+)\\):$", RegexOptions.CultureInvariant)]
    private static partial Regex PresetCountRegex();

    [GeneratedRegex("^Current test devices: (\\d+)$", RegexOptions.CultureInvariant)]
    private static partial Regex CurrentTestDevicesRegex();

    [GeneratedRegex("^Real device visibility: visible=(\\d+), hidden=(\\d+)$", RegexOptions.CultureInvariant)]
    private static partial Regex RealDeviceVisibilityRegex();

    [GeneratedRegex("^Note: UI and affinity masks are capped at (\\d+) LPs\\.$", RegexOptions.CultureInvariant)]
    private static partial Regex AffinityCapRegex();

    [GeneratedRegex("^time: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex TimeRegex();

    [GeneratedRegex("^preview: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex PreviewRegex();

    [GeneratedRegex("^queues: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex QueuesRegex();

    [GeneratedRegex("^enter 1 or (\\d+) values$", RegexOptions.CultureInvariant)]
    private static partial Regex EnterValuesRegex();

    [GeneratedRegex("^alpha version (.+) - developed by (@[^ ]+)$", RegexOptions.CultureInvariant)]
    private static partial Regex SubtitleRegex();

    [GeneratedRegex("^Value: ReservedCpuSets = (.+) \\| CPUs: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex CpuValueRegex();

    [GeneratedRegex(@"^Apply changes \(Ctrl\+S\) — modified: (\d+)$", RegexOptions.CultureInvariant)]
    private static partial Regex ApplyChangesDirtyRegex();

    [GeneratedRegex(@"^CPU (\d+): (P-core SMT sibling|P-core|E-core)(, CCD \d+)?(, CCX \d+)?, Group (\d+), Core (\d+), Local (\d+)\. CPPC: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex CpuTooltipRegex();

    [GeneratedRegex(@"^rating (\d+), (preferred rank #1|rank #\d+)$", RegexOptions.CultureInvariant)]
    private static partial Regex CppcRatingRankRegex();

    [GeneratedRegex(@"^(preferred rank #1|rank #(\d+))$", RegexOptions.CultureInvariant)]
    private static partial Regex CppcRankOnlyRegex();

    private static string TranslateCpuTooltip(Match match)
    {
        string lp = match.Groups[1].Value;
        string coreType = match.Groups[2].Value switch
        {
            "P-core SMT sibling" => "Поток SMT (P-ядро)",
            "E-core" => "E-ядро",
            _ => "P-ядро",
        };
        string ccd = match.Groups[3].Value;
        string ccx = match.Groups[4].Value;
        string group = match.Groups[5].Value;
        string core = match.Groups[6].Value;
        string local = match.Groups[7].Value;
        string cppcRaw = match.Groups[8].Value;

        string cppcTranslated;
        if (cppcRaw == "unavailable")
        {
            cppcTranslated = "недоступно";
        }
        else
        {
            Match cppcFull = CppcRatingRankRegex().Match(cppcRaw);
            if (cppcFull.Success)
            {
                string rankText = cppcFull.Groups[2].Value == "preferred rank #1"
                    ? "приоритетный ранг #1"
                    : $"ранг #{cppcFull.Groups[2].Value.Replace("rank #", "", StringComparison.Ordinal)}";
                cppcTranslated = $"рейтинг {cppcFull.Groups[1].Value}, {rankText}";
            }
            else
            {
                Match cppcRankOnly = CppcRankOnlyRegex().Match(cppcRaw);
                if (cppcRankOnly.Success)
                {
                    cppcTranslated = cppcRankOnly.Groups[1].Value == "preferred rank #1"
                        ? "приоритетный ранг #1"
                        : $"ранг #{cppcRankOnly.Groups[2].Value}";
                }
                else
                {
                    cppcTranslated = cppcRaw;
                }
            }
        }

        return $"CPU {lp}: {coreType}{ccd}{ccx}, Группа {group}, Ядро {core}, Локальный индекс {local}. CPPC: {cppcTranslated}";
    }
}
