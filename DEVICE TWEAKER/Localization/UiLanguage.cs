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
                    && !entry.Value.Contains(requiredToken, StringComparison.OrdinalIgnoreCase)
                    && !(requiredToken == "interrupter" && entry.Value.Contains("интерраптер", StringComparison.OrdinalIgnoreCase))
                    && !(requiredToken == "IMOD persistence" && entry.Value.Contains("IMOD", StringComparison.Ordinal)
                        && entry.Value.Contains("автозапуск", StringComparison.OrdinalIgnoreCase)))
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
        => IsRussian ? TranslateRussianRole(value) : value;

    private static string TranslateRussianRole(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
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

            if (value.StartsWith(english + " ", StringComparison.OrdinalIgnoreCase)
                || value.StartsWith(english + ":", StringComparison.OrdinalIgnoreCase))
            {
                return russian + value[english.Length..];
            }
        }

        return value;
    }

    private static string TranslateDriverPreflightState(string source) => source
        .Replace("not explicitly configured", "нет явного параметра", StringComparison.Ordinal)
        .Replace("not found", "не найдена", StringComparison.Ordinal)
        .Replace("found (", "найдена (", StringComparison.Ordinal)
        .Replace("configured off", "задано выкл", StringComparison.Ordinal)
        .Replace("configured on", "задано вкл", StringComparison.Ordinal)
        .Replace("unavailable", "недоступно", StringComparison.Ordinal)
        .Replace("disabled", "выключена", StringComparison.Ordinal)
        .Replace("enabled", "включена", StringComparison.Ordinal)
        .Replace("running", "работает", StringComparison.Ordinal)
        .Replace("stopped", "остановлена", StringComparison.Ordinal)
        .Replace("starting", "запускается", StringComparison.Ordinal)
        .Replace("stopping", "останавливается", StringComparison.Ordinal)
        .Replace("continuing", "возобновляется", StringComparison.Ordinal)
        .Replace("pausing", "приостанавливается", StringComparison.Ordinal)
        .Replace("paused", "приостановлена", StringComparison.Ordinal);

    private static string TranslateLine(string source)
    {
        if (Russian.TryGetValue(source, out string? exact))
        {
            return exact;
        }

        if (source.StartsWith("Vulnerable Driver Blocklist (registry): ", StringComparison.Ordinal))
            return "Vulnerable Driver Blocklist (реестр): " + TranslateDriverPreflightState(source["Vulnerable Driver Blocklist (registry): ".Length..]);
        if (source.StartsWith("Memory Integrity / HVCI: runtime ", StringComparison.Ordinal))
            return "Целостность памяти / HVCI: сейчас " + TranslateDriverPreflightState(source["Memory Integrity / HVCI: runtime ".Length..].Replace(", registry ", ", в реестре ", StringComparison.Ordinal));
        if (source.StartsWith("Microsoft Defender service: ", StringComparison.Ordinal))
            return "Служба Microsoft Defender: " + TranslateDriverPreflightState(source["Microsoft Defender service: ".Length..]);
        if (source.StartsWith("FACEIT service: ", StringComparison.Ordinal))
            return "Служба FACEIT: " + TranslateDriverPreflightState(source["FACEIT service: ".Length..]);
        if (source.StartsWith("Vanguard: vgk ", StringComparison.Ordinal))
            return "Vanguard: vgk " + TranslateDriverPreflightState(source["Vanguard: vgk ".Length..]);

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

        match = TargetsRegex().Match(source);
        if (match.Success)
        {
            return $"Устройств: {match.Groups[1].Value}";
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
            return $"версия {match.Groups[1].Value} - разработчик {match.Groups[2].Value}";
        }

        match = CpuValueRegex().Match(source);
        if (match.Success)
        {
            return $"Значение: ReservedCpuSets = {match.Groups[1].Value} | CPU: {match.Groups[2].Value}";
        }

        match = ApplyChangesDirtyRegex().Match(source);
        if (match.Success)
        {
            return $"Применить настройки (Ctrl+S) - изменено: {match.Groups[1].Value}";
        }

        match = PendingApplyRegex().Match(source);
        if (match.Success)
            return $"ожидает ПРИМЕНИТЬ: {TranslateInlineStatus(match.Groups[1].Value)}";

        match = ReservedValueRegex().Match(source);
        if (match.Success)
        {
            string prefix = match.Groups[1].Value == "Pending APPLY" ? "Ожидает ПРИМЕНИТЬ" : "Значение";
            string value = match.Groups[2].Value == "not set" ? "не задано" : match.Groups[2].Value.Replace(" | CPUs:", " | CPU:", StringComparison.Ordinal);
            return $"{prefix}: ReservedCpuSets = {value}";
        }

        if (source.StartsWith("HID: ", StringComparison.Ordinal)
            || source.StartsWith("Polling: ", StringComparison.Ordinal)
            || source.StartsWith("Polling live: ", StringComparison.Ordinal))
        {
            int colon = source.IndexOf(':');
            string prefix = source[..colon] switch { "HID" => "HID", "Polling" => "Опрос", _ => "Опрос live" };
            return prefix + ": " + string.Join(", ", source[(colon + 2)..].Split(',', StringSplitOptions.TrimEntries).Select(TranslateRussianRole));
        }

        if (source.StartsWith("Power Saving: ", StringComparison.Ordinal))
            return "Power Saving: " + TranslateInlineStatus(source[14..]);

        if (source.StartsWith("Topology: estimated ", StringComparison.Ordinal))
            return "Топология: ориентировочно " + source[20..];

        match = RssRuntimeSummaryRegex().Match(source);
        if (match.Success)
            return $"RSS: {TranslateInlineStatus(match.Groups[1].Value)} | RSS Base: {match.Groups[2].Value} | RSS Queues: {match.Groups[3].Value}";

        match = IrqSummaryRegex().Match(source);
        if (match.Success)
            return $"IRQ: {TranslateInlineStatus(match.Groups[1].Value)} | WMI hint: {TranslateInlineStatus(match.Groups[2].Value)} | HW: {TranslateInlineStatus(match.Groups[3].Value)}";

        match = IrqHardwareSummaryRegex().Match(source);
        if (match.Success)
            return $"IRQ: {TranslateInlineStatus(match.Groups[1].Value)} | HW: {TranslateInlineStatus(match.Groups[2].Value)}";

        match = ImodAppliedRegex().Match(source);
        if (match.Success)
            return $"IMOD применён на USB-контроллерах: {match.Groups[1].Value}.";

        match = ImodWarningRegex().Match(source);
        if (match.Success)
            return $"USB IMOD применён с предупреждениями (записано {match.Groups[1].Value}/{match.Groups[2].Value}).";

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
            return string.Join(" | ", pipeParts.Select(TranslateRussianRole));
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
        if (QaRun.TryBlock("LANGUAGE PREFERENCE")) return false;
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
        QaRun.DemandWrite("LANGUAGE LEGACY CLEANUP");
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
        QaRun.DemandWrite("LANGUAGE SETTINGS CLEANUP");
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
            ["Loaded DTIMOD does not support PCI queries. Restart Windows, then read IRQ Mode again."] = "Загруженный DTIMOD не поддерживает чтение PCI. Перезагрузите Windows и повторите чтение IRQ Mode.",
            ["DEVICE TWEAKER"] = "DEVICE TWEAKER",
            ["APPLY"] = "ПРИМЕНИТЬ",
            ["AUTO-OPTIMIZATION"] = "АВТООПТИМИЗАЦИЯ",
            ["REFRESH"] = "ОБНОВИТЬ",
            ["RESTORE"] = "ВОССТАНОВЛЕНИЕ",
            ["RESET WINDOWS DEFAULT"] = "СБРОС WINDOWS",
            ["DELETE SELECTED"] = "УДАЛИТЬ",
            ["RESTORE LAST"] = "ПОСЛЕДНИЙ БЕКАП",
            ["RESTORE SELECTED"] = "ВЫБРАННЫЙ БЕКАП",
            ["CANCEL"] = "ОТМЕНА",
            ["YES"] = "ДА",
            ["NO"] = "НЕТ",
            ["OK"] = "ОК",
            ["DETAILS"] = "ПОДРОБНОСТИ",
            ["HIDE"] = "СКРЫТЬ",
            ["HIDE DETAILS"] = "СКРЫТЬ ПОДРОБНОСТИ",
            ["Show / hide IMOD values by device and interrupter."] = "Показать / скрыть значения IMOD по устройствам и интерраптерам.",
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
            ["Storage affinity is managed by Windows / the storage driver. Editing it here is disabled."] = "Affinity накопителей управляется Windows / драйвером накопителя. Редактирование здесь отключено.",
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
            ["Changed"] = "Изменено",
            ["Read-only"] = "Только чтение",
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
            ["No devices match this filter."] = "По этому фильтру устройства не найдены.",
            ["COPY REGISTRY"] = "СКОПИРОВАТЬ РЕЕСТР",
            ["Click to copy the registry path."] = "Нажмите, чтобы скопировать путь реестра.",
            ["Registry path copied to clipboard."] = "Путь в реестре скопирован в буфер обмена.",
            ["Select all logical cores"] = "Выбрать все логические ядра",
            ["Clear all cores"] = "Снять выбор со всех ядер",
            ["Select physical P-cores only"] = "Выбрать только физические P-ядра",
            ["Select CCD0 cores only"] = "Выбрать только ядра CCD0",
            ["MSI Mode:"] = "MSI Mode:",
            ["IRQ Mode:"] = "IRQ Mode:",
            ["IRQ count: WMI data. HW: supported modes. CHECK beside IRQ Mode reads the current mode through DTIMOD."] = "Количество IRQ берется из WMI. HW показывает поддерживаемые режимы. Кнопка ПРОВЕРИТЬ рядом с IRQ Mode считывает текущий режим через DTIMOD.",
            ["not read"] = "не считан",
            ["Simulated Windows allocation. "] = "Тестовое распределение ресурсов Windows. ",
            ["Reading Windows allocated interrupt resources. No driver is loaded."] = "Чтение выделенных ресурсов Windows. Драйвер не загружается.",
            ["IRQ count comes from WMI. IRQ Mode uses Windows allocated resources. HW lists supported modes. Refresh updates the readings without loading a driver."] = "Количество IRQ берётся из WMI. IRQ Mode определяется по выделенным ресурсам Windows. HW показывает поддерживаемые режимы. ОБНОВИТЬ повторяет чтение без загрузки драйвера.",
            ["Source: Windows allocated resources. Separate message descriptors allocated. PCI Enable bits were not read."] = "Источник: выделенные ресурсы Windows. Выделены отдельные дескрипторы сообщений. Биты PCI Enable не считывались.",
            ["Source: Windows allocated resources. One descriptor with multiple messages allocated. PCI Enable bits were not read."] = "Источник: выделенные ресурсы Windows. Выделен один дескриптор с несколькими сообщениями. Биты PCI Enable не считывались.",
            ["Source: Windows allocated resources. Message interrupt allocated; hardware supports MSI, not MSI-X. PCI Enable bits were not read."] = "Источник: выделенные ресурсы Windows. Выделено прерывание сообщениями. Устройство поддерживает MSI, поддержка MSI-X отсутствует. Биты PCI Enable не считывались.",
            ["Source: Windows allocated resources. Message interrupt allocated; hardware supports MSI-X, not MSI. PCI Enable bits were not read."] = "Источник: выделенные ресурсы Windows. Выделено прерывание сообщениями. Устройство поддерживает MSI-X, поддержка MSI отсутствует. Биты PCI Enable не считывались.",
            ["Source: Windows allocated resources. One message allocated; hardware supports MSI and MSI-X. These resources cannot distinguish MSI from MSI-X. PCI Enable bits were not read."] = "Источник: выделенные ресурсы Windows. Выделено одно сообщение. Устройство поддерживает оба режима; этих данных недостаточно, чтобы различить MSI и MSI-X. Биты PCI Enable не считывались.",
            ["Source: Windows allocated resources. Message interrupts allocated; hardware capabilities unavailable. PCI Enable bits were not read."] = "Источник: выделенные ресурсы Windows. Выделены прерывания сообщениями. Возможности устройства прочитать не удалось. Биты PCI Enable не считывались.",
            ["Source: Windows allocated resources. Line-based interrupt resources allocated. PCI Enable bits were not read."] = "Источник: выделенные ресурсы Windows. Выделены линейные прерывания. Биты PCI Enable не считывались.",
            ["Source: Windows allocated resources. No interrupt resources allocated. PCI Enable bits were not read."] = "Источник: выделенные ресурсы Windows. Ресурсы прерываний не выделены. Биты PCI Enable не считывались.",
            ["An older DTIMOD driver is loaded. Restart Windows, then press CHECK again."] = "Загружена предыдущая версия DTIMOD. Перезагрузите Windows, затем нажмите ПРОВЕРИТЬ.",
            ["DTIMOD is unavailable.\nRestart Windows, then press CHECK."] = "DTIMOD недоступен.\nПерезагрузите Windows и нажмите ПРОВЕРИТЬ.",
            ["DTIMOD loading did not finish normally.\nRestart Windows, then press CHECK."] = "Загрузка DTIMOD не завершилась штатно.\nПерезагрузите Windows и нажмите ПРОВЕРИТЬ.",
            ["not PCI"] = "не PCI",
            ["no IRQ"] = "нет IRQ",
            ["No IRQ"] = "Нет IRQ",
            ["Not read"] = "Не считан",
            ["Read failed"] = "Ошибка чтения",
            ["Read the current interrupt mode. Loads DTIMOD if needed. Device settings are not changed."] = "Считать текущий режим прерываний. При необходимости загрузит DTIMOD. Настройки устройства не меняются.",
            ["IRQ Mode: Windows allocated resources / PCI MSI Enable and MSI-X Enable bits. HW lists supported modes, not the active mode."] = "IRQ Mode определяется по выделенным ресурсам Windows / битам MSI Enable и MSI-X Enable в PCI. HW показывает поддерживаемые режимы.",
            ["This device has no PCI interrupt mode."] = "У этого устройства нет режима прерываний PCI.",
            ["Windows has allocated no interrupt resources to this device."] = "Windows не выделил этому устройству ресурсы прерываний.",
            ["Simulated interrupt mode. No hardware was queried."] = "Тестовый режим прерываний. Железо не опрашивалось.",
            ["Loaded DTIMOD does not support PCI queries. An older driver may still be loaded."] = "Загруженный DTIMOD не поддерживает чтение PCI. Возможно, в памяти осталась старая версия драйвера.",
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
            ["RESULT: WARNINGS"] = "РЕЗУЛЬТАТ: ПРЕДУПРЕЖДЕНИЯ",
            ["RESULT: PARTIAL (3x)"] = "РЕЗУЛЬТАТ: ЧАСТИЧНО (3x)",
            ["Partially applied (5/8 successful, 3 skipped or failed)."] = "Применено частично (успешно: 5/8; пропущено или с ошибкой: 3).",
            ["Stress preview: verify clipping, wrapping, scrolling and DETAILS."] = "Стресс-тест: проверьте обрезку текста, переносы, прокрутку и ПОДРОБНОСТИ.",
            ["A deliberately long summary verifies wrapping without exposing raw diagnostics in the main result."] = "Длинное сообщение для проверки переносов: технические подробности выводятся в отдельном блоке.",
            ["Kernel CI blocked driver execution (code 577)."] = "Kernel CI заблокировал запуск драйвера (код 577).",
            ["RESULT: PARTIAL"] = "РЕЗУЛЬТАТ: ЧАСТИЧНО",
            ["RESULT: FAILED"] = "РЕЗУЛЬТАТ: ОШИБКА",
            ["RESULT: STRESS"] = "РЕЗУЛЬТАТ: СТРЕСС-ТЕСТ",
            ["PROMPT: IMOD"] = "ДИАЛОГ: IMOD",
            ["PROMPT: CONFIRM"] = "ДИАЛОГ: ПОДТВЕРЖДЕНИЕ",
            ["RESET"] = "СБРОСИТЬ",
            ["PROMPT: RESTORE (0)"] = "ДИАЛОГ: ВОССТАНОВЛЕНИЕ (0)",
            ["RESET ADAPTER SETTINGS"] = "СБРОС НАСТРОЕК АДАПТЕРА",
            ["Reset all active network tweaks and restore default Windows adapter parameters?"] = "Сбросить сетевые настройки к стандартным значениям Windows?",
            ["This will re-enable interrupt moderation and standard RSS ring buffers."] = "Будут включены Interrupt Moderation и стандартные RSS ring buffers.",
            ["PROMPT: BACKUP"] = "ДИАЛОГ: БЕКАП",
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
            ["Backup"] = "Бекап",
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
            ["Optional. Format: 0=120, 1=110, 2=100 or just 120,110,100. Empty = CPPC off in test mode."] = "Необязательно. Формат: 0=120, 1=110, 2=100 или просто 120,110,100. Если поле пустое, CPPC выключен в тестовом режиме.",
            ["Optional fake hardware ID. Example: PCI\\VEN_10EC&DEV_8125\\TEST for NIC ITR profile preview."] = "Необязательный тестовый ID оборудования. Пример: PCI\\VEN_10EC&DEV_8125\\TEST для предпросмотра профиля NIC ITR.",
            ["Optional fake IRQ count shown in device blocks. Empty/Auto = role-based default."] = "Необязательное тестовое количество IRQ в карточках устройств. Если поле пустое или выбрано Авто, используется значение по роли.",
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
            ["Automatic backup"] = "Автоматический бекап",
            ["AUTOMATIC BACKUP"] = "АВТОМАТИЧЕСКИЙ БЕКАП",
            ["BACKUP"] = "БЕКАП",
            ["BACKUP RESTORE"] = "ВОССТАНОВЛЕНИЕ БЕКАПА",
            ["DEVICE SETTINGS"] = "НАСТРОЙКИ УСТРОЙСТВ",
            ["AUTO PLAN"] = "ПЛАН АВТООПТИМИЗАЦИИ",
            ["USB POWER"] = "ПИТАНИЕ USB",
            ["USB IMOD"] = "USB IMOD",
            ["REFRESH UI"] = "ОБНОВЛЕНИЕ ИНТЕРФЕЙСА",
            ["Reserved CPU sets"] = "Reserved CPU sets",
            ["Raw mouse throttle"] = "Raw mouse throttle",
            ["IMOD persistence"] = "Автозапуск IMOD",
            ["USB Selective Suspend power plan"] = "USB Selective Suspend в плане питания",
            ["Applied"] = "Применено",
            ["Applied with warnings. Core settings were successfully configured."] = "Основные настройки применены с замечаниями.",
            ["The operation stopped before any changes were made."] = "Настройки не изменены: операция остановлена до применения.",
            ["Built"] = "Сформирован",
            ["Selected settings were applied and saved."] = "Выбранные настройки применены и сохранены.",
            ["Please reboot your PC to finish applying the changes."] = "Перезагрузите компьютер, чтобы завершить применение изменений.",
            ["Applied steps were saved."] = "Примененные настройки сохранены.",
            ["Applied steps were saved. One optional step was not completed."] = "Настройки сохранены. Один дополнительный шаг не выполнен.",
            ["DTIMOD was unavailable. USB IMOD was not changed."] = "DTIMOD недоступен. Значение USB IMOD не изменено.",
            ["DTIMOD driver was unavailable. USB IMOD was not changed."] = "Драйвер DTIMOD недоступен. USB IMOD не изменён.",
            ["DTIMOD was rejected by Windows or antivirus. USB IMOD was not changed."] = "Загрузка DTIMOD отклонена Windows или антивирусом. USB IMOD не изменён.",
            ["DTIMOD was rejected by Windows or antivirus."] = "Загрузка DTIMOD отклонена Windows или антивирусом.",
            ["Windows rejected DTIMOD due to code integrity/signature checks. USB IMOD was not changed."] = "Windows отклонила DTIMOD при проверке целостности кода / подписи. USB IMOD не изменён.",
            ["Windows rejected DTIMOD due to code integrity/signature checks."] = "Windows отклонила DTIMOD при проверке целостности кода / подписи.",
            ["DTIMOD.sys loader rejected by Windows"] = "загрузчик DTIMOD.sys отклонён Windows",
            ["DTIMOD.sys rejected by Windows"] = "DTIMOD.sys отклонён Windows",
            ["Driver preflight (read-only):"] = "Проверка системы (только чтение):",
            ["Apply USB IMOD during AUTO-OPTIMIZATION?"] = "Применить USB IMOD при автооптимизации?",
            ["SYSTEM CHECK +"] = "ПРОВЕРКА СИСТЕМЫ +",
            ["SYSTEM CHECK"] = "ПРОВЕРКА СИСТЕМЫ",
            ["DEVICE TWEAKER has crashed and will close.\n\nA crash report is saved in logs. No further settings will be applied."] = "В DEVICE TWEAKER произошел сбой. Программа будет закрыта.\n\nОтчет о сбое сохранен в logs. Новые настройки применяться не будут.",
            ["Run DEVICE TWEAKER as administrator to apply settings.\n\nRight-click the EXE and choose 'Run as administrator'."] = "Для применения настроек запустите DEVICE TWEAKER от имени администратора.\n\nНажмите правой кнопкой по EXE → Запуск от имени администратора.",
            ["NIC ITR CHECK"] = "ПРОВЕРКА NIC ITR",
            ["USB IMOD DELETE"] = "УДАЛЕНИЕ USB IMOD",
            ["USB IMOD CHECK"] = "ПРОВЕРКА USB IMOD",
            ["current: driver check skipped (test)"] = "текущее: чтение драйвера пропущено (тест)",
            ["RESET WINDOWS DEFAULT restores supported settings to their defaults."] = "СБРОС WINDOWS возвращает поддерживаемые настройки к стандартным значениям.",
            ["Registry settings / IMOD startup profile from the backup."] = "Настройки реестра / профиль автозапуска IMOD из бекапа.",
            ["Select a backup to restore settings:"] = "Выберите бекап для восстановления настроек:",
            ["The original-settings backup (BASELINE) is protected from deletion."] = "Бекап исходных настроек защищён от удаления.",
            ["Larger values request longer intervals for background listeners."] = "При большем значении запрашивается более длинный интервал для фоновых приложений.",
            ["Accepted values: DWORD 1–20 ms."] = "Допустимые значения: DWORD 1–20 мс.",
            ["RawMouseThrottleDuration: HKCU\\Control Panel\\Mouse."] = "RawMouseThrottleDuration: HKCU\\Control Panel\\Mouse.",
            ["High mouse polling rates can increase this load."] = "Высокая частота опроса мыши может увеличивать эту нагрузку.",
            ["Can reduce the message load from background listeners."] = "Может уменьшить нагрузку от обработки фоновых сообщений.",
            ["Limits Raw Input delivery to eligible background listeners on supported Windows builds."] = "Ограничивает доставку Raw Input подходящим фоновым приложениям на поддерживаемых сборках Windows.",
            ["IMOD startup script saved. USB live apply skipped (no custom USB IMOD config)."] = "Скрипт автозапуска IMOD сохранен. Текущие значения не изменены: пользовательский профиль USB IMOD не задан.",
            ["USB IMOD is active for this session, but startup persistence was not saved."] = "USB IMOD применен в текущей сессии, но профиль автозапуска не сохранен.",
            ["DTIMOD driver unavailable. Open the session log for diagnostics."] = "Драйвер DTIMOD недоступен. Причину смотрите в логе.",
            ["The cause of rejection is unconfirmed. Other policies or anti-cheats may still block DTIMOD."] = "Причина отказа не подтверждена. Другие политики / античиты могут по-прежнему блокировать DTIMOD.",
            ["Disable Vulnerable Driver Blocklist?"] = "Отключить Vulnerable Driver Blocklist?",
            ["Windows rejected DTIMOD. Vulnerable Driver Blocklist is enabled in the registry."] = "Windows отклонила загрузку DTIMOD. Vulnerable Driver Blocklist включен в реестре.",
            ["For USB: controller / root hubs / power-plan Selective Suspend."] = "Для USB: Selective Suspend контроллера / корневых хабов / плана питания.",
            ["Original baseline snapshot is permanently protected."] = "Бекап исходных настроек защищен от удаления.",
            ["Removing IMOD persistence..."] = "Удаление автозапуска IMOD...",
            ["Building device list (0/0)"] = "Подготовка списка устройств (0/0)",
            ["no devices to optimize"] = "Нет устройств для авто-оптимизации.",
            ["Unknown DTIMOD driver load error."] = "Не удалось загрузить DTIMOD. Причина не определена.",
            ["SYSTEM CHECK: Windows settings / anti-cheat services."] = "Проверка системы: параметры Windows / службы античитов.",
            ["SKIP continues auto-optimization without USB IMOD."] = "При пропуске авто-оптимизация продолжится без USB IMOD.",
            ["It may conflict with DTIMOD.sys."] = "Возможен конфликт с DTIMOD.sys.",
            ["They may conflict with DTIMOD.sys."] = "Возможен конфликт с DTIMOD.sys.",
            ["Vanguard detected."] = "Найдены компоненты Vanguard.",
            ["FACEIT detected."] = "Найдены компоненты FACEIT.",
            ["FACEIT and Vanguard detected."] = "Найдены компоненты FACEIT / Vanguard.",
            ["Include USB IMOD in auto-optimization?"] = "Включить USB IMOD в авто-оптимизацию?",
            ["The remaining results are available in DETAILS."] = "Остальные результаты доступны в подробностях.",
            ["See DETAILS for applied / skipped settings and their reasons."] = "Результаты всех шагов смотрите в подробностях.",
            ["Full results / warnings are in DETAILS."] = "Результаты и замечания смотрите в подробностях.",
            ["See DETAILS and LOGS for the error."] = "Причину ошибки смотрите в подробностях и логах.",
            ["APPLY includes USB IMOD. SKIP continues auto-optimization without it."] = "ПРИМЕНИТЬ включает USB IMOD в авто-оптимизацию. ПРОПУСТИТЬ продолжает без него.",
            ["SYSTEM CHECK shows the detected Windows settings and anti-cheat services."] = "ПРОВЕРКА СИСТЕМЫ показывает найденные параметры Windows и службы античитов.",
            ["Windows settings may prevent DTIMOD.sys from loading."] = "Настройки Windows могут препятствовать загрузке DTIMOD.",
            ["Some system checks are unavailable."] = "Часть проверок системы недоступна.",
            ["Windows settings / service status. This check does not test driver loading or anti-cheat compatibility."] = "Параметры Windows / состояние служб. Загрузка драйвера и совместимость с античитами здесь не проверяются.",
            ["Not configured = no registry value. This does not mean the feature is disabled."] = "Не задано = параметра нет в реестре. Это не означает, что функция отключена.",
            ["Skipped"] = "пропущено",
            ["Not applied"] = "не выполнено",
            ["Warning"] = "замечание",
            ["Where should the additional backup be saved?"] = "Куда сохранить дополнительный бекап?",
            ["The EXE folder keeps the backup next to the program."] = "В папке EXE бекап будет рядом с программой.",
            ["AppData keeps it in your user profile."] = "В AppData он останется в профиле пользователя.",
            ["You can skip this copy. The original-settings backup is kept either way."] = "Можно пропустить. Бекап исходных настроек сохраняется в любом случае.",
            ["Blocklist (registry)"] = "Blocklist (реестр)",
            ["HVCI (runtime)"] = "HVCI (сейчас)",
            ["HVCI (registry)"] = "HVCI (реестр)",
            ["Not configured"] = "Не задано",
            ["Not found"] = "Не найдено",
            ["Running"] = "Работает",
            ["Stopped"] = "Остановлено",
            ["Starting"] = "Запускается",
            ["Stopping"] = "Останавливается",
            ["Continuing"] = "Возобновляется",
            ["Pausing"] = "Приостанавливается",
            ["Paused"] = "Приостановлено",
            ["State unavailable"] = "Состояние недоступно",
            ["Registry configuration is not a driver-load verdict. Driver loading and anti-cheat compatibility are untested."] = "Настройки реестра не подтверждают возможность загрузки драйвера. Загрузка и совместимость с античитами не проверялись.",
            ["Vanguard components found. A conflict with DTIMOD/KDU is possible."] = "Обнаружены компоненты Vanguard. Возможен конфликт с DTIMOD/KDU.",
            ["FACEIT components found. A conflict with DTIMOD/KDU is possible."] = "Обнаружены компоненты FACEIT. Возможен конфликт с DTIMOD/KDU.",
            ["FACEIT and Vanguard components found. A conflict with DTIMOD/KDU is possible."] = "Обнаружены компоненты FACEIT и Vanguard. Возможен конфликт с DTIMOD/KDU.",
            ["SYSTEM CHECK -"] = "ПРОВЕРКА СИСТЕМЫ −",
            ["Vanguard components found. Compatibility with DTIMOD/KDU is untested; a conflict is possible."] = "Обнаружены компоненты Vanguard. Возможен конфликт с DTIMOD/KDU; совместимость не проверена.",
            ["FACEIT components found. Compatibility with DTIMOD/KDU is untested; a conflict is possible."] = "Обнаружены компоненты FACEIT. Возможен конфликт с DTIMOD/KDU; совместимость не проверена.",
            ["FACEIT and Vanguard components found. Compatibility with DTIMOD/KDU is untested; a conflict is possible."] = "Обнаружены компоненты FACEIT и Vanguard. Возможен конфликт с DTIMOD/KDU; совместимость не проверена.",
            ["Windows settings may prevent DTIMOD loading. See the system check for details."] = "Настройки Windows могут препятствовать загрузке DTIMOD. Подробности доступны в проверке системы.",
            ["Some system checks are unavailable. See details."] = "Часть проверок системы недоступна. Подробности доступны в проверке системы.",
            ["USB IMOD tuning is available for detected XHCI controller(s)."] = "Для обнаруженных XHCI-контроллеров доступна настройка USB IMOD.",
            ["Apply it during AUTO-OPTIMIZATION?"] = "Применить её во время АВТООПТИМИЗАЦИИ?",
            ["Anti-cheat components were found. Compatibility with DTIMOD/KDU is not confirmed; a conflict is possible, not detected."] = "Найдены компоненты античита. Совместимость с DTIMOD/KDU не подтверждена; конфликт возможен, но не обнаружен.",
            ["These checks do not test driver loading. A missing registry value does not mean the protection is disabled."] = "Эта проверка не загружает драйвер. Отсутствие параметра в реестре не означает, что защита отключена.",
            ["Windows rejected DTIMOD. Vulnerable Driver Blocklist is configured as enabled."] = "Windows отклонила DTIMOD. В реестре Vulnerable Driver Blocklist задан как включённый.",
            ["Change the explicit Vulnerable Driver Blocklist setting to off? This does not guarantee that DTIMOD will load."] = "Изменить параметр Vulnerable Driver Blocklist на выключенный? Это не гарантирует загрузку DTIMOD.",
            ["After reboot, run CHECK or AUTO-OPTIMIZATION again to apply IMOD."] = "После перезагрузки снова нажмите ПРОВЕРИТЬ или АВТООПТИМИЗАЦИЮ, чтобы применить IMOD.",
            ["This does not identify the cause of the rejection. Other policies may still prevent loading; anti-cheat compatibility is not confirmed."] = "Это не устанавливает причину отказа. Другие политики также могут запрещать загрузку; совместимость с античитами не подтверждена.",
            ["AUTO profile: Mouse IMOD = 0 (interrupt moderation off). Manual IMOD is set separately."] = "Профиль AUTO: IMOD для мыши = 0 (модерация прерываний отключена). Ручное значение IMOD задаётся отдельно.",
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
            ["DTIMOD.sys rejected by Windows or antivirus"] = "Загрузка DTIMOD.sys отклонена Windows или антивирусом",
            ["OPEN BACKUP"] = "ОТКРЫТЬ БЕКАП",
            ["BACKUP"] = "БЕКАП",
            ["BACKUPS"] = "БЕКАПЫ",
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
            ["The selected backup is invalid."] = "Бекап поврежден или имеет неподходящий формат.",
            ["The backup file could not be created."] = "Не удалось создать файл бекапа.",
            ["backup could not be created; no settings were changed"] = "не удалось создать бекап; настройки не изменены",
            ["backup could not be created; changes were not applied"] = "не удалось создать бекап; изменения не применены",
            ["Settings were restored. Please reboot your PC."] = "Настройки восстановлены. Перезагрузите компьютер.",
            ["Some restore steps may already have run."] = "Часть настроек уже могла восстановиться.",
            ["Restore did not finish. Open DETAILS before rebooting."] = "Перед перезагрузкой откройте подробности восстановления.",
            ["RESET WINDOWS DEFAULT was cancelled because the rollback backup failed."] = "Сброс Windows отменён: не удалось создать бекап для отката.",
            ["Supported tweaks were reset."] = "Поддерживаемые настройки сброшены.",
            ["Please reboot your PC to restore runtime hardware defaults."] = "Перезагрузите компьютер, чтобы восстановить стандартные параметры оборудования.",
            ["RESET WINDOWS DEFAULT finished with errors. Some settings may still be active."] = "Сброс Windows завершён с ошибками. Некоторые настройки могут оставаться активными.",
            ["Choose how to restore DEVICE TWEAKER settings."] = "Выберите способ восстановления настроек DEVICE TWEAKER.",
            ["RESTORE LAST uses the newest snapshot. RESTORE SELECTED uses the highlighted snapshot."] = "ПОСЛЕДНИЙ восстанавливает свежий снимок. ВЫБРАННЫЙ восстанавливает выделенный снимок.",
            ["RESET WINDOWS DEFAULT restores supported settings to Windows defaults."] = "СБРОС WINDOWS возвращает поддерживаемые настройки к значениям Windows по умолчанию.",
            ["ORIGINAL STATE is protected and never pruned."] = "ИСХОДНОЕ СОСТОЯНИЕ защищено и не удаляется автоматически.",
            ["No backup snapshot was found."] = "Снимки бекапов не найдены.",
            ["ORIGINAL STATE"] = "ИСХОДНОЕ СОСТОЯНИЕ",
            ["ORIGINAL"] = "ИСХОДНЫЙ",
            ["Windows settings"] = "Параметры Windows",
            ["Services"] = "Службы",
            ["Vulnerable driver blocklist"] = "Блокировка уязвимых драйверов",
            ["Memory integrity (HVCI)"] = "Целостность памяти (HVCI)",
            ["HVCI registry setting"] = "Параметр HVCI в реестре",
            ["No registry setting"] = "Нет параметра",
            ["Could not read"] = "Не удалось считать",
            ["Service not installed"] = "Не установлена",
            ["Service running"] = "Работает",
            ["Service stopped"] = "Остановлена",
            ["This check reads settings and service status. It does not test DTIMOD.sys loading."] = "Считаны параметры Windows и состояния служб. Загрузка DTIMOD.sys не проверялась.",
            ["Refresh interrupt mode"] = "Обновить режим прерываний",
            ["Read interrupt modes for all PCI devices"] = "Считать режимы прерываний всех PCI-устройств",
            ["READ IRQ MODE"] = "СЧИТАТЬ IRQ",
            ["Read interrupt modes for all PCI devices, including hidden cards. Loads DTIMOD.sys if needed. Device settings are not changed."] = "Считать режимы прерываний всех PCI-устройств, включая скрытые карточки. При необходимости загрузит DTIMOD.sys. Настройки устройств не меняются.",
            ["refresh"] = "обновить",
            ["Refresh this device's interrupt mode. Loads DTIMOD.sys if needed. Device settings are not changed."] = "Обновить режим прерываний этого устройства. При необходимости загрузит DTIMOD.sys. Настройки устройства не меняются.",
            ["Click READ IRQ MODE at the top to read interrupt modes for all PCI devices. Loads DTIMOD.sys if needed. Device settings are not changed."] = "Нажмите СЧИТАТЬ IRQ вверху окна, чтобы считать режимы прерываний всех PCI-устройств. При необходимости загрузит DTIMOD.sys. Настройки устройства не меняются.",
            ["IRQ count comes from WMI. HW lists supported modes. Click READ IRQ MODE at the top to read all PCI interrupt modes through DTIMOD.sys."] = "Количество IRQ берётся из WMI. HW показывает поддерживаемые режимы. Кнопка СЧИТАТЬ IRQ вверху окна считывает режимы всех PCI-устройств через DTIMOD.sys.",
            ["pre-restore"] = "до восстановления",
            ["pre-release"] = "перед обновлением",
            ["pre-imod"] = "до настройки IMOD",
            ["pre-nic-itr"] = "до настройки NIC ITR",
            ["pre-nic-itr-save"] = "до сохранения NIC ITR",
            ["pre-apply"] = "до применения",
            ["pre-auto"] = "до автооптимизации",
            ["pre-reset"] = "до сброса",
            ["pre-reset-tweaks"] = "до сброса твиков",
            ["NVIDIA/AMD video driver not detected."] = "Драйвер NVIDIA / AMD не найден.",
            ["Install the GPU driver and press REFRESH."] = "Установите драйвер видеокарты и нажмите ОБНОВИТЬ.",
            ["Run DEVICE TWEAKER as administrator to apply settings."] = "Запустите DEVICE TWEAKER от имени администратора для применения настроек.",
            ["Right-click the EXE and choose 'Run as administrator'."] = "Нажмите правой кнопкой по EXE и выберите \"Запуск от имени администратора\".",
            ["DEVICE TWEAKER has crashed and will close."] = "В DEVICE TWEAKER произошел сбой. Программа будет закрыта.",
            ["A crash report is saved in logs. No further settings will be applied."] = "Отчет о сбое сохранен в logs. Новые настройки применяться не будут.",
            ["DEVICE TWEAKER: CRITICAL ERROR"] = "DEVICE TWEAKER: КРИТИЧЕСКАЯ ОШИБКА",
            ["Reads HKLM:\\System\\CurrentControlSet\\Control\\Session Manager\\Kernel\\ReservedCpuSets"] = "Читает HKLM:\\System\\CurrentControlSet\\Control\\Session Manager\\Kernel\\ReservedCpuSets",
            ["Registry: HKLM\\System\\CurrentControlSet\\Control\\Session Manager\\kernel"] = "Реестр: HKLM\\System\\CurrentControlSet\\Control\\Session Manager\\kernel",
            ["current: varies by interrupter"] = "текущее: зависит от interrupter",
            ["current: unavailable"] = "текущее: недоступно",
            ["current: reading..."] = "текущее: чтение...",
            ["current: loading driver..."] = "текущее: загрузка драйвера...",
            ["current: values shown below"] = "текущее: значения показаны ниже",
            ["current: test profile"] = "текущее: тестовый профиль",
            ["current: read successfully"] = "текущее: успешно прочитано",
            ["devices: driver unavailable"] = "устройства: драйвер недоступен",
            ["devices: reading..."] = "устройства: чтение...",
            ["devices: unavailable"] = "устройства: недоступно",
            ["devices: invalid IMOD input"] = "устройства: некорректное значение IMOD",
            ["detected USB device roles"] = "найденные роли USB-устройств",
            ["current: unavailable (Windows 11 22H2+)"] = "текущее: недоступно (Windows 11 22H2+)",
            ["0 = unlocked"] = "0 = без ограничения",
            ["(0 = unlocked)"] = "(0 = без ограничения)",
            ["(driver default)"] = "(настройка драйвера)",
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
            ["backup failed"] = "ошибка бекапа",
            ["applying..."] = "применение...",
            ["saved for startup"] = "сохранено для автозапуска",
            ["save failed"] = "ошибка сохранения",
            ["off"] = "выключено",
            ["on"] = "включено",
            ["driver blocked"] = "драйвер заблокирован",
            ["kernel CI blocked"] = "заблокировано kernel CI",
            ["signature blocked"] = "подпись драйвера заблокирована",
            ["driver not loaded — press CHECK"] = "драйвер не загружен, нажмите ПРОВЕРИТЬ",
            ["admin required"] = "требуются права администратора",
            ["No additional details."] = "Дополнительных данных нет.",
            ["single value for controller"] = "одно значение для контроллера",
            ["values by interrupter index"] = "значения по индексам interrupter",
            ["input: hex / list / roles"] = "ввод: hex / список / роли",
            ["time: reading..."] = "время: чтение...",
            ["time: unavailable"] = "время: недоступно",
            ["EXE FOLDER"] = "ПАПКА EXE",
            ["APPDATA"] = "APPDATA",
            ["AUTO BACKUP"] = "БЕКАП АВТООПТИМИЗАЦИИ",
            ["backup"] = "бекап",
            ["Where should DEVICE TWEAKER save the pre-auto backup?"] = "Где сохранить бекап перед автооптимизацией?",
            ["EXE FOLDER = portable backup next to the app."] = "ПАПКА EXE сохраняет переносимый бекап рядом с программой.",
            ["APPDATA = user profile backup that survives app folder cleanup."] = "APPDATA сохраняет бекап в профиле пользователя, отдельно от папки программы.",
            ["SKIP = do not create an additional pre-auto backup."] = "При пропуске дополнительный бекап перед авто-оптимизацией не создается.",
            ["The protected ORIGINAL STATE snapshot is always retained."] = "Защищённый снимок ИСХОДНОЕ СОСТОЯНИЕ сохраняется всегда.",
            ["Close (X) = cancel AUTO-OPTIMIZATION."] = "Закрытие окна (X) отменяет АВТООПТИМИЗАЦИЮ.",
            ["TEST DEVICE (no registry writes)"] = "ТЕСТОВОЕ УСТРОЙСТВО (без записи в реестр)",
            ["Protection: Wi-Fi settings preserved (no changes)"] = "Защита: настройки Wi-Fi не изменяются",
            ["Type: Storage controller"] = "Тип: контроллер накопителя",
            ["Type: NVMe storage controller"] = "Тип: контроллер NVMe-накопителя",
            ["Type: SATA/AHCI controller"] = "Тип: контроллер SATA/AHCI",
            ["Refresh devices (F5 / Ctrl+R)"] = "Обновить устройства (F5 / Ctrl+R)",
            ["Apply changes (Ctrl+S)"] = "Применить настройки (Ctrl+S)",
            ["Auto-optimization (Ctrl+O)"] = "Автооптимизация (Ctrl+O)",
            ["Restore settings (Ctrl+Z)"] = "Восстановить настройки (Ctrl+Z)",
            ["MSISupported requests MSI for the device. The active MSI / MSI-X mode depends on the hardware / driver."] = "MSISupported запрашивает MSI для устройства. Активный режим MSI / MSI-X зависит от устройства / драйвера.",
            ["MessageNumberLimit sets the requested message limit. Unlimited removes this limit (0 is also accepted). Hardware / Windows / the driver determine the actual message count."] = "MessageNumberLimit задает лимит сообщений. Unlimited снимает лимит (0 тоже принимается). Фактическое количество зависит от устройства / Windows / драйвера.",
            ["DevicePriority requests interrupt priority. High applies to interrupts, not game or DPC priority; actual scheduling depends on Windows / the driver."] = "DevicePriority задает запрашиваемый приоритет прерываний. High относится к прерываниям, а не приоритету игры / DPC. Фактическое выполнение зависит от Windows / драйвера.",
            ["DevicePolicy requests how Windows assigns processors to device interrupts:\n• SpecCPU uses the selected Affinity Mask\n• MachineDefault leaves placement to Windows\n• AllClose / Single request nearby or one processor\n• SpreadMessages may distribute multiple MSI-X messages when the device and driver use them.\nA driver can target DPCs separately."] =
                "DevicePolicy задаёт политику выбора процессоров для прерываний устройства:\n• SpecCPU использует выбранную маску Affinity Mask\n• MachineDefault оставляет выбор Windows\n• AllClose / Single запрашивают близкие процессоры либо один процессор\n• SpreadMessages может распределять сообщения MSI-X, если устройство и драйвер используют несколько векторов.\nДрайвер может отдельно выбирать CPU для DPC.",
            ["AssignmentSetOverride selects CPUs for device interrupts. The driver can send DPCs to other CPUs; check actual ISR / DPC placement with a trace."] = "AssignmentSetOverride задает CPU для прерываний устройства. Драйвер может отправить DPC на другой CPU. Фактическое распределение ISR / DPC проверяется трассировкой.",
            ["IRQ count: WMI data. WMI hint: MSI / MSI-X estimate from IRQ numbers above 999, not the active interrupt mode. It cannot tell MSI from MSI-X. HW: device capabilities."] = "Количество IRQ берется из WMI. WMI hint оценивает MSI / MSI-X по IRQ выше 999, но не определяет активный режим и не различает MSI и MSI-X. HW показывает возможности устройства.",
            ["Net type: NetAdapterCx"] = "Тип сети: NetAdapterCx",
            ["Copied"] = "Скопировано",
            ["RSS: receive queues / CPUs. IRQ: interrupt affinity. BOTH: both settings, if RSS is supported by the adapter / driver."] = "RSS настраивает receive queues / CPU. IRQ настраивает affinity прерываний. BOTH применяет обе настройки, если адаптер / драйвер поддерживает RSS.",
            ["RSS receive queue count. The adapter / driver sets the supported range. Assignment starts at RSS Base CPU."] = "Количество RSS receive queues. Диапазон зависит от адаптера / драйвера. Распределение начинается с RSS Base CPU.",
            ["Load DTIMOD.sys and read current NIC ITR values."] = "Загрузить DTIMOD.sys и считать текущие значения NIC ITR.",
            ["Supported IMOD input:"] = "Поддерживаемый формат IMOD:",
            ["XHCI: one value for all controller interrupters. Device mapping: values by device role; uses one controller-wide value if mapping is uncertain. Interrupter: values by index."] = "XHCI задает одно значение всем интерраптерам контроллера. Device mapping задает значения по ролям устройств. При неточном сопоставлении используется одно значение всему контроллеру. Interrupter задает значения по индексам.",
            ["Input format for the selected IMOD mode."] = "Формат ввода для выбранного IMOD mode.",
            ["Apply IMOD and re-read the values to check the write."] = "Применить IMOD и повторно считать значения для проверки записи.",
            ["Load DTIMOD.sys and read current IMOD values."] = "Загрузить DTIMOD.sys и считать текущие значения IMOD.",
            ["Role mode applies values to the detected interrupter for each USB role."] = "Role mode применяет значения к обнаруженному interrupter каждой USB-роли.",
            ["Click to use the default IMOD interval."] = "Нажмите, чтобы подставить стандартный интервал IMOD.",
            ["Wi-Fi settings are preserved by DEVICE TWEAKER."] = "DEVICE TWEAKER не изменяет настройки Wi-Fi.",
            ["Same idea as Device Manager → Power Management → 'Allow the computer to turn off this device to save power'."] = "Аналог параметра \"Разрешить отключение этого устройства для экономии энергии\" в диспетчере устройств.",
            ["For USB this disables Selective Suspend on the controller + root hubs and the power-plan USB SS setting."] = "Для USB отключает Selective Suspend у контроллера и корневых USB-концентраторов, а также USB SS в текущем плане электропитания.",
            ["Unchecked = Disabled. Applied with APPLY / AUTO. A reboot may be required."] = "Если флажок снят, функция выключена. Применяется через ПРИМЕНИТЬ или АВТООПТИМИЗАЦИЮ. Может потребоваться перезагрузка.",
            ["Device Manager → Power Management → 'Allow the computer to turn off this device to save power'."] = "Диспетчер устройств → Управление электропитанием → \"Разрешить отключение этого устройства для экономии энергии\".",
            ["Unchecked sets PnPCapabilities bit 0x08 (do-not-turn-off) and clears MSPower_DeviceEnable on the wired NIC."] = "Снятый флажок устанавливает бит PnPCapabilities 0x08 и отключает MSPower_DeviceEnable у проводного сетевого адаптера.",
            ["Applied with APPLY / AUTO. A reboot may be required."] = "Применяется через ПРИМЕНИТЬ или АВТООПТИМИЗАЦИЮ. Может потребоваться перезагрузка.",
            ["Failed to write ReservedCpuSets."] = "Не удалось записать ReservedCpuSets.",
            ["See the log for details."] = "Подробности сохранены в логе.",
            ["DIAGNOSTIC LOGGING IS UNAVAILABLE"] = "НЕ УДАЕТСЯ ЗАПИСАТЬ ЛОГ",
            ["The program can continue, but the log may be incomplete. Include this message when reporting an issue."] = "Программа продолжит работу, но лог может быть неполным. Приложите это сообщение, если будете сообщать о баге.",
            ["The driver check did not complete."] = "Проверка драйвера не завершена.",
            ["DTIMOD driver was unavailable."] = "Драйвер DTIMOD недоступен.",
            ["NIC ITR DRIVER"] = "ДРАЙВЕР NIC ITR",
            ["Backup saved."] = "Бекап сохранён.",
            ["No backup was created."] = "Бекап не создан.",
            ["No backup files were found in the EXE folder or APPDATA."] = "Файлы бекапов не найдены ни в папке EXE, ни в APPDATA.",
            ["Restart Windows to finish restoring settings."] = "Перезагрузите Windows, чтобы завершить восстановление настроек.",
            ["BACKUP FILE"] = "ФАЙЛ БЕКАПА",
            ["Full path (select to copy):"] = "Полный путь (можно выделить и скопировать):",
            ["Driver loading and anti-cheat compatibility are not checked here."] = "Загрузка драйвера и совместимость с античитами здесь не проверяются.",
            ["Not configured: no registry value; this does not mean disabled."] = "Не задано: параметра нет в реестре. Это не означает \"выключено\".",
            ["Backup deleted. No backups remain."] = "Бекап удалён. Других бекапов нет.",
            ["Backup deleted. Select another backup or close this window."] = "Бекап удалён. Выберите другой бекап или закройте окно.",
            ["The backup could not be deleted. See LOGS for details."] = "Не удалось удалить бекап. Причину смотрите в логах.",
            ["Skipping this copy keeps the original-settings backup."] = "Исходный бекап сохраняется и при пропуске.",
            ["Restore the selected backup?"] = "Восстановить выбранный бекап?",
            ["Backup restored."] = "Бекап восстановлен.",
            ["Please reboot your PC to finish applying restored settings."] = "Перезагрузите компьютер, чтобы завершить применение восстановленных настроек.",
            ["Restore did not finish. Any completed rollback steps were preserved."] = "Восстановление завершено не полностью. Уже выполненный откат сохранен.",
            ["The backup could not be created. No changes were made."] = "Не удалось создать бекап. Настройки не изменены.",
            ["NIC ITR was cancelled before any changes were made."] = "Настройка NIC ITR отменена до внесения изменений.",
            ["RAW INPUT THROTTLE"] = "ОГРАНИЧЕНИЕ RAW INPUT",
            ["The setting was not changed."] = "Настройка не изменена.",
            ["The setting could not be updated."] = "Не удалось обновить настройку.",
            ["APPLY was cancelled because the automatic backup failed."] = "Применение отменено: не удалось создать автоматический бекап.",
            ["APPLY preview completed."] = "Предпросмотр применения завершён.",
            ["Sandbox dry-run is ON (no registry changes)."] = "Тестовый режим включён: реестр не изменялся.",
            ["Applied steps were saved. One or more steps were not completed."] = "Примененные настройки сохранены. Часть изменений не применена.",
            ["USB IMOD tuning is available for detected XHCI controller(s)."] = "Для обнаруженных XHCI-контроллеров доступна настройка USB IMOD.",
            ["DTIMOD driver access can be blocked by Windows security features, antivirus software, or anti-cheats."] = "Доступ к драйверу DTIMOD могут блокировать средства безопасности Windows, антивирус или античит.",
            ["Apply it during AUTO-OPTIMIZATION?"] = "Применить USB IMOD во время АВТООПТИМИЗАЦИИ?",
            ["USB IMOD TUNING"] = "НАСТРОЙКА USB IMOD",
            ["AUTO-OPTIMIZATION was cancelled because the automatic backup failed."] = "Автооптимизация отменена: не удалось создать автоматический бекап.",
            ["Original-state backup"] = "Бекап исходного состояния",
            ["the initial recovery snapshot could not be created; changes were not applied"] = "не удалось создать исходный снимок восстановления; изменения не применены",
            ["AUTO-OPTIMIZATION was cancelled because the original-state backup failed."] = "Автооптимизация отменена: не удалось сохранить исходное состояние.",
            ["IMOD apply was cancelled because the automatic backup failed."] = "Применение IMOD отменено: не удалось создать автоматический бекап.",
            ["No eligible USB IMOD target for SET."] = "Подходящий USB-контроллер для применения IMOD не найден.",
            ["IMOD applied."] = "IMOD применён.",
            ["USB IMOD completed with warnings or failed steps."] = "USB IMOD завершён с предупреждениями или ошибками.",
            ["USB IMOD was applied with warnings."] = "USB IMOD применён с предупреждениями.",
            ["No eligible target was found."] = "Подходящее устройство не найдено.",
            ["IMOD note: high-rate input detected; Mouse AUTO uses 0, verify any manual non-zero interval."] = "Примечание IMOD: обнаружено устройство ввода с высокой частотой; для мыши AUTO использует 0. Если задаёте ненулевой интервал вручную, проверьте результат.",
            ["likely MSI/MSI-X"] = "вероятно MSI/MSI-X",
            ["likely Line"] = "вероятно Line",
            ["unknown (registry disabled)"] = "неизвестно (в реестре отключено)",
            ["unknown (registry enabled)"] = "неизвестно (в реестре включено)",
            ["registry key is unavailable"] = "ключ реестра недоступен",
            ["failed to clear ReservedCpuSets in registry"] = "не удалось удалить ReservedCpuSets из реестра",
            ["failed to write the staged ReservedCpuSets value"] = "не удалось записать выбранное значение ReservedCpuSets",
            ["USB IMOD was not applied."] = "USB IMOD не применён.",
            ["NDIS class key was not found"] = "ключ класса NDIS не найден",
            ["NDIS class key is unavailable"] = "ключ класса NDIS недоступен",
            ["Windows CPU topology APIs failed; no affinity changes were applied"] = "не удалось определить топологию CPU через Windows API; настройки Affinity не изменены",
            ["CPU map unavailable; affinity plan was not built"] = "карта CPU недоступна; план Affinity не составлен",
            ["no performance P-cores available for affinity plan"] = "нет доступных производительных P-cores для плана Affinity",
            ["Some power settings require a device restart."] = "Для части настроек питания требуется перезапуск устройства.",
            ["No suitable physical core was found."] = "Подходящее физическое ядро не найдено.",
            ["The register value could not be verified."] = "Не удалось подтвердить значение регистра.",
            ["Failed to disable Vulnerable Driver Blocklist."] = "Не удалось отключить Vulnerable Driver Blocklist.",
            ["Sandbox dry-run is ON."] = "Включён тестовый режим.",
            ["No registry or driver writes are performed."] = "Запись в реестр и драйвер не выполняется.",
            ["This developer build cannot apply settings."] = "Эта сборка для разработчика не применяет настройки.",
            ["off"] = "выкл.",
            ["unsupported"] = "не поддерживается",
            ["read successfully"] = "считано",
            ["read failed"] = "ошибка чтения",
            ["test profile"] = "тестовый профиль",
            ["unknown"] = "неизвестно",
            ["IMOD delete"] = "Удаление IMOD",
            ["IMOD reset to defaults."] = "IMOD сброшен к стандартным значениям.",
            ["Startup script removed. Reboot your PC to unload DTIMOD.sys from memory if it was loaded."] = "Скрипт автозапуска удалён. Перезагрузите компьютер, чтобы выгрузить DTIMOD.sys из памяти, если драйвер был загружен.",
            ["IMOD delete finished with errors. Some persistence files may still remain. If the driver was loaded, reboot to unload DTIMOD.sys from memory."] = "Удаление IMOD завершено с ошибками. Некоторые файлы автозапуска могли сохраниться. Если драйвер был загружен, перезагрузите компьютер для его выгрузки.",
            ["Dry-run reset complete (UI only)."] = "Тестовый сброс завершён только в интерфейсе.",
            ["No registry / IMOD startup files were changed."] = "Реестр и файлы автозапуска IMOD не изменялись.",
            ["This developer build cannot reset Windows settings."] = "Эта сборка для разработчика не сбрасывает настройки Windows.",
            ["RESET WINDOWS DEFAULT preview finished with errors."] = "Предпросмотр сброса Windows завершён с ошибками.",
            ["RESET WINDOWS DEFAULT PREVIEW"] = "ПРЕДПРОСМОТР СБРОСА WINDOWS",
            ["(affinity masks not supported on SSD/HDD)"] = "(affinity masks не поддерживаются для накопителей)",
            ["Scanning devices..."] = "Определение устройств...",
            ["Clearing device list..."] = "Очистка списка устройств...",
            ["Enumerating devices..."] = "Поиск устройств...",
            ["Building reserved CPU sets..."] = "Подготовка reserved CPU sets...",
            ["Laying out devices..."] = "Подготовка интерфейса...",
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
            ["On supported Windows builds, eligible background Raw Input listeners can be throttled."] = "На поддерживаемых сборках Windows ограничение может применяться к фоновым обработчикам Raw Input.",
            ["Throttling and coalescing can reduce the message load from background listeners."] = "Ограничение частоты и объединение событий могут снизить нагрузку от сообщений фоновых обработчиков.",
            ["Mice with a polling rate from 1 to 8 kHz can create a high message load."] = "Мыши с частотой опроса от 1 до 8 кГц могут создавать высокую нагрузку сообщениями.",
            ["This setting does not target foreground apps."] = "Эта настройка не нацелена на активные приложения.",
            ["Some background listeners can bypass this limit."] = "Некоторые фоновые приложения могут обходить это ограничение.",
            ["The setting writes RawMouseThrottleDuration to HKCU\\Control Panel\\Mouse."] = "Настройка записывает RawMouseThrottleDuration в HKCU\\Control Panel\\Mouse.",
            ["DEVICE TWEAKER accepts DWORD values from 1 to 20 milliseconds."] = "DEVICE TWEAKER принимает значения DWORD от 1 до 20 миллисекунд.",
            ["A larger value requests a longer throttle interval for background listeners."] = "Большее значение задаёт более длинный интервал ограничения для фоновых обработчиков.",
            ["Nominal intervals: 20 ms ≈ 50 Hz, 8 ms ≈ 125 Hz, 1 ms ≈ 1 kHz. Actual event rates also depend on input traffic and leeway."] = "Номинальные интервалы: 20 мс ≈ 50 Гц, 8 мс ≈ 125 Гц, 1 мс ≈ 1 кГц. Фактическая частота событий также зависит от входного потока и leeway.",
            ["Background raw input throttling appeared in the Windows 11 22H2 preview update KB5027303."] = "Ограничение фонового Raw Input появилось в предварительном обновлении Windows 11 22H2 KB5027303.",
            ["The change was included in the July 2023 update KB5028185."] = "Изменение вошло в обновление KB5028185 за июль 2023 года.",
            ["DEVICE TWEAKER does not offer this setting on Windows 10 or an older unpatched Windows 11 22H2 build."] = "DEVICE TWEAKER не предлагает эту настройку в Windows 10 и старых сборках Windows 11 22H2 без необходимых обновлений.",
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
        ("Backup files deleted:", "Удалено файлов бекапов:"),
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
        (": set affinity", ": настройка Affinity"),
        (": clear RSS IRQ policy", ": сброс RSS IRQ policy"),
        (": set RSS base CPU", ": настройка RSS Base CPU"),
        (": set RSS queue count", ": настройка RSS Queues"),
        (": set extended RSS values", ": настройка параметров RSS"),
        (": clear RSS base CPU", ": сброс RSS Base CPU"),
        (": clear RSS queue count", ": сброс RSS Queues"),
        (": clear extended RSS values", ": сброс параметров RSS"),
        (": reset", ": сброс"),
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

    [GeneratedRegex("^(\\d+) targets$", RegexOptions.CultureInvariant)]
    private static partial Regex TargetsRegex();

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

    [GeneratedRegex("^version (.+) - developed by (@[^ ]+)$", RegexOptions.CultureInvariant)]
    private static partial Regex SubtitleRegex();

    [GeneratedRegex("^Value: ReservedCpuSets = (.+) \\| CPUs: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex CpuValueRegex();

    [GeneratedRegex(@"^Apply changes \(Ctrl\+S\)\s*[-—:] modified: (\d+)$", RegexOptions.CultureInvariant)]
    private static partial Regex ApplyChangesDirtyRegex();

    [GeneratedRegex(@"^pending APPLY: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex PendingApplyRegex();

    [GeneratedRegex(@"^(Value|Pending APPLY): ReservedCpuSets = (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex ReservedValueRegex();

    [GeneratedRegex(@"^RSS: (Enabled|Disabled|Unknown) base (\S+) queues (\S+)$", RegexOptions.CultureInvariant)]
    private static partial Regex RssRuntimeSummaryRegex();

    [GeneratedRegex(@"^IRQ: (.+) \| WMI hint: (.+) \| HW: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex IrqSummaryRegex();

    [GeneratedRegex(@"^IRQ: (.+) \| HW: (.+)$", RegexOptions.CultureInvariant)]
    private static partial Regex IrqHardwareSummaryRegex();

    [GeneratedRegex(@"^IMOD applied to (\d+) USB controller\(s\)\.$", RegexOptions.CultureInvariant)]
    private static partial Regex ImodAppliedRegex();

    [GeneratedRegex(@"^USB IMOD was applied with warnings \((\d+)/(\d+) writes\)\.$", RegexOptions.CultureInvariant)]
    private static partial Regex ImodWarningRegex();

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
