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

        // The value is the driver's own option text, already in the system language.
        if (source.StartsWith("Interrupt Moderation: ", StringComparison.Ordinal))
            return source;

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
        bool changed = false;
        string translated = string.Join(", ", content.Split(',', StringSplitOptions.TrimEntries).Select(part =>
        {
            string[] pipeParts = part.Split('|', StringSplitOptions.TrimEntries);
            return string.Join(" | ", pipeParts.Select(role =>
            {
                string russian = TranslateRussianRole(role);
                changed |= !string.Equals(russian, role, StringComparison.Ordinal);
                return russian;
            }));
        }));
        // Brackets without a role ("[ ВСЕ ]", "[02]") keep their exact spacing.
        return changed ? value[..(open + 1)] + translated + value[close..] : value;
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

    private static readonly IReadOnlyDictionary<string, string> Russian = UiResources.LoadRussianPhrases();

    private static readonly (string English, string Russian)[] InlineReplacements = UiResources.LoadInlineReplacements();

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
