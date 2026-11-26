using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public class UiLanguageTests
{
    [Theory]
    [InlineData("\n", false)]
    [InlineData("\r\n", false)]
    [InlineData("\n", true)]
    [InlineData("\r\n", true)]
    public void DriverMessagesTranslateWithEitherLineEnding(string newline, bool pending)
    {
        string source = pending
            ? "DTIMOD loading did not finish normally.\nRestart Windows, then press CHECK."
            : DriverRuntime.RestartMessage;
        source = source.Replace("\n", newline, StringComparison.Ordinal);
        var translate = typeof(UiLanguage).GetMethod("TranslateRussian", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!;
        string actual = (string)translate.Invoke(null, [source])!;
        string firstLine = pending ? "Загрузка DTIMOD не завершилась штатно." : "DTIMOD недоступен.";
        Assert.Equal(firstLine + newline + "Перезагрузите Windows и нажмите ПРОВЕРИТЬ.", actual);
    }

    [Theory]
    [InlineData("The operation stopped before any changes were made.", "Настройки не изменены: операция остановлена до применения.")]
    [InlineData("Apply changes (Ctrl+S) — modified: 3", "Изменено: 3")]
    [InlineData("Apply changes (Ctrl+S): modified: 3", "Изменено: 3")]
    [InlineData("pending APPLY: off", "ожидает ПРИМЕНИТЬ: выкл.")]
    [InlineData("pending APPLY: 1K (DWORD=1)", "ожидает ПРИМЕНИТЬ: 1K (DWORD=1)")]
    [InlineData("Pending APPLY: ReservedCpuSets = not set", "Ожидает ПРИМЕНИТЬ: ReservedCpuSets = не задано")]
    [InlineData("Pending APPLY: ReservedCpuSets = [02] | CPUs: 1", "Ожидает ПРИМЕНИТЬ: ReservedCpuSets = [02] | CPU: 1")]
    [InlineData("HID: Mouse 1K, Keyboard 8K", "HID: Мышь 1K, Клавиатура 8K")]
    [InlineData("Polling live: Mouse 1K", "Опрос live: Мышь 1K")]
    [InlineData("Polling: Mouse: 8K, Keyboard: 1K", "Опрос: Мышь: 8K, Клавиатура: 1K")]
    [InlineData("Polling live: Mouse: 1K", "Опрос live: Мышь: 1K")]
    [InlineData("Topology: estimated CHIP 0 | CPU-direct | Alder Lake (12th Gen)", "Топология: ориентировочно CHIP 0 | CPU-direct | Alder Lake (12th Gen)")]
    [InlineData("RSS: Enabled base 0 queues 4", "RSS: Включено | Базовый CPU: 0 | Очереди RSS: 4")]
    [InlineData("Power Saving: Disabled", "Энергосбережение: Выключено")]
    [InlineData("IRQ: 1 | WMI hint: likely MSI/MSI-X | HW: Line/MSI/MSI-X (max 3)", "WMI hint: вероятно MSI/MSI-X")]
    [InlineData("IRQ: unavailable | WMI hint: Unknown | HW: unknown", "IRQ: недоступно | WMI hint: Неизвестно | HW: неизвестно")]
    [InlineData("IMOD applied to 2 USB controller(s).", "IMOD применён на USB-контроллерах: 2.")]
    [InlineData("USB IMOD was applied with warnings (1/2 writes).", "USB IMOD применён с предупреждениями (записано 1/2).")]
    [InlineData("Failed to disable Vulnerable Driver Blocklist.\nWindows error 5", "Не удалось отключить Vulnerable Driver Blocklist.\nWindows error 5")]
    [InlineData("Unknown DTIMOD driver load error.", "Не удалось загрузить DTIMOD. Причина не определена.")]
    [InlineData("no devices to optimize", "Нет устройств для автооптимизации.")]
    [InlineData("current: saved; moderation setup incomplete", "настройка модерации не завершена")]
    [InlineData("4 targets", "Устройств: 4")]
    [InlineData("12 targets", "Устройств: 12")]
    [InlineData("RESET", "СБРОСИТЬ")]
    [InlineData("IRQ: unavailable | HW: unknown", "IRQ: недоступно | HW: неизвестно")]
    [InlineData("not read", "не считан")]
    [InlineData("An older DTIMOD driver is loaded. Restart Windows, then press CHECK again.", "Загружена предыдущая версия DTIMOD. Перезагрузите Windows, затем нажмите ПРОВЕРИТЬ.")]
    [InlineData("no IRQ", "нет IRQ")]
    [InlineData("MSI Limit exceeds the hardware maximum reported by Windows (2).", "MSI Limit превышает аппаратный максимум, указанный Windows (2).")]
    [InlineData("MSI Limit exceeds the hardware maximum reported by Windows (2048). The invalid value was not applied.", "MSI Limit превышает аппаратный максимум, указанный Windows (2048). Некорректное значение не применено.")]
    [InlineData("This is an MSI-only device. The allowed multi-message limits are 1, 2, 4, 8, or 16.", "Устройство поддерживает только MSI. Допустимые лимиты сообщений: 1, 2, 4, 8 или 16.")]
    [InlineData("pre-restore", "до восстановления")]
    [InlineData("pre-release", "перед обновлением")]
    [InlineData("pre-imod", "до настройки IMOD")]
    [InlineData("pre-nic-itr", "до настройки NIC ITR")]
    [InlineData("pre-nic-itr-save", "до сохранения NIC ITR")]
    [InlineData("Refresh interrupt mode", "Обновить режим прерываний")]
    [InlineData("Read interrupt modes for all PCI devices", "Считать режимы прерываний всех PCI-устройств")]
    [InlineData("READ IRQ MODE", "СЧИТАТЬ IRQ")]
    [InlineData("Service stopped", "Остановлена")]
    [InlineData("No registry setting", "Нет параметра")]
    [InlineData("Simulated interrupt mode. No hardware was queried.", "Тестовый режим прерываний. Железо не опрашивалось.")]
    [InlineData("current: 0x0 x8 · 14:02", "текущее: 0x0 x8 · 14:02")]
    [InlineData("current: 0x0 x8 · 14:02 · devices changed", "текущее: 0x0 x8 · 14:02 · устройства изменились")]
    [InlineData("Last read: 14:02", "Считано: 14:02")]
    [InlineData("Saved settings were restored.", "Сохранённые настройки восстановлены.")]
    [InlineData("Saved settings", "Сохранённые настройки")]
    [InlineData("Saved setting", "Сохранённый параметр")]
    [InlineData("Registry and startup profile restored and verified.", "Реестр и профиль автозапуска восстановлены и проверены.")]
    [InlineData("This setting was unreadable when the backup was created. Its current value was kept.", "оставлено без изменений")]
    [InlineData("Power scheme", "Схема электропитания")]
    [InlineData("The restored power scheme could not be activated.", "Не удалось активировать восстановленную схему электропитания.")]
    [InlineData("Some device data could not be refreshed within the time limit.", "за отведённое время")]
    [InlineData("Some device settings could not be applied.", "Не все настройки устройства удалось применить.")]
    [InlineData("The device list could not be refreshed.", "Не удалось обновить список устройств.")]
    public void UserStatusesTranslateWhilePreservingNumbersAndDiagnosticText(string source, string required)
    {
        var method = typeof(UiLanguage).GetMethod("TranslateRussian", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!;
        string translated = (string)method.Invoke(null, [source])!;
        Assert.Contains(required, translated, StringComparison.Ordinal);
    }

    // Controls keep their source text, but some code assigns already translated text;
    // translating it again must not change it (e.g. "[ ВСЕ ]" lost its spaces).
    [Fact]
    public void TranslatingRussianTextAgainChangesNothing()
    {
        var flags = System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic;
        var translate = typeof(UiLanguage).GetMethod("TranslateRussian", flags)!;
        var dictionary = (IReadOnlyDictionary<string, string>)typeof(UiLanguage).GetField("Russian", flags)!.GetValue(null)!;
        List<string> changed = [];
        foreach (string russian in dictionary.Values.Distinct())
        {
            string again = (string)translate.Invoke(null, [russian])!;
            if (!string.Equals(again, russian, StringComparison.Ordinal))
            {
                changed.Add($"{russian} -> {again}");
            }
        }

        Assert.True(changed.Count == 0, string.Join("\n", changed));
    }

    [Fact]
    public void RussianTranslationContract_PreservesExpectedTerms()
    {
        Assert.True(UiLanguage.ValidateRussianTranslationContract(out string error), error);
    }
}
