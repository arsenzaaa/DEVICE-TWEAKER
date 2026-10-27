using DeviceTweakerCS;
using Xunit;

namespace DeviceTweaker.Tests;

public class UiLanguageTests
{
    [Theory]
    [InlineData("The operation stopped before any changes were made.", "Настройки не изменены: операция остановлена до применения.")]
    [InlineData("Apply changes (Ctrl+S) — modified: 3", "изменено: 3")]
    [InlineData("Apply changes (Ctrl+S): modified: 3", "изменено: 3")]
    [InlineData("pending APPLY: off", "ожидает ПРИМЕНИТЬ: выкл.")]
    [InlineData("pending APPLY: 1K (DWORD=1)", "ожидает ПРИМЕНИТЬ: 1K (DWORD=1)")]
    [InlineData("Pending APPLY: ReservedCpuSets = not set", "Ожидает ПРИМЕНИТЬ: ReservedCpuSets = не задано")]
    [InlineData("Pending APPLY: ReservedCpuSets = [02] | CPUs: 1", "Ожидает ПРИМЕНИТЬ: ReservedCpuSets = [02] | CPU: 1")]
    [InlineData("HID: Mouse 1K, Keyboard 8K", "HID: Мышь 1K, Клавиатура 8K")]
    [InlineData("Polling live: Mouse 1K", "Опрос live: Мышь 1K")]
    [InlineData("Polling: Mouse: 8K, Keyboard: 1K", "Опрос: Мышь: 8K, Клавиатура: 1K")]
    [InlineData("Polling live: Mouse: 1K", "Опрос live: Мышь: 1K")]
    [InlineData("Topology: estimated CHIP 0 | CPU-direct | Alder Lake (12th Gen)", "Топология: ориентировочно CHIP 0 | CPU-direct | Alder Lake (12th Gen)")]
    [InlineData("RSS: Enabled base 0 queues 4", "RSS: Включено | RSS Base: 0 | RSS Queues: 4")]
    [InlineData("Power Saving: Disabled", "Power Saving: Выключено")]
    [InlineData("IRQ: 1 | WMI hint: likely MSI/MSI-X | HW: Line/MSI/MSI-X (max 3)", "WMI hint: вероятно MSI/MSI-X")]
    [InlineData("IRQ: unavailable | WMI hint: Unknown | HW: unknown", "IRQ: недоступно | WMI hint: Неизвестно | HW: неизвестно")]
    [InlineData("IMOD applied to 2 USB controller(s).", "IMOD применён на USB-контроллерах: 2.")]
    [InlineData("USB IMOD was applied with warnings (1/2 writes).", "USB IMOD применён с предупреждениями (записано 1/2).")]
    [InlineData("Failed to disable Vulnerable Driver Blocklist.\nWindows error 5", "Не удалось отключить Vulnerable Driver Blocklist.\nWindows error 5")]
    [InlineData("Unknown DTIMOD driver load error.", "Не удалось загрузить DTIMOD. Причина не определена.")]
    [InlineData("no devices to optimize", "Нет устройств для авто-оптимизации.")]
    [InlineData("4 targets", "Устройств: 4")]
    [InlineData("12 targets", "Устройств: 12")]
    [InlineData("RESET", "СБРОСИТЬ")]
    [InlineData("IRQ: unavailable | HW: unknown", "IRQ: недоступно | HW: неизвестно")]
    [InlineData("not read", "не считан")]
    [InlineData("An older DTIMOD driver is loaded. Restart Windows, then press CHECK again.", "Загружена предыдущая версия DTIMOD. Перезагрузите Windows, затем нажмите ПРОВЕРИТЬ.")]
    [InlineData("no IRQ", "нет IRQ")]
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
    public void UserStatusesTranslateWhilePreservingNumbersAndDiagnosticText(string source, string required)
    {
        var method = typeof(UiLanguage).GetMethod("TranslateRussian", System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic)!;
        string translated = (string)method.Invoke(null, [source])!;
        Assert.Contains(required, translated, StringComparison.Ordinal);
    }

    [Fact]
    public void RussianTranslationContract_PreservesExpectedTerms()
    {
        Assert.True(UiLanguage.ValidateRussianTranslationContract(out string error), error);
    }
}
