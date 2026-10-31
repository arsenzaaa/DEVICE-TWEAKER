<div align="right"><a href="./README.en.md">English</a></div>

# DEVICE TWEAKER

Настройка прерываний устройств в Windows 10 / 11 x64. MSI, affinity, приоритет IRQ, ReservedCpuSets, RSS, USB IMOD и NIC ITR в одной программе, с авто-оптимизацией под топологию процессора: P- и E-ядра, SMT, CCD / CCX, CPPC.

**[Скачать 0.0.4](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest)** · [Описание настроек](./src/docs/SETTINGS.md) · [История изменений](./CHANGELOG.md)

<a href="./assets/screenshots/showcase_intel_14900k_full_ru.png"><img src="./assets/screenshots/showcase_intel_14900k_full_ru.png" alt="DEVICE TWEAKER: USB-контроллер на Intel Core i9-14900K" width="100%"></a>

<table>
  <tr>
    <td width="33%"><a href="./assets/screenshots/showcase_amd_imod_full_ru.png"><img src="./assets/screenshots/thumbs/showcase_amd_imod_full_ru.png" alt="USB IMOD на AMD Ryzen 9 9950X3D"></a></td>
    <td width="33%"><a href="./assets/screenshots/showcase_gpu_full_ru.png"><img src="./assets/screenshots/thumbs/showcase_gpu_full_ru.png" alt="Прерывания GPU"></a></td>
    <td width="33%"><a href="./assets/screenshots/showcase_nic_full_ru.png"><img src="./assets/screenshots/thumbs/showcase_nic_full_ru.png" alt="NIC ITR и RSS"></a></td>
  </tr>
  <tr>
    <td><sub>USB IMOD, Ryzen 9 9950X3D</sub></td>
    <td><sub>GPU, пара физических ядер</sub></td>
    <td><sub>NIC ITR и RSS, Intel I226-V</sub></td>
  </tr>
</table>

<sub>На скриншотах тестовые устройства и значения.</sub>

## Что меняет программа

| Настройка | Куда записывается | Когда действует |
| --- | --- | --- |
| MSI Mode, MSI Limit | `MSISupported`, `MessageNumberLimit` в `Interrupt Management\MessageSignaledInterruptProperties` | после перезагрузки |
| IRQ Priority | `DevicePriority` в `Interrupt Management\Affinity Policy` | после перезагрузки |
| Policy, Interrupt Affinity | `DevicePolicy`, `AssignmentSetOverride` там же | после перезагрузки |
| RSS | `*RssBaseProcNumber`, `*NumRssQueues`, `*MaxRssProcessors`, `*RssMaxProcNumber` в ключе класса адаптера | после перезагрузки |
| ReservedCpuSets | `ReservedCpuSets` в `Session Manager\kernel` | после перезагрузки |
| Mouse Throttle | `RawMouseThrottleDuration` в `HKCU\Control Panel\Mouse`, Windows 11 | после перезагрузки |
| Энергосбережение | `PnPCapabilities`, `SelectiveSuspendEnabled`, `EnhancedPowerManagementEnabled`, USB selective suspend в плане питания | после перезагрузки |
| USB IMOD | регистр IMOD каждого interrupter'а контроллера xHCI | сразу, до перезагрузки или сна |
| NIC ITR | регистры модерации контроллера: EITR / ITR у Intel, IntrMit у Realtek | сразу, до перезагрузки или сна |

Перед применением создаётся проверенный бекап. Регистры IMOD и NIC ITR после записи читаются повторно. Что применилось, что пропущено и почему — в логе.

## Авто-оптимизация

Строит план по топологии процессора и ролям устройств.

- Дискретная GPU получает два разных физических ядра. SMT-соседи одного ядра не считаются двумя ядрами.
- Контроллер с мышью и клавиатурой, сеть и аудио размещаются по своим правилам. На процессорах с несколькими CCD выбирается целевой CCD.
- Накопителям включается MSI и высокий приоритет, affinity остаётся за Windows. Встроенной GPU affinity возвращается Windows, Wi-Fi не трогается.
- Если топология определена ненадёжно, запись не начинается.

Перед первой авто-оптимизацией сохраняется исходное состояние. Его можно вернуть в любой момент.

## USB IMOD и NIC ITR

IMOD — интервал модерации прерываний interrupter'а xHCI с шагом 250 нс: `0xC8` = 50 мкс, `0x0` выключает модерацию. Значение задаётся на весь контроллер, по interrupter'ам или по ролям устройств. Карта «устройство → interrupter» читается из контекстов xHCI в памяти.

NIC ITR поддерживается для Intel I210 / I211, I219, I225 / I226, I350, 82576 / 82580 и Realtek RTL8111 / 8168, RTL8125 / 8126. При сохранении профиля в реестре выключается собственная модерация драйвера (`*InterruptModeration`), иначе она перезаписывает регистр.

Регистры пишет собственный драйвер **DTIMOD.sys**, и только в память, которую Windows выделила выбранному устройству. Значения в регистрах живут до перезагрузки или сна, поэтому сохранённый профиль применяет задача Планировщика `\DEVICE TWEAKER\IMOD`: при входе в Windows, после выхода из сна и при подключении USB. Задача запускает открытый скрипт `ApplyIMOD.ps1` из `%ProgramData%\DEVICE TWEAKER\Startup` и работает без EXE.

## Установка

| Файл | Требования |
| --- | --- |
| [DEVICE.TWEAKER.SelfContained.exe](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe) | нет, .NET встроен |
| [DEVICE.TWEAKER.exe](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.exe) | [.NET 8 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/8.0) |

Запускайте от имени администратора. DTIMOD.sys и загрузчик KDU встроены в обе сборки. SHA-256 файлов — в `SHA256SUMS.txt` на [странице релиза](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest).

1. **АВТООПТИМИЗАЦИЯ** строит план и применяет его. Если найден подходящий контроллер xHCI, программа предложит включить USB IMOD. Ручная настройка — **ПРИМЕНИТЬ**.
2. Перезагрузите компьютер.
3. **ПРОВЕРИТЬ** в блоке IMOD или NIC ITR считывает фактические значения регистров.

## Откат

**ВОССТАНОВЛЕНИЕ → ИСХОДНЫЙ** возвращает реестр к состоянию до первой авто-оптимизации и удаляет задачу автозапуска. Там же — любой из бекапов и сброс к настройкам Windows. Бекапы хранятся в папке `Backups` рядом с EXE.

## Вопросы

<details>
<summary>Зачем драйвер и почему антивирус может ругаться на KDU</summary>
<br>

Регистры IMOD и NIC ITR недоступны из пользовательского режима. У DTIMOD.sys нет публичной подписи, поэтому его загружает [KDU](./src/IMOD/KDU/README.md) через подписанный уязвимый драйвер, а антивирусы относят такие загрузчики к HackTool. Исходники DTIMOD и закреплённая версия KDU лежат в репозитории, бинарники сверяются по SHA-256. Драйвер нужен только для IMOD, NIC ITR и **СЧИТАТЬ IRQ**.
</details>

<details>
<summary>HVCI, Vulnerable Driver Blocklist, FACEIT и Vanguard</summary>
<br>

Целостность памяти (HVCI) и Microsoft Vulnerable Driver Blocklist блокируют загрузку через KDU, античиты FACEIT и Vanguard могут с ней конфликтовать. Перед включением USB IMOD программа показывает их состояние в **ПРОВЕРКЕ СИСТЕМЫ**. Настройки реестра от драйвера не зависят.
</details>

## Обратная связь

Ошибки — в [issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml), с логом из папки `logs` рядом с EXE, версией Windows и конфигурацией ПК. Для авто-оптимизации полезнее всего три лога: до применения, сразу после и после перезагрузки. Новости — в [Telegram](https://t.me/arsenzaa).

[Описание настроек](./src/docs/SETTINGS.md) · [Сборка](./src/docs/BUILD.md) · [Проверки](./dev/qa/README.md) · [Безопасность](./.github/SECURITY.md) · [GPL-3.0](./LICENSE)
