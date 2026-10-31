<p align="right"><a href="./README.en.md">English</a></p>

<h1>
<picture>
<source media="(max-width: 600px) and (prefers-color-scheme: dark)" srcset="./assets/branding/hero-mobile-dark.svg">
<source media="(max-width: 600px)" srcset="./assets/branding/hero-mobile-light.svg">
<source media="(prefers-color-scheme: dark)" srcset="./assets/branding/hero-dark.svg">
<img src="./assets/branding/hero-light.svg" alt="DEVICE TWEAKER — Windows 10 / 11 x64: MSI, affinity, RSS, IMOD, ITR" width="1080">
</picture>
</h1>

<p align="center"><strong>DEVICE TWEAKER настраивает прерывания устройств в Windows 10/11 x64.</strong></p>

<p align="center">MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + RSS + управление IMOD / ITR в одном интерфейсе с интеллектуальной авто-оптимизацией под топологию процессора.</p>

<p align="center">
<a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/branding/button-download-ru-dark.svg"><img src="./assets/branding/button-download-ru-light.svg" alt="Скачать 0.0.4" width="169" height="40"></picture></a>
<a href="./docs/SETTINGS.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/branding/button-settings-ru-dark.svg"><img src="./assets/branding/button-settings-ru-light.svg" alt="Описание настроек" width="132" height="40"></picture></a>
<a href="./CHANGELOG.md"><picture><source media="(prefers-color-scheme: dark)" srcset="./assets/branding/button-changelog-ru-dark.svg"><img src="./assets/branding/button-changelog-ru-light.svg" alt="История изменений" width="132" height="40"></picture></a>
</p>

## <samp>[ ИНТЕРФЕЙС ]</samp>

<p align="center"><a href="./assets/screenshots/showcase_intel_14900k_full_ru.png"><img src="./assets/screenshots/showcase_intel_14900k_full_ru.png" alt="DEVICE TWEAKER: USB IMOD на Intel Core i9-14900K" width="100%"></a></p>

<table>
<tr><th align="left" colspan="3"><samp>Другие конфигурации</samp></th></tr>
<tr>
<td width="33%" valign="top"><a href="./assets/screenshots/showcase_amd_imod_full_ru.png"><img src="./assets/screenshots/thumbs/showcase_amd_imod_full_ru.png" alt="USB IMOD на AMD Ryzen 9 9950X3D"></a><br><sub>USB IMOD, Ryzen 9 9950X3D</sub></td>
<td width="33%" valign="top"><a href="./assets/screenshots/showcase_gpu_full_ru.png"><img src="./assets/screenshots/thumbs/showcase_gpu_full_ru.png" alt="Прерывания GPU на двух физических ядрах"></a><br><sub>GPU, два физических ядра</sub></td>
<td width="33%" valign="top"><a href="./assets/screenshots/showcase_nic_full_ru.png"><img src="./assets/screenshots/thumbs/showcase_nic_full_ru.png" alt="NIC ITR и RSS на Intel I226-V"></a><br><sub>NIC ITR и RSS, Intel I226-V</sub></td>
</tr>
</table>

<sub>На скриншотах тестовые устройства и значения.</sub>

## <samp>[ ФУНКЦИОНАЛЬНОСТЬ ]</samp>

<table>
<tr><th align="left">Интеллектуальная авто-оптимизация</th></tr>
<tr><td>
<p>Учитывает P-Core / E-Core, SMT / HT, CPPC-рейтинг, CCD / CCX, количество устройств и их роли. Дискретная GPU получает два разных физических ядра: SMT-соседи одного ядра за два не считаются.</p>
<p align="center"><a href="./docs/SETTINGS.md#интеллектуальная-авто-оптимизация-и-восстановление"><picture><source media="(max-width: 600px) and (prefers-color-scheme: dark)" srcset="./assets/branding/cpu-allocation-mobile-ru-dark.svg"><source media="(max-width: 600px)" srcset="./assets/branding/cpu-allocation-mobile-ru-light.svg"><source media="(prefers-color-scheme: dark)" srcset="./assets/branding/cpu-allocation-ru-dark.svg"><img src="./assets/branding/cpu-allocation-ru-light.svg" alt="GPU: CPU 4 на ядре 2 и CPU 8 на ядре 4, SMT-соседи CPU 5 и CPU 9 свободны" width="560"></picture></a></p>
<p>Контроллер с мышью и клавиатурой, сеть и аудио распределяются по своим правилам, на процессорах с несколькими CCD выбирается целевой CCD. Накопителям включается MSI и высокий приоритет, affinity остаётся за Windows. Встроенной GPU affinity возвращается Windows, Wi-Fi не трогается.</p>
<p>Если топология определена ненадёжно, запись не начинается. Перед первой авто-оптимизацией сохраняется исходное состояние, его возвращает <strong>ВОССТАНОВЛЕНИЕ → ИСХОДНЫЙ</strong>.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Настройка прерываний</th></tr>
<tr><td>
<p>MSI Mode / MSI Limit / IRQ Priority / IRQ Policy / Interrupt Affinity для каждого устройства.</p>
<p>Настройка MSI Mode из реестра и текущий режим прерываний в IRQ Mode показаны отдельно, поддерживаемые устройством режимы — в HW. При открытии программа берёт режим из ресурсов Windows или у уже загруженного DTIMOD.sys и драйвер не загружает. Если по ресурсам Windows нельзя отличить MSI от MSI-X, показывается MSI/MSI-X.</p>
<p><strong>СЧИТАТЬ IRQ</strong> вверху окна читает MSI / MSI-X Enable всех PCI-устройств из списка через DTIMOD.sys и при необходимости загружает драйвер.</p>
</td></tr>
</table>

<table>
<tr><th align="left">USB IMOD</th></tr>
<tr><td>
<p>Читает и меняет IMOD — интервал модерации прерываний каждого interrupter'а контроллера xHCI, шаг 250 нс: <code>0xC8</code> = 50 мкс, <code>0x0</code> выключает модерацию.</p>
<p>Режим <strong>Устройства</strong> задаёт значения по USB-ролям, <strong>XHCI</strong> — одно значение на контроллер, <strong>Interrupters</strong> — по индексам. Карта «устройство → interrupter» читается из контекстов xHCI в памяти. После записи значения считываются повторно.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Настройка сети</th></tr>
<tr><td>
<p>NIC ITR у Intel I210 / I211, I219, I225 / I226, I350, 82576 / 82580 и Realtek RTL8111 / 8168, RTL8125 / 8126, включая совместимые Killer. Очереди и RX / TX показываются в зависимости от семейства контроллера. <strong>СОХРАНИТЬ</strong> выключает в реестре собственную модерацию драйвера (<code>*InterruptModeration</code>), иначе она перезаписывает регистр, <strong>УДАЛИТЬ</strong> возвращает значение драйвера.</p>
<p>RSS у проводных адаптеров NDIS: базовый CPU, число очередей и режим RSS / IRQ. Процессоры берутся по одному на физическое ядро, SMT-соседи пропускаются. Число очередей ограничено значениями драйвера и числом векторов MSI-X.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Автозапуск IMOD / NIC ITR</th></tr>
<tr><td>
<p>Значения в регистрах держатся до перезагрузки или сна. Сохранённый профиль применяет задача Планировщика <code>\DEVICE TWEAKER\IMOD</code> от SYSTEM: при входе в Windows, после выхода из сна и при подключении USB-устройства.</p>
<p>Задача запускает открытый скрипт <code>ApplyIMOD.ps1</code> из <code>%ProgramData%\DEVICE TWEAKER\Startup</code> и работает без EXE. Перед запуском сверяются SHA-256 runtime, драйвера и загрузчика. Изменять файлы могут только администраторы и SYSTEM, журналы хранятся 14 дней.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Определение устройств</th></tr>
<tr><td>
<p>Находит GPU / USB-контроллеры / сетевые адаптеры / аудиоконтроллеры / контроллеры накопителей и показывает их по категориям.</p>
<p>Определяет, на каком USB-контроллере находятся мышь / клавиатура / геймпад / микрофон / аудио / камера. Если к одному контроллеру подключено несколько устройств, показываются все найденные роли.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Топология процессора / ReservedCpuSets</th></tr>
<tr><td>
<p>Определяет P-Core / E-Core, SMT / HT, AMD CCD / CCX и CPPC-рейтинг, если он доступен. Показывает, к какому физическому ядру относится каждый логический поток. Выбранные CPU можно зарезервировать через ReservedCpuSets.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Polling Rate / Mouse Throttle</th></tr>
<tr><td>
<p>Показывает Polling Rate мыши и клавиатуры и замеряет частоту событий Raw Input, после замера выводится измеренная частота. У клавиатуры частота событий зависит от нажатий. RawMouseThrottleDuration настраивается на совместимых сборках Windows 11, текущее значение видно в интерфейсе.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Энергосбережение</th></tr>
<tr><td>
<p>Энергосбережение USB-контроллеров / корневых USB-хабов / проводных сетевых адаптеров и USB selective suspend в плане питания. После записи значения перечитываются через WMI: неподдерживаемая функция, ошибка и несовпадение значений выводятся отдельно.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Бекапы / логи</th></tr>
<tr><td>
<p>Перед применением создаётся проверенный бекап, он включает и профиль автозапуска. Бекап можно восстановить или удалить, настройки — сбросить к значениям Windows через <strong>СБРОС WINDOWS</strong>. Бекапы хранятся в папке <code>Backups</code> рядом с EXE.</p>
<p>Подробные логи определения устройств и применения настроек — что применилось, что пропущено и почему — пишутся в папку <code>logs</code> рядом с EXE.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Интерфейс / навигация</th></tr>
<tr><td>
<p>Поиск по устройствам, фильтры по категориям, навигация с клавиатуры и переключение RU / EN без перезапуска. Изменённые, но ещё не применённые настройки отмечаются индикатором.</p>
</td></tr>
</table>

## <samp>[ ЧТО МЕНЯЕТ ПРОГРАММА ]</samp>

| Настройка | Куда записывается |
| :--- | :--- |
| MSI Mode, MSI Limit | `MSISupported`, `MessageNumberLimit` в `Interrupt Management\MessageSignaledInterruptProperties` |
| IRQ Priority | `DevicePriority` в `Interrupt Management\Affinity Policy` |
| IRQ Policy, Interrupt Affinity | `DevicePolicy`, `AssignmentSetOverride` там же |
| RSS | `*RssBaseProcNumber`, `*NumRssQueues`, `*MaxRssProcessors`, `*RssMaxProcNumber` в ключе класса адаптера |
| ReservedCpuSets | `ReservedCpuSets` в `Session Manager\kernel` |
| Mouse Throttle | `RawMouseThrottleDuration` в `HKCU\Control Panel\Mouse` |
| Энергосбережение | `PnPCapabilities`, `SelectiveSuspendEnabled`, `EnhancedPowerManagementEnabled`, USB selective suspend в плане питания |
| USB IMOD | регистр IMOD каждого interrupter'а контроллера xHCI |
| NIC ITR | EITR / ITR у Intel, IntrMit у Realtek |

Настройки реестра вступают в силу после перезагрузки. Регистры IMOD и NIC ITR меняются сразу и после записи читаются повторно.

## <samp>[ УСТАНОВКА ]</samp>

<table>
<tr><th align="left">Скачать</th></tr>
<tr><td>
<p>В <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest">релизе</a> доступны две сборки:</p>
<ul>
<li><a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.SelfContained.exe"><code>DEVICE.TWEAKER.SelfContained.exe</code></a> — со встроенным .NET, устанавливать ничего не нужно;</li>
<li><a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/latest/download/DEVICE.TWEAKER.exe"><code>DEVICE.TWEAKER.exe</code></a> — обычная сборка, требуется <a href="https://dotnet.microsoft.com/download/dotnet/8.0">.NET 8 Desktop Runtime x64</a>.</li>
</ul>
<p>Драйвер DTIMOD.sys и загрузчик KDU встроены в обе сборки. SHA-256 файлов указаны в <code>SHA256SUMS.txt</code>, исходников — в <code>SOURCE-SHA256SUMS.txt</code>.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Использование</th></tr>
<tr><td>
<p>Запустите программу от имени администратора.</p>
<p><strong>АВТООПТИМИЗАЦИЯ</strong> строит план по топологии процессора и применяет его. Если найден подходящий контроллер xHCI, программа предложит включить USB IMOD. Кнопка <strong>ПРОВЕРКА СИСТЕМЫ</strong> в этом окне показывает состояние HVCI, Blocklist, Microsoft Defender и служб FACEIT / Vanguard. NIC ITR настраивается вручную.</p>
<p>Для ручной настройки выберите параметры и нажмите <strong>ПРИМЕНИТЬ</strong>. Кнопка применяет настройки прерываний, ReservedCpuSets, RSS, Mouse Throttle, энергосбережения и USB IMOD.</p>
<p>У IMOD / NIC ITR есть отдельные кнопки:</p>
<ul>
<li>IMOD: <strong>ЗАДАТЬ</strong> применяет значения и сохраняет отличающиеся от стандартных для автозапуска.</li>
<li>NIC ITR: <strong>ЗАДАТЬ</strong> меняет текущие значения, <strong>СОХРАНИТЬ</strong> сохраняет профиль для автозапуска.</li>
<li><strong>ПРОВЕРИТЬ</strong> считывает фактические значения регистров.</li>
<li>IMOD: <strong>УДАЛИТЬ</strong> убирает общий автозапуск IMOD / NIC ITR и возвращает поля IMOD к стандартному значению. Чтобы завершить сброс, перезагрузите компьютер.</li>
</ul>
<p>После <strong>ПРИМЕНИТЬ</strong> / <strong>АВТООПТИМИЗАЦИЯ</strong> перезагрузите компьютер. <strong>ОБНОВИТЬ</strong> повторно определяет устройства.</p>
<p>Откат — <strong>ВОССТАНОВЛЕНИЕ</strong>: <strong>ИСХОДНЫЙ</strong> возвращает реестр к состоянию до первой авто-оптимизации и удаляет задачу автозапуска. Там же любой из бекапов и сброс к настройкам Windows.</p>
</td></tr>
</table>

<table>
<tr><th align="left">Драйвер / античиты</th></tr>
<tr><td>
<p>Регистры IMOD и NIC ITR недоступны из пользовательского режима. Их пишет собственный драйвер DTIMOD.sys, и только в память, которую Windows выделила выбранному устройству. У драйвера нет публичной подписи, поэтому его загружает <a href="./src/IMOD/KDU/README.md">KDU</a> через подписанный уязвимый драйвер, а антивирусы относят такие загрузчики к HackTool. Исходники DTIMOD и закреплённая версия KDU лежат в репозитории, бинарники сверяются по SHA-256.</p>
<p>Целостность памяти (HVCI) и Microsoft Vulnerable Driver Blocklist блокируют загрузку через KDU, античиты FACEIT и Vanguard могут с ней конфликтовать. Драйвер нужен только для IMOD, NIC ITR и <strong>СЧИТАТЬ IRQ</strong>, настройки реестра от него не зависят.</p>
</td></tr>
</table>

## <samp>[ ОБРАТНАЯ СВЯЗЬ ]</samp>

<table>
<tr><th align="left">Баги / тестирование</th></tr>
<tr><td>
<p>Сообщайте о багах и недоработках в <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml">issues</a>. Прикладывайте лог из папки <code>logs</code> рядом с EXE, версию программы, сборку Windows и конфигурацию ПК. Также напишите, что делали и что произошло.</p>
<p>Если хотите помочь с тестированием интеллектуальной авто-оптимизации, отправьте <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=hardware_report.yml">отчёт о железе</a> с логами в таком порядке:</p>
<ol>
<li>До применения авто-оптимизации или после сброса настроек.</li>
<li>Сразу после применения.</li>
<li>После перезагрузки компьютера и повторного запуска программы.</li>
</ol>
<p>Обновления и обсуждение проекта — в <a href="https://t.me/arsenzaa">Telegram</a>.</p>
</td></tr>
</table>

<p align="center"><samp><a href="./docs/SETTINGS.md">Описание настроек</a> · <a href="./docs/DEVELOPMENT.md">Разработка</a> · <a href="./CHANGELOG.md">История изменений</a> · <a href="./docs/THIRD_PARTY_NOTICES.md">Сторонние компоненты</a> · <a href="./.github/SECURITY.md">Безопасность</a> · <a href="./LICENSE">GPL-3.0</a></samp></p>
