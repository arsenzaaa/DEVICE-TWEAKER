<p align="right"><a href="./README.en.md">English</a></p>

<h1>
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER#gh-dark-mode-only"><picture><source media="(max-width: 600px)" srcset="./assets/branding/hero-mobile-dark.svg"><img src="./assets/branding/hero-dark.svg" alt="DEVICE TWEAKER / Windows 10/11 x64 / 0.0.4" width="1080"></picture></a>
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER#gh-light-mode-only"><picture><source media="(max-width: 600px)" srcset="./assets/branding/hero-mobile-light.svg"><img src="./assets/branding/hero-light.svg" alt="DEVICE TWEAKER / Windows 10/11 x64 / 0.0.4" width="1080"></picture></a>
</h1>

<p align="center"><strong>DEVICE TWEAKER настраивает прерывания устройств в Windows 10/11 x64.</strong></p>

MSI Utility V3 + Interrupt Affinity Policy Tool (IAPT) + ReservedCpuSets + управление IMOD / ITR в одном интерфейсе с интеллектуальной авто-оптимизацией.

<p align="center">
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4#gh-dark-mode-only"><img src="./assets/branding/download-ru-dark.svg" alt="Скачать 0.0.4" width="170"></a>
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4#gh-light-mode-only"><img src="./assets/branding/download-ru-light.svg" alt="Скачать 0.0.4" width="170"></a>
  <a href="./DEVICE%20TWEAKER/docs/SETTINGS.md#gh-dark-mode-only"><img src="./assets/branding/settings-ru-dark.svg" alt="Настройки" width="140"></a>
  <a href="./DEVICE%20TWEAKER/docs/SETTINGS.md#gh-light-mode-only"><img src="./assets/branding/settings-ru-light.svg" alt="Настройки" width="140"></a>
</p>

## <samp>[ ИНТЕРФЕЙС ]</samp>

<p align="center"><a href="./assets/screenshots/showcase_intel_14900k_full_ru.png"><img src="./assets/screenshots/showcase_intel_14900k_full_ru.png" alt="USB IMOD / Intel Core i9-14900K" width="860"></a></p>

<table width="100%">
  <tr><th><samp>USB / AMD</samp></th><th><samp>GPU</samp></th><th><samp>NIC ITR</samp></th></tr>
<tr>
  <td width="33%" valign="top"><a href="./assets/screenshots/showcase_amd_imod_full_ru.png" title="USB IMOD / AMD Ryzen 9 9950X3D"><img src="./assets/screenshots/showcase_amd_imod_full_ru.png" alt="USB IMOD / AMD Ryzen 9 9950X3D" width="100%"></a></td>
  <td width="33%" valign="top"><a href="./assets/screenshots/showcase_gpu_full_ru.png" title="GPU / NVIDIA GeForce RTX 4090"><img src="./assets/screenshots/showcase_gpu_full_ru.png" alt="GPU / NVIDIA GeForce RTX 4090" width="100%"></a></td>
  <td width="33%" valign="top"><a href="./assets/screenshots/showcase_nic_full_ru.png" title="NIC ITR / RSS / Intel I226-V"><img src="./assets/screenshots/showcase_nic_full_ru.png" alt="NIC ITR / RSS / Intel I226-V" width="100%"></a></td>
</tr>
</table>

<sub>Скриншоты из тест-панели, устройства и значения тестовые.</sub>

## <samp>[ ФУНКЦИОНАЛЬНОСТЬ ]</samp>

### <samp>Устройства / CPU</samp>

- Автоматически определяет GPU / USB-контроллеры / сетевые адаптеры / аудиоконтроллеры / контроллеры накопителей.
- Автоматически определяет, на каком USB-контроллере находятся мышь / клавиатура / геймпад / микрофон / аудио / камера. Если несколько устройств подключены к одному контроллеру, отображаются все найденные роли.
- Автоматически определяет и отображает P-Core / E-Core, SMT / HT, AMD CCD / CCX и доступный CPPC-рейтинг. Показывает, к какому физическому ядру относится каждый логический поток.

### <samp>Интеллектуальная авто-оптимизация</samp>

- Поддерживает интеллектуальную авто-оптимизацию с движком планирования, который учитывает P-Core / E-Core, SMT / HT, CPPC-рейтинг, CCD / CCX, количество устройств и их роли. Для дискретной GPU подбираются два разных физических ядра, для USB / сети / аудио используются свои правила распределения.
- Поддерживает резервирование выбранных CPU через ReservedCpuSets.

### <samp>MSI / IRQ / Interrupt Affinity</samp>

- Поддерживает настройку MSI Mode / MSI Limit / IRQ Priority / IRQ Policy / Interrupt Affinity.
- Отдельно отображает настройку MSI Mode из реестра и текущий режим прерываний в IRQ Mode. Поддерживаемые устройством режимы отображаются в HW.

  При открытии использует уже загруженный DTIMOD.sys или ресурсы Windows, без загрузки драйвера. Если по ресурсам Windows нельзя отличить MSI от MSI-X, показывает MSI/MSI-X.

  **СЧИТАТЬ IRQ** вверху окна читает режимы всех отображаемых PCI-устройств через DTIMOD.sys; при необходимости загружает драйвер.

### <samp>USB IMOD / NIC ITR / RSS</samp>

- Считывает и отображает IMOD у поддерживаемых USB-контроллеров xHCI. Режим **Устройства** задаёт профиль по USB-ролям, **XHCI** использует одно значение для контроллера, **Interrupters** позволяет задать значения по индексам.

  Для чтения и записи использует самописный драйвер DTIMOD.sys, после записи повторно считывает значения. Поддерживает сохранение изменённых значений для следующего входа в Windows.
- Считывает и отображает NIC ITR у поддерживаемых сетевых адаптеров Intel / Realtek, включая совместимые Killer. Информация по очередям и RX/TX зависит от семейства контроллера. Позволяет изменить текущие значения и сохранить профиль для автоматического применения при следующем входе в Windows.
- Поддерживает настройку RSS у проводных сетевых адаптеров NDIS: базовый CPU, количество очередей и режим RSS/IRQ.

### <samp>Polling Rate / энергосбережение</samp>

- Отображает информацию о Polling Rate мыши / клавиатуры и поддерживает замер частоты событий Raw Input. У клавиатуры частота событий зависит от нажатий.
- Поддерживает настройку RawMouseThrottleDuration на совместимых сборках Windows 11 с отображением текущего значения в интерфейсе.
- Поддерживает настройку энергосбережения USB-контроллеров / корневых USB-хабов / проводных сетевых адаптеров.

### <samp>Интерфейс / навигация</samp>

- Поддерживает поиск по устройствам, переход по категориям и переключение RU/EN без перезапуска.

### <samp>Бекапы / логи</samp>

- Поддерживает создание бекапов, восстановление выбранного бекапа, удаление старых бекапов и сброс настроек.
- Записывает подробные логи определения устройств и применения настроек (что применилось / что пропущено / по какой причине).

## <samp>[ СКАЧАТЬ ]</samp>

В [релизе](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4) доступны две сборки:

- **DEVICE.TWEAKER.exe**. Для запуска требуется [.NET 8 Desktop Runtime x64](https://dotnet.microsoft.com/download/dotnet/8.0).
- **DEVICE.TWEAKER.SelfContained.exe** включает .NET, отдельно устанавливать его не нужно.

Драйвер DTIMOD.sys и загрузчик KDU встроены в обе сборки. SHA-256 каждого EXE указан в SHA256SUMS.txt.

## <samp>[ ИСПОЛЬЗОВАНИЕ ]</samp>

Скачайте нужную сборку и запустите программу от имени администратора.

Для интеллектуальной авто-оптимизации нажмите **АВТООПТИМИЗАЦИЯ**. Если есть подходящие USB-контроллеры xHCI, программа предложит включить настройку IMOD перед применением. NIC ITR настраивается вручную.

Для ручной настройки выберите нужные параметры и нажмите **ПРИМЕНИТЬ**. Эта кнопка применяет настройки прерываний, ReservedCpuSets, RSS, Mouse Throttle, энергосбережения и USB IMOD.

У IMOD / NIC ITR также есть отдельные кнопки:

- IMOD: **ЗАДАТЬ** применяет настройки и сохраняет значения, отличающиеся от стандартных, для следующего входа в Windows.
- NIC ITR: **ЗАДАТЬ** меняет текущие значения, **СОХРАНИТЬ** сохраняет профиль для следующего входа в Windows.
- **ПРОВЕРИТЬ** считывает текущие значения IMOD / NIC ITR.
- IMOD: **УДАЛИТЬ** убирает общий автозапуск IMOD / NIC ITR и возвращает поля IMOD к стандартному значению. Чтобы завершить сброс, перезагрузите компьютер.

**ОБНОВИТЬ** повторно определяет устройства. Для отката настроек нажмите **ВОССТАНОВЛЕНИЕ** и выберите нужный бекап.

После **ПРИМЕНИТЬ** / **АВТООПТИМИЗАЦИЯ** перезагрузите компьютер, чтобы завершить применение настроек.

Перед ручным применением и сбросом программа создаёт бекап. При интеллектуальной авто-оптимизации сохраняются исходные настройки.

Подробное описание настроек: [SETTINGS.md](./DEVICE%20TWEAKER/docs/SETTINGS.md).

## <samp>[ БАГИ / ТЕСТИРОВАНИЕ ]</samp>

Сообщайте о найденных багах / недоработках в [issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml). Прикладывайте лог, версию программы, сборку Windows и конфигурацию ПК. Также напишите, что делали и что произошло.

Если хотите помочь с тестированием интеллектуальной авто-оптимизации, отправляйте логи в следующем порядке:

1. До применения интеллектуальной авто-оптимизации (или после ресета настроек).
2. После применения интеллектуальной авто-оптимизации.
3. После перезагрузки компьютера и повторного запуска программы.

Обновления и обсуждение проекта: [Telegram](https://t.me/arsenzaa).

## <samp>[ ИСХОДНИКИ ]</samp>

[Сборка и проверки](./DEVICE%20TWEAKER/README.md) · [История изменений](./CHANGELOG.md) · [GPL-3.0](./LICENSE)
