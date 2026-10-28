<p align="right"><a href="./README.en.md">English</a></p>

<h1>
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER#gh-dark-mode-only"><picture><source media="(max-width: 600px)" srcset="./assets/branding/hero-mobile-dark.svg"><img src="./assets/branding/hero-dark.svg" alt="DEVICE TWEAKER / Windows 10/11 x64 / 0.0.4" width="1080"></picture></a>
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER#gh-light-mode-only"><picture><source media="(max-width: 600px)" srcset="./assets/branding/hero-mobile-light.svg"><img src="./assets/branding/hero-light.svg" alt="DEVICE TWEAKER / Windows 10/11 x64 / 0.0.4" width="1080"></picture></a>
</h1>

<p align="center"><strong>Настройка прерываний устройств в Windows 10/11 x64.</strong><br>
Ручная настройка и интеллектуальная авто-оптимизация.
</p>

<p align="center">
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4#gh-dark-mode-only"><img src="./assets/branding/download-ru-dark.svg" alt="Скачать 0.0.4" width="170"></a>
  <a href="https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4#gh-light-mode-only"><img src="./assets/branding/download-ru-light.svg" alt="Скачать 0.0.4" width="170"></a>
  <a href="./DEVICE%20TWEAKER/docs/SETTINGS.md#gh-dark-mode-only"><img src="./assets/branding/settings-ru-dark.svg" alt="Настройки" width="140"></a>
  <a href="./DEVICE%20TWEAKER/docs/SETTINGS.md#gh-light-mode-only"><img src="./assets/branding/settings-ru-light.svg" alt="Настройки" width="140"></a>
</p>

<p align="center">
  <a href="./SCREENSHOTS.md"><img src="./assets/screenshots/showcase_intel_14900k_full_ru.png" alt="DEVICE TWEAKER / USB IMOD / Intel Core i9-14900K" width="860"></a>
</p>

<p align="center"><a href="./SCREENSHOTS.md">USB IMOD / GPU / NIC ITR: все скриншоты</a></p>

## Интеллектуальная авто-оптимизация

Автоматически распределяет устройства по CPU с учетом P-Core / E-Core, SMT / HT, CPPC-рейтинга, AMD CCD / CCX и роли каждого устройства. Для дискретной GPU подбирает два разных физических ядра; для USB, сети и аудио используются свои правила распределения.

## Настройки

**Прерывания / CPU**

MSI Mode / MSI Limit, IRQ Priority / Policy, Interrupt Affinity, ReservedCpuSets и определение текущего IRQ Mode.

**USB / сеть**

Чтение и настройка IMOD у поддерживаемых xHCI-контроллеров, NIC ITR у поддерживаемых сетевых адаптеров. Для проводных NDIS-адаптеров доступны базовый CPU, количество очередей и режим RSS/IRQ.

Профили IMOD / NIC ITR можно сохранить для применения при следующем входе в Windows.

**Устройства / диагностика**

Определение ролей USB-устройств, отображение Polling Rate мыши / клавиатуры и замер частоты событий Raw Input. Бекапы и подробные логи применения настроек.

Также доступны Mouse Throttle на совместимых сборках Windows 11 и настройка энергосбережения USB / проводной сети.

<details>
<summary><strong>Запуск, бекапы и проверка результата</strong></summary>

1. Запустите программу от имени администратора.
2. Нажмите **АВТООПТИМИЗАЦИЯ** или выберите параметры вручную и нажмите **ПРИМЕНИТЬ**. При авто-оптимизации программа отдельно предлагает включить USB IMOD. NIC ITR настраивается вручную.
3. После применения перезагрузите компьютер. Для восстановления выбранного бекапа используйте **ВОССТАНОВЛЕНИЕ**.

Перед ручным применением создаётся бекап; интеллектуальная авто-оптимизация сохраняет исходные настройки.

Описание параметров и отдельных кнопок находится в [руководстве по настройкам](./DEVICE%20TWEAKER/docs/SETTINGS.md).

## Баги / тестирование

Сообщайте о найденных багах / недоработках в [issues](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/new?template=bug_report.yml). Прикладывайте лог, версию программы, сборку Windows и конфигурацию ПК. Напишите, что делали и что произошло.

Для проверки интеллектуальной авто-оптимизации нужны логи до применения, после применения и после перезагрузки с повторным запуском программы.

</details>

## Документация

- [История изменений](./CHANGELOG.md)
- [Сборка и инструменты разработчика](./DEVICE%20TWEAKER/README.md)

Обновления и обсуждение: [Telegram](https://t.me/arsenzaa).
