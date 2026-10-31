# Разработка DEVICE TWEAKER

[Главная](../README.md) · [Описание настроек](./SETTINGS.md) · [English](./DEVELOPMENT.en.md)

## Требования

- Windows 10 / 11 x64, [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0), Windows PowerShell 5.1.
- Для пересборки `DTIMOD.sys` и нативной модели в QA: Visual Studio с C++ Build Tools, Windows SDK и Windows Driver Kit. Готовый драйвер уже лежит в репозитории, для обычной сборки они не нужны.

## Сборка и тесты

Из корня репозитория:

```powershell
dotnet build .\src\DeviceTweakerCS.csproj -c Release -p:TreatWarningsAsErrors=true
dotnet test .\dev\qa\tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj -c Release
```

Предупреждения считаются ошибками и в CI. `src/.editorconfig` включает правила CA1304 / CA1305 / CA1310: форматирование и сравнения строк не должны зависеть от языка Windows.

## Релизные EXE

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\src\build.ps1 -Flavor both -Configuration Release -SkipImodDriverBuild
```

| Параметр | Что делает |
| --- | --- |
| `-Flavor both` / `with-net` / `without-net` | обе сборки, только со встроенным .NET или только обычная |
| `-SkipImodDriverBuild` | использовать готовый `IMOD/DTIMOD.sys` |
| `-TrustImodDriverCert` | установить сертификат драйвера на тестовом компьютере |
| `-NoClean` | не очищать промежуточные файлы |

Результат:

- `src/bin/Publish/DEVICE TWEAKER/DEVICE TWEAKER.exe` — обычная сборка, нужен .NET 8 Desktop Runtime x64;
- `src/bin/Publish/DEVICE TWEAKER (SELF-CONTAINED)/…` — со встроенным .NET, сжатый single-file;
- `src/bin/ReleasePackages/v<версия>/` — два EXE и `SHA256SUMS.txt` для релиза.

Папки `Backups` и `logs` рядом с EXE при чистой сборке сохраняются.

## Драйвер DTIMOD

Исходники — `src/IMOD/Driver/imod_driver.c` и `.h`. После их изменения:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\src\publish-variants.ps1 -BuildDriverOnly
```

Команда собирает и подписывает `IMOD/DTIMOD.sys` и обновляет `DTIMOD.sys.sha256` и `DTIMOD.build.json` с хешами исходников. Если в MSBuild нет задач WDK, используется прямой вызов `cl` / `link`. QA останавливается, пока бинарник не совпадает с исходниками. Скрипты тестового сертификата лежат в `src/Scripts`; приватные ключи, `.pfx` и `.p12` в репозиторий не добавляются.

Загрузчик KDU закреплён на версии 1.4.5.2512, его исходники и манифесты — в [`src/IMOD/KDU`](../src/IMOD/KDU/README.md).

## Автозапуск IMOD / NIC ITR

`src/Startup/ApplyIMOD.ps1` и `src/Startup/DeviceTweakerImodRuntime.cs` встраиваются в EXE и при сохранении профиля копируются в `%ProgramData%\DEVICE TWEAKER\Startup`. Runtime общий для программы и скрипта и компилируется ещё и в Windows PowerShell 5.1, поэтому в нём допустим только синтаксис C# 5.

## Проверки

Полный локальный QA — сборки, тесты, намеренные поломки, нативная модель драйвера и проверки интерфейса на скрытом рабочем столе — описан в [`dev/qa`](../dev/qa/README.md). Он работает только на чтение и не загружает драйвер.

## Сайт и графика

Сайт [arsenzaaa.github.io/DEVICE-TWEAKER](https://arsenzaaa.github.io/DEVICE-TWEAKER/) собирается из `docs/`, `CHANGELOG.md` / `CHANGELOG.en.md` и текста главной страницы в `dev/site/build.py`. Workflow `Site` публикует его на GitHub Pages после изменений в `main`. Локально:

```powershell
pip install markdown==3.7
python dev/site/build.py --out _site
```

Знак и схемы в `assets/branding` рисует `dev/branding/render.py`: текст переводится в кривые, поэтому SVG выглядят одинаково в любой системе. Рамку и тень скриншотов добавляет `dev/branding/screens.py`, исходные снимки делает `dev/qa/Capture-Showcase.ps1`.

## Устройство кода

| Папка | Содержимое |
| --- | --- |
| `src/Affinity` | топология CPU, маски affinity, ReservedCpuSets |
| `src/Devices` | обнаружение и классификация устройств, RSS, энергосбережение |
| `src/Core` | применение и восстановление настроек, IMOD, NIC ITR, автозапуск, логи |
| `src/GUI`, `src/Localization` | интерфейс и строки RU / EN |
| `src/Interop` | вызовы Windows API |
| `src/Startup` | скрипт автозапуска и общий runtime регистров |
| `src/IMOD` | драйвер DTIMOD и загрузчик KDU |
| `dev/qa` | тесты, тест-панель и QA |
| `dev/site`, `dev/branding` | сайт, знак, схемы и рамка скриншотов |

Логи приложения пишутся в папку `logs` рядом с EXE, отдельный файл на каждый запуск. При необработанной ошибке там же появляется crash-файл и `last-crash.txt`.

Выпуск версии описан в [RELEASE.md](./RELEASE.md), правила участия — в [CONTRIBUTING](../.github/CONTRIBUTING.md).
