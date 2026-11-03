# Разработка DEVICE TWEAKER

[Документация](./README.ru.md) / Разработка · [English](./DEVELOPMENT.md)

## Требования

- Windows 10 / 11 x64, [.NET 10 SDK](https://dotnet.microsoft.com/download/dotnet/10.0), Windows PowerShell 5.1.
- Для пересборки `DTIMOD.sys` и нативной модели в QA: Visual Studio с C++ Build Tools, Windows SDK и Windows Driver Kit. Готовый драйвер уже лежит в репозитории, для обычной сборки они не нужны.

`global.json` выбирает стабильный SDK 10.0 и допускает более новые feature bands внутри этой версии. Общий `src/Startup/DeviceTweakerImodRuntime.cs` также компилируется Windows PowerShell 5.1: в нём сохраняется синтаксис C# 5.

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

- `-Flavor both` / `with-net` / `without-net` - обе сборки, только со встроенным .NET или только обычная
- `-SkipImodDriverBuild` - использовать готовый `IMOD/DTIMOD.sys`
- `-TrustImodDriverCert` - установить сертификат драйвера на тестовом компьютере
- `-NoClean` - не очищать промежуточные файлы

Результат:

- `src/bin/Publish/DEVICE TWEAKER/DEVICE TWEAKER.exe` - обычная сборка, нужен .NET 10 Desktop Runtime x64;
- `src/bin/Publish/DEVICE TWEAKER (SELF-CONTAINED)/…` - со встроенным .NET, сжатый single-file;
- `src/bin/ReleasePackages/v<версия>/` - два EXE и `SHA256SUMS.txt` для релиза.

Папки `Backups` и `logs` рядом с EXE при чистой сборке сохраняются.

## Драйвер DTIMOD

Исходники - `src/IMOD/Driver/imod_driver.c` и `.h`. Сборка/подпись драйвера отделена от обычной сборки приложения и выполняется в изолированной среде разработки:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\src\publish-variants.ps1 -BuildDriverOnly
```

Команда собирает и подписывает `IMOD/DTIMOD.sys` и обновляет `DTIMOD.sys.sha256` и `DTIMOD.build.json` с хешами исходников. Если в MSBuild нет задач WDK, используется прямой вызов `cl` / `link`. QA останавливается, пока бинарник не совпадает с исходниками. Скрипты тестового сертификата лежат в `src/Scripts`; приватные ключи, `.pfx` и `.p12` в репозиторий не добавляются.

WDK signing targets могут создавать тестовые сертификаты; `-TrustImodDriverCert` явно меняет доверенные хранилища системы. Не используйте эту процедуру как обычную проверку GUI и не загружайте драйвер ради компиляции.

Загрузчик KDU закреплён на версии 1.4.5.2512, его исходники и манифесты - в [`src/IMOD/KDU`](../src/IMOD/KDU/README.md).

## Автозапуск IMOD / NIC ITR

`src/Startup/ApplyIMOD.ps1` и `src/Startup/DeviceTweakerImodRuntime.cs` встраиваются в EXE и при сохранении профиля копируются в `%ProgramData%\DEVICE TWEAKER\Startup`. Runtime общий для программы и скрипта и компилируется ещё и в Windows PowerShell 5.1, поэтому в нём допустим только синтаксис C# 5.

## Проверки

Полный локальный QA - сборки, тесты, намеренные поломки, нативная модель драйвера и проверки интерфейса на скрытом рабочем столе - описан в [`dev/qa`](../dev/qa/README.ru.md). Он работает только на чтение и не загружает драйвер.

## Выбор проверок

Для изменения кода запускайте сборку и проверки затронутого поведения. Полный QA нужен при интеграции сквозных изменений; его успешный дорогой прогон не требуется повторять после каждой правки текста.

Для GUI сначала воспроизведите сценарий, затем выполните его после изменения:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Invoke-DevSession.ps1 -Scenario .\dev\qa\Scenarios\device-flow.json -OutputRoot .\local\dev-sessions
# Для сравнения добавьте -CompareWith <каталог предыдущего запуска>.
```

Просмотрите PNG, различия и переходы. После просмотра заполните review-input.json и используйте `dev/qa/Review-DevSession.ps1`. MATCH подтверждает совпадение пикселей, а не одобрение дизайна. Протокол - в [WORKBENCH.md](../dev/qa/WORKBENCH.ru.md).

Стенд работает на скрытом отдельном desktop и в неизменяемом `--qa-read-only`. Он не должен трогать пользовательские процессы или активный рабочий стол. Модели аппаратуры и раскладки не заменяют реальный driver/readback, сон, перезагрузку или DPI Windows. Ограничения размеров окна сохраняются в отчёте.

## RU/EN и состояние UI

Нейтральные английские строки находятся в `src/Localization/Strings.resx`, русские - в `Strings.ru.resx`. Стабильные ключи основных действий объявлены в UiText, доступ к ресурсам - через UiResources. Идентификаторы контролов независимы от переведённых подписей.

Существующий UiLanguage сохраняет адаптер для динамических диагностических фраз. При добавлении текста обновляйте обе культуры и проверяйте параметры форматирования. Изменение языка или масштаба не должно заново читать устройства, терять pending/baseline, фокус или положение прокрутки.

## Скриншоты

Снимайте разработческую read-only сборку через `dev/qa/Capture-Showcase.ps1`. В [галерее](./SCREENSHOTS.ru.md) используются полные окна приложения без добавленной рамки и тени. Подписывайте демонстрационные конфигурации и согласовывайте значения контролов со сводкой. Команда захвата приведена в [руководстве по выпуску](./RELEASE.ru.md#страница-проекта).

## Автоматизация выпуска

По умолчанию workflow только готовит файлы. Замена использует ранее принятый комплект и проверенный архив старого выпуска. Проверки отказов выполняются локально без доступа к GitHub:

```powershell
python -m unittest discover -s dev/release -p test_replace_release.py -v
```

Python 3.11+ нужен только для этой автоматизации. Подробности - в [руководстве по выпуску](./RELEASE.ru.md).

## Устройство кода

- `src/Affinity` - топология CPU, маски affinity, ReservedCpuSets
- `src/Devices` - обнаружение и классификация устройств, RSS, энергосбережение
- `src/Core` - применение и восстановление настроек, IMOD, NIC ITR, автозапуск, логи
- `src/GUI`, `src/Localization` - интерфейс и строки RU / EN
- `src/Interop` - вызовы Windows API
- `src/Startup` - скрипт автозапуска и общий runtime регистров
- `src/IMOD` - драйвер DTIMOD и загрузчик KDU
- `dev/qa` - тесты, тест-панель и QA

Логи приложения пишутся в папку `logs` рядом с EXE, отдельный файл на каждый запуск. При необработанной ошибке там же появляется crash-файл и `last-crash.txt`.

Выпуск версии описан в [RELEASE.md](./RELEASE.ru.md), правила участия - в [CONTRIBUTING](../.github/CONTRIBUTING.ru.md).
