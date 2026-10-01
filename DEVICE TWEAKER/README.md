# Сборка DEVICE TWEAKER

[Главная страница](../README.md) | [English](./README.en.md) | [История изменений](../CHANGELOG.md)

## Требования

- Windows 10/11 x64
- .NET 8 SDK
- PowerShell 5.1 или новее
- Visual Studio C++ Build Tools и Windows Driver Kit для пересборки `DTIMOD.sys`

## Каталоги

- `Affinity` содержит топологию CPU, маски и настройки affinity.
- `Core` содержит применение настроек, восстановление, IMOD, NIC ITR и диагностику.
- `Devices` отвечает за обнаружение и классификацию устройств.
- `GUI` содержит Windows Forms интерфейс.
- `Interop` содержит вызовы Windows API.
- `Localization` содержит русские и английские строки.
- `IMOD` содержит драйвер, загрузчик и связанные файлы.
- `tests` содержит модульные тесты.

Обычное обновление списка устройств не загружает драйвер Ring 0. Доступ к MMIO выполняется только при явной проверке или записи IMOD/ITR и только для распознанного профиля оборудования.

## Сборка

Команды выполняются из корня репозитория.

```powershell
dotnet restore ".\DEVICE TWEAKER\DeviceTweakerCS.csproj"
dotnet build ".\DEVICE TWEAKER\DeviceTweakerCS.csproj" -c Release -p:TreatWarningsAsErrors=true
dotnet test ".\DEVICE TWEAKER\tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj" -c Release
```

## Релизные файлы

```powershell
powershell -ExecutionPolicy Bypass -File ".\DEVICE TWEAKER\publish-variants.ps1" `
  -Flavor both -Configuration Release -SkipImodDriverBuild
```

Скрипт создаёт компактную и автономную сборки, а также файл `SHA256SUMS.txt`.

## Драйвер

В репозитории находится тестово подписанный `IMOD/DTIMOD.sys`. Ожидаемый хеш хранится в `IMOD/DTIMOD.sys.sha256` и проверяется программой перед использованием.

Не добавляйте в репозиторий приватные ключи, PFX, локальные сертификаты и случайные бинарные файлы. Для нового Device ID нельзя копировать MMIO-смещения похожего контроллера без проверки документации и регистров.

Правила для изменений и отчётов об ошибках находятся в [CONTRIBUTING.md](../.github/CONTRIBUTING.md) и [SECURITY.md](../.github/SECURITY.md).
