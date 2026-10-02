# Исходники DEVICE TWEAKER

[Описание программы](../README.md) · [Как работают настройки](./docs/SETTINGS.md) · [English](./README.en.md)

Здесь лежат исходники Windows Forms приложения. Если вам нужен готовый EXE, скачайте его со [страницы релиза](https://github.com/arsenzaaa/DEVICE-TWEAKER/releases/tag/v0.0.4).

## Собрать и проверить

Нужны Windows 10/11 x64, .NET 8 SDK и PowerShell 5.1 или новее. Запускайте команды из корня репозитория:

```powershell
dotnet restore ".\DEVICE TWEAKER\DeviceTweakerCS.csproj"
dotnet build ".\DEVICE TWEAKER\DeviceTweakerCS.csproj" -c Release -p:TreatWarningsAsErrors=true
dotnet test ".\DEVICE TWEAKER\tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj" -c Release
```

Готовый `IMOD/DTIMOD.sys` уже есть в репозитории. Visual Studio C++ Build Tools и Windows Driver Kit нужны, только если вы хотите пересобрать драйвер.

Два EXE и `SHA256SUMS.txt` собираются так:

```powershell
powershell -ExecutionPolicy Bypass -File ".\DEVICE TWEAKER\publish-variants.ps1" `
  -Flavor both -Configuration Release -SkipImodDriverBuild
```

Подробные параметры сборки и порядок выпуска описаны в [BUILD.md](./docs/BUILD.md) и [RELEASE.md](./docs/RELEASE.md).

## Где искать код

- `Affinity` — топология CPU и маски.
- `Devices` — обнаружение и классификация устройств.
- `Core` — применение настроек, восстановление, IMOD, NIC ITR и диагностика.
- `GUI` и `Localization` — интерфейс и строки RU/EN.
- `Interop` — вызовы Windows API.
- `IMOD` — драйвер и загрузчик.
- `tests` — тесты.

Если нашли ошибку или хотите прислать изменение, начните с [CONTRIBUTING.md](../.github/CONTRIBUTING.md).
