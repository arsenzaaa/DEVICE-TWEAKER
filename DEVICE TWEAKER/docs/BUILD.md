# Сборка DEVICE TWEAKER

## Требования

- Windows 10/11 x64.
- [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0).
- Windows PowerShell 5.1 или новее.
- Для пересборки `DTIMOD.sys` необходимы Visual Studio с C++ Build Tools, MSBuild и Windows Driver Kit.

## Сборка обеих версий

Запустите команду из корневой папки проекта.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Flavor both -Configuration Release -SkipImodDriverBuild
```

Доступные параметры.

- `-Flavor both` - собрать обе версии.
- `-Flavor with-net` - собрать автономную версию со встроенным .NET 8 Runtime.
- `-Flavor without-net` - собрать обычную версию, для которой требуется .NET Desktop Runtime x64 версии 8 или новее (в сборке включён `RollForward=Major`).
- `-Configuration Release` - релизная сборка.
- `-SkipImodDriverBuild` - использовать готовый `IMOD/DTIMOD.sys`.
- `-TrustImodDriverCert` - установить сертификат драйвера на тестовом компьютере.
- `-NoClean` - не очищать промежуточные файлы перед сборкой.

После сборки файлы находятся по следующим путям.

- `bin\Publish\DEVICE TWEAKER\DEVICE TWEAKER.exe` - обычная версия.
- `bin\Publish\DEVICE TWEAKER (SELF-CONTAINED)\DEVICE TWEAKER (SELF-CONTAINED).exe` - автономная версия со встроенным .NET 8 Runtime.
- `bin\ReleasePackages\v<InformationalVersion>\` - готовый набор для GitHub Releases (два EXE и `SHA256SUMS.txt`); номер берётся из `DeviceTweakerCS.csproj`.

## Обычная сборка через dotnet

```powershell
dotnet build .\DeviceTweakerCS.csproj -c Release -p:BuildImodDriver=false
```

Параметр `BuildImodDriver=false` отключает отдельную пересборку драйвера. Без него проект потребует Visual Studio с C++ Build Tools и WDK.

Для публичного релиза используйте `build.ps1`, чтобы обе версии собирались одинаково.

## Сборка и подпись драйвера

В папке `Scripts` находятся вспомогательные скрипты.

- `Create-DevCodeSignCert.ps1` - создание тестового сертификата.
- `Install-CodeSignCert.ps1` - установка сертификата.
- `CreateCertAndBuildSigned.ps1` - создание сертификата и сборка подписанного драйвера.

Не добавляйте в репозиторий приватные ключи, `.pfx`, `.p12` и другие файлы сертификатов с закрытым ключом.

## Проверка перед релизом

1. Обновите `Version`, `AssemblyVersion`, `FileVersion` и `InformationalVersion` в `DeviceTweakerCS.csproj`, а также версию в `assets/app.manifest`.
2. Обновите описание изменений в `CHANGELOG.md` и `docs/releases/v*.md`.
3. Соберите обе версии через `build.ps1`.
4. Проверьте, что SHA-256 файла `IMOD/DTIMOD.sys` совпадает с `IMOD/DTIMOD.sys.sha256`.
5. Запустите автономную версию от имени администратора.
6. Проверьте загрузку устройств, интерфейс, подсказки и выпадающие списки.
7. Убедитесь, что без кнопки `CHECK` драйвер не загружается.
8. Добавьте в GitHub Release два EXE и `SHA256SUMS.txt` из папки сборки в `bin\ReleasePackages\`. Драйвер уже встроен в EXE; папки `logs` и `Backups` не публикуйте.
9. Перед публикацией сверьте SHA-256 обоих EXE со значениями из `SHA256SUMS.txt`. Для GitHub используйте workflow `Release package`, описанный в [RELEASE.md](RELEASE.md).

Подробный чеклист публикации находится в [RELEASE.md](RELEASE.md).

## Логи и промежуточные файлы

- `bin`, `obj`, `build`, `.vs`, `*.log`, `*.tmp`, `.pdb` и кэши сборки не должны попадать в репозиторий.
- Подробный лог приложения создается автоматически при запуске в папке `logs` рядом с EXE. Для каждого запуска используется отдельный файл `DeviceTweaker_дата_время.log`.
- При необработанной ошибке в папке `logs` создаются отдельный crash-файл и обновленный `last-crash.txt`.
- Приватные сертификаты и локальные вспомогательные файлы не должны публиковаться в GitHub Releases.
