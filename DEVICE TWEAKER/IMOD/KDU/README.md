# Исходный код KDU

В этой папке находятся только компоненты KDU, необходимые для пересборки загрузчиков DEVICE TWEAKER.

Файлы, используемые программой.

- `IMOD/Loader/kdu.exe`
- `IMOD/Loader/drv64.dll`

Соответствие исходного кода.

- `Source/Hamakaze/KDU.vcxproj` собирает `kdu.exe`.
- `Source/Tanikaze/Tanikaze.vcxproj` собирает `drv64.dll`.
- `Source/Utils/GenAsIo2Unlock/GenAsIo2Unlock.vcxproj` собирает вспомогательный файл для post-build этапа Hamakaze.
- `Source/Shared` содержит общие заголовки и исходный код KDU.

Оригинальный проект находится в репозитории [hfiref0x/KDU](https://github.com/hfiref0x/KDU).

KDU распространяется по лицензии MIT. Текст лицензии находится в [LICENSE.txt](LICENSE.txt).

Файлы в `IMOD/Loader`.

- `kdu.exe`, версия `1.4.5.2512`, SHA-256 `B340DAD4DDBE8607F9FDDB79F679375B4FF5080FE1A7EDB6CE015F69D3A0CD4F`.
- `drv64.dll`, версия `1.4.5.2512`, SHA-256 `D032D855A0FF0EF3E3AD6EC8DAFFB8649048F82D3F1EEB89BEE3652EE6F01F80`.

## Сборка

Для сборки требуются Visual Studio 2019 или новее с C++ Build Tools, а также компоненты Windows SDK и WDK, необходимые KDU.

Скрипт по умолчанию заменяет upstream `PlatformToolset` на `v143`. Другой toolset можно передать через `-PlatformToolset`.

Команда запуска из корневой папки DEVICE TWEAKER.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\IMOD\KDU\build-kdu.ps1
```

Результаты сборки.

- `Source/Utils/GenAsIo2Unlock/output/x64/Release/GenAsIo2Unlock.exe`
- `Source/Tanikaze/output/x64/Release/drv64.dll`
- `Source/Hamakaze/output/x64/Release/kdu.exe`

Команда для копирования собранных файлов в `IMOD/Loader`.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\IMOD\KDU\build-kdu.ps1 -CopyToLoader
```

Не заменяйте файлы из `IMOD/Loader` без проверки загрузки драйвера на чистой системе.
