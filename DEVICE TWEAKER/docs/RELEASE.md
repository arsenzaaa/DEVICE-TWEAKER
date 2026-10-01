# Релиз DEVICE TWEAKER

Порядок подготовки следующего релиза на GitHub. `v0.0.3` приведён как пример.

## Структура репозитория

На GitHub проект лежит во вложенной папке. Локальная папка на рабочем столе соответствует `DEVICE TWEAKER/` внутри репозитория.

```text
DEVICE-TWEAKER/
  LICENSE
  README.md
  DEVICE TWEAKER/
    ...исходники проекта...
```

Корневой `README.md` поддерживается в самом репозитории; английская версия — `README.en.md`.

## Локальная сборка пакета

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Flavor both -Configuration Release -SkipImodDriverBuild
```

Готовый набор будет здесь:

`bin\ReleasePackages\v0.0.3\`

На GitHub Release прикрепляются:

- `DEVICE.TWEAKER.exe`
- `DEVICE.TWEAKER.SelfContained.exe`
- `SHA256SUMS.txt`

Драйвер `DTIMOD.sys` и загрузчик KDU уже встроены в EXE. Отдельные копии в assets не нужны.

## Публикация

1. Сверить изменения с исходным кодом и обновить `CHANGELOG.md`, `CHANGELOG.en.md` и `docs/releases/v*.md`.
2. Собрать пакет и проверить два EXE по `SHA256SUMS.txt`.
3. Загрузить актуальные исходники, README и заметки к релизу на `main`.
4. Создать или обновить релиз с нужным тегом и текстом из `docs/releases/v*.md` (без заголовка и навигационных ссылок).
5. Прикрепить два EXE и `SHA256SUMS.txt` из соответствующей папки `bin\ReleasePackages\`.

## Проверка после публикации

1. Сверить опубликованные SHA-256 обоих EXE с локальным `SHA256SUMS.txt`.
2. Проверить ссылки README, текст релиза и состав assets на GitHub.
3. Запустить автономную версию от имени администратора и проверить основные сценарии.
