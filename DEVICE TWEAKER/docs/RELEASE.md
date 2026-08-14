# Релиз DEVICE TWEAKER

Краткая инструкция для публикации `v0.0.4-alpha.2` на GitHub.

## Структура репозитория

На GitHub проект лежит во вложенной папке. Локальная папка на рабочем столе соответствует `DEVICE TWEAKER/` внутри репозитория.

```text
DEVICE-TWEAKER/
  LICENSE
  README.md
  DEVICE TWEAKER/
    ...исходники проекта...
```

Короткий корневой README берется из `docs/github-root-README.md`.

## Локальная сборка пакета

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Flavor both -Configuration Release -SkipImodDriverBuild
```

Готовый набор будет здесь:

`bin\ReleasePackages\v0.0.4-alpha.2\`

На GitHub Release (как в alpha.1) прикрепляются **только**:

- `DEVICE.TWEAKER.exe`
- `DEVICE.TWEAKER.NET.FRAMEWORK.exe`

Остальные файлы пакета (`DTIMOD.sys`, notes, `SHA256SUMS.txt` и т.п.) в assets релиза не загружать — драйвер уже встроен в EXE.

## Публикация

1. Загрузить исходники локальной папки проекта в `DEVICE TWEAKER/` на `main`.
2. Обновить корневой `README.md` из `docs/github-root-README.md`.
3. Создать pre-release с тегом `v0.0.4-alpha.2`.
4. Прикрепить **только два EXE** из `bin\ReleasePackages\v0.0.4-alpha.2\`.
5. В описание релиза вставить текст из `RELEASE_NOTES_v0.0.4-alpha.2.md` (стиль как у alpha.1: обычные пункты для пользователя).
6. При необходимости ответить на [issue #1](https://github.com/arsenzaaa/DEVICE-TWEAKER/issues/1) и попросить проверить новую alpha.2.

## Проверка после публикации

1. Скачать оба EXE с GitHub Releases.
2. Запустить автономную версию от имени администратора.
3. Убедиться, что без `CHECK` драйвер не загружается.
