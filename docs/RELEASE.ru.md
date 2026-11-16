# Выпуск DEVICE TWEAKER

[Документация](./README.ru.md) / Выпуск · [English](./RELEASE.md)

Соберите одну проверенную ревизию, примите её конкретные файлы и опубликуйте новый выпуск. Workflow комплекта только собирает и проверяет; прав публикации у него нет.

## Исходники и описание выпуска

Используйте чистую изолированную рабочую копию. До сборки согласуйте версию продукта, assembly/file versions, manifest приложения, ссылки и скриншоты. Версии внутренних протоколов драйвера и профилей меняйте только при изменении самих форматов.

Пишите changelog по отличиям реализации от предыдущего опубликованного выпуска. Заголовки коммитов помогают найти изменения, но сами по себе их не подтверждают. Первый выпуск задаёт исходную точку и описывает включённую реализацию. Согласуйте EN/RU, сохраняйте исторические даты и укажите окончательный SHA исходников в описании выпуска.

## Локальный комплект

Выполняйте из корня репозитория в чистой рабочей копии:

```powershell
[xml]$project = Get-Content .\src\DeviceTweakerCS.csproj
$version = [string]($project.Project.PropertyGroup.InformationalVersion | Where-Object { $_ } | Select-Object -First 1)
$package = ".\src\bin\ReleasePackages\v$version"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\src\build.ps1 -Flavor both -Configuration Release -SkipImodDriverBuild
if ($LASTEXITCODE -ne 0) { throw 'Build failed' }
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Test-ReleasePackage.ps1 -PackageDirectory $package
if ($LASTEXITCODE -ne 0) { throw 'Package check failed' }

$sourceRoot = '.\local\release-sources'
New-Item -ItemType Directory -Force -Path $sourceRoot | Out-Null
$sourceCommit = (git rev-parse HEAD).Trim()
if ($LASTEXITCODE -ne 0) { throw 'Source revision lookup failed' }
$sourceName = "DEVICE.TWEAKER.Source.$sourceCommit.zip"
git archive --format=zip --output="$sourceRoot/$sourceName" HEAD
if ($LASTEXITCODE -ne 0) { throw 'Source archive failed' }
Get-FileHash -LiteralPath "$sourceRoot/$sourceName" -Algorithm SHA256
```

Публикуемый комплект содержит `DEVICE.TWEAKER.exe`, `DEVICE.TWEAKER.SelfContained.exe` и `SHA256SUMS.txt`. Исходники доступны через Git и стандартные архивы GitHub для тега. Точный локальный source ZIP и его сумму сохраняйте отдельно от публикуемых файлов.

Обычный EXE требует .NET 10 Desktop Runtime x64. Self-contained EXE включает его. [Проверка комплекта](../dev/qa/Test-ReleasePackage.ps1) сравнивает версии файлов с исходным проектом, проверяет общую идентичность приложения, встроенные компоненты и отсутствие тест-панели через read-only запуск probe. Она не загружает DTIMOD и не подтверждает аппаратное поведение. Сверьте содержимое архива исходников с выбранным Git tree.

Завершите проверки изменённого поведения, просмотрите снимки EN/RU и проведите приёмку окончательных EXE. Зафиксируйте проверенные Windows/устройства/драйверы и непокрытые конфигурации. Загрузка драйверов, запись регистров и изменение настроек Windows требуют отдельного разрешения на тестовой системе.

## Сборка без публикации

Запустите `Release package` с точным 40-символьным `source_sha`. Workflow восстанавливает зависимости, выполняет тесты, собирает оба EXE, проверяет комплект и сохраняет его как Actions artifact. Версия и каталог берутся из выбранного исходного проекта. Теги и выпуски не создаются и не заменяются.

Скачайте artifact, проверьте состав и запишите workflow run, SHA исходников и суммы файлов. Приёмка и публикация должны использовать именно эти файлы. Пересборка после приёмки создаёт другой кандидат.

## Публикация нового выпуска

После разрешения публикации:

1. Убедитесь, что выбран проверенный коммит, CI завершился успешно, а скачанный комплект имеет принятые суммы.
2. Сверьте удалённые refs и выпуски. Если выбранный тег или версия уже существуют, остановитесь; не перезаписывайте их неявно.
3. Создайте новый тег на проверенном SHA и опубликуйте принятые файлы с вычитанным описанием. Укажите SHA исходников и требования Windows/runtime.
4. Скачайте публичные файлы и сравните суммы. Проверьте назначение тега, архивы исходников, обе версии README, ссылки галереи и поддержки.

Изменение прежних тегов, страниц выпусков или публичной истории требует отдельного решения о переходе. До такого перехода сохраните все прежние refs, объекты аннотированных тегов, описания выпусков и файлы локально; проверьте восстановление из архива. Личная история и архивы восстановления не должны попадать в публичный выпуск или Actions artifacts. Обычный новый выпуск не требует переписывать предыдущие.

## Страница проекта

README использует обычные Markdown-заголовки, текстовую навигацию и один снимок интерфейса. Дополнительные сценарии размещаются в [галерее](./SCREENSHOTS.ru.md). Снимки получаются из read-only сборки на скрытом desktop:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\dev\qa\Invoke-DevSession.ps1 -Scenario .\dev\qa\Scenarios\gallery-intel-gpu.json -OutputRoot .\local\screenshots
```

Сценарии `gallery-*` охватывают affinity Intel/AMD, роли и interrupters USB, сеть, восстановление, отчёты операций и проверку системы. Каждый сохраняет снимки EN/RU. После просмотра выберите `images/en.png` и `images/ru.png` для соответствующих файлов галереи; сессию и ссылку на сборку оставьте локально. `Capture-Showcase.ps1 -ReadmeOnly` создаёт снимки с прежними именами и не пересоздаёт эту галерею.

Просмотрите оба языка, подпишите демонстрационные данные и проверьте согласованность контролов, сводок и readback. Используйте исходные снимки без декоративных рамок. Описание About должно соответствовать английской главной странице.
