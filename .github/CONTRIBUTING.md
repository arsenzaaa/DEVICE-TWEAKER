# Как помочь DEVICE TWEAKER

Нашли баг, устройство определяется неправильно или есть идея для новой настройки? Откройте issue. Логи и точная конфигурация железа помогут разобраться быстрее, чем описание «не работает».

## Если нашли баг

Укажите версию DEVICE TWEAKER, сборку Windows, CPU, материнскую плату, BIOS, Hardware ID устройства и версию драйвера. Напишите, что нажали, чего ожидали, что произошло и перезагружали ли компьютер. Приложите фрагмент лога с этим действием.

Перед отправкой уберите из лога личные пути, имена пользователей и серийные номера. Для обычной ошибки не нужны дампы всей системы или экспорт реестра.

## Если хотите отправить PR

Делайте один PR на одну задачу. Напишите, на каком железе или сборке Windows проверили изменение и как повторить проверку. Если заявляете об уменьшении задержки, приложите методику и исходные результаты измерений.

Перед PR запустите сборку и тесты по [руководству](../DEVICE%20TWEAKER/README.md). Для изменений записи в регистры, IMOD/ITR и восстановления проверьте также случай, когда запись или чтение не удались. Для интерфейса проверьте русский и английский языки и масштаб Windows 100%, 125% и 150%.

Код проекта распространяется по [GPL-3.0](../LICENSE).

---

# Contributing

Bug reports, hardware compatibility results, documentation fixes, and focused pull requests are welcome.

For a bug, include the DEVICE TWEAKER version, Windows build, CPU, motherboard, BIOS, device Hardware ID, driver version, steps to reproduce, expected and actual results, and a relevant log excerpt. Mention whether you restarted Windows. Remove personal paths, usernames, and serial numbers from logs before posting.

Keep each PR focused. Explain which hardware or Windows build you tested and how to repeat the check. Claims about lower latency need a repeatable method and raw results. Run the build and tests from the [build guide](../DEVICE%20TWEAKER/README.en.md) before submitting. Check read/write failures for changes to registers, IMOD/ITR, and restore; check RU/EN and 100%, 125%, and 150% scaling for UI changes.

Contributions are distributed under [GPL-3.0](../LICENSE).
