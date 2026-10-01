# Contributing / Участие в разработке

Bug reports, hardware compatibility results, documentation corrections, and focused pull requests are welcome.

## Before opening an issue

- Use the latest preview release and search existing issues.
- Include the Windows build, CPU, motherboard, BIOS, device Hardware ID, and driver version.
- Describe the exact action, expected result, actual result, and whether a restart was performed.
- Attach the relevant log after removing usernames, paths, serial numbers, and other private data.
- Never upload registry exports, certificates, keys, or unrelated system dumps.

## Pull requests

Keep each pull request focused. Explain the hardware or Windows behavior behind the change and how it was verified. Do not claim a latency improvement without a repeatable measurement method and raw results.

Before submitting:

```powershell
dotnet build ".\DEVICE TWEAKER\DeviceTweakerCS.csproj" -c Release -p:TreatWarningsAsErrors=true
dotnet test ".\DEVICE TWEAKER\tests\DeviceTweaker.Tests\DeviceTweaker.Tests.csproj" -c Release
```

Changes to MMIO profiles, IOCTL handling, backups, restore, or registry writes require negative-path tests. UI changes require RU and EN checks and should preserve usable layout at 100%, 125%, and 150% scaling.

By contributing, you agree that your work is distributed under GPL-3.0.

---

Перед issue укажите сборку Windows, железо, Hardware ID и версию драйвера, точные шаги и факт перезагрузки. Очистите логи от личных данных. PR должен решать одну задачу, проходить сборку и тесты; изменения MMIO, IOCTL, backup/restore и реестра требуют проверки ошибочных сценариев. Заявления об улучшении задержки принимаются только с воспроизводимой методикой и исходными результатами.
