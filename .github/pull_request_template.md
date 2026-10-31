## Что изменено / What changed



## Как проверено / How it was tested

- [ ] `dotnet build src/DeviceTweakerCS.csproj -c Release -p:TreatWarningsAsErrors=true`
- [ ] `dotnet test dev/qa/tests/DeviceTweaker.Tests/DeviceTweaker.Tests.csproj -c Release`
- [ ] Железо и сборка Windows / hardware and Windows build:
- [ ] Для записи в регистры, IMOD/ITR и восстановления проверен отказ записи или чтения / register, IMOD/ITR and restore changes: write and read failures checked
- [ ] Для интерфейса: RU и EN, масштаб 100 / 125 / 150 % / UI: RU and EN, 100 / 125 / 150 % scaling
