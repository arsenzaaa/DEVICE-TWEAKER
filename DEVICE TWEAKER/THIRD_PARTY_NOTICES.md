# Сторонние компоненты

## KDU

В проекте используется [hfiref0x/KDU](https://github.com/hfiref0x/KDU) для загрузки `DTIMOD.sys`, чтобы программа могла читать и записывать USB IMOD и NIC ITR.

Файлы, которые использует программа:

- `IMOD/Loader/kdu.exe`
- `IMOD/Loader/drv64.dll`

Версия загрузчиков: `1.4.5.2512`.

Лицензия KDU: MIT. Полный текст находится в `IMOD/KDU/LICENSE.txt`.

Контрольные суммы:

```text
B340DAD4DDBE8607F9FDDB79F679375B4FF5080FE1A7EDB6CE015F69D3A0CD4F  kdu.exe
D032D855A0FF0EF3E3AD6EC8DAFFB8649048F82D3F1EEB89BEE3652EE6F01F80  drv64.dll
```

## Прочее

Приложение собирается под .NET 8 для Windows и использует Windows Forms, WMI и Win32 API.

Таблицы PCI ID для оценки `CHIP 0` / `CHIP 1` адаптированы из проекта [MariusHeier/cpu-direct-usb](https://github.com/MariusHeier/cpu-direct-usb) и связанных источников. Это оценка класса контроллера, а не измеренная задержка.
