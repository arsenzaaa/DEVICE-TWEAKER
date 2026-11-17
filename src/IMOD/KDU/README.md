# KDU

[Documentation](../../../docs/DEVELOPMENT.md) / KDU · [Русский](./README.ru.md)

The application embeds `kdu.exe` / `drv64.dll` version **1.4.5.2512** from [DEVICE TWEAKER commit 6094c866](https://github.com/arsenzaaa/DEVICE-TWEAKER/commit/6094c866ac2d8baaf362d393f4911b76a4fc7908), which `v0.0.4` referenced at import time. Exact provenance and SHA-256 hashes are recorded in `../Loader/KDU.build.json`.

This directory preserves the corresponding source tree from that commit. The repository, commit and input hashes are recorded in `upstream.json`. See the author's [README.upstream.md](README.upstream.md) and the MIT [LICENSE.txt](LICENSE.txt).

`build-kdu.ps1` builds a source copy under `bin/KduBuild`. The application uses the separately pinned pair in `IMOD/Loader`; a regular application build does not rebuild it, and this script does not replace it automatically. A byte-identical rebuild of the published binaries is not claimed.
