#define NOMINMAX
#include <windows.h>
#include <cfgmgr32.h>
#include <setupapi.h>

#include <bitset>
#include <cstdint>
#include <cwchar>
#include <cwctype>
#include <algorithm>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <iterator>
#include <limits>
#include <optional>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "cfgmgr32.lib")
#pragma comment(lib, "setupapi.lib")

namespace {

constexpr uint32_t kDefaultInterval = 0x0;
constexpr uint32_t kDefaultHcsparamsOffset = 0x4;
constexpr uint32_t kDefaultRtsoff = 0x18;

constexpr const wchar_t* kWinIoServiceName = L"WINIO";
constexpr const wchar_t* kDriverFileName = L"winio.sys";
constexpr const wchar_t* kConfigFileName = L"imod-config.ini";

constexpr uint32_t FILE_DEVICE_WINIO = 0x00008010;
constexpr uint32_t WINIO_IOCTL_INDEX = 0x810;

constexpr uint32_t IoctlWinioMapPhysToLin =
    CTL_CODE(FILE_DEVICE_WINIO, WINIO_IOCTL_INDEX, METHOD_BUFFERED, FILE_ANY_ACCESS);
constexpr uint32_t IoctlWinioUnmapPhysAddr =
    CTL_CODE(FILE_DEVICE_WINIO, WINIO_IOCTL_INDEX + 1, METHOD_BUFFERED, FILE_ANY_ACCESS);
constexpr uint32_t IoctlWinioEnableDirectIo =
    CTL_CODE(FILE_DEVICE_WINIO, WINIO_IOCTL_INDEX + 2, METHOD_BUFFERED, FILE_ANY_ACCESS);
constexpr uint32_t IoctlWinioDisableDirectIo =
    CTL_CODE(FILE_DEVICE_WINIO, WINIO_IOCTL_INDEX + 3, METHOD_BUFFERED, FILE_ANY_ACCESS);

#pragma pack(push, 1)
struct PhysStruct {
    uint64_t physMemSizeInBytes;
    uint64_t physAddress;
    uint64_t physicalMemoryHandle;
    uint64_t physMemLin;
    uint64_t physSection;
};
#pragma pack(pop)

struct UsbControllerInfo {
    std::wstring deviceId;
    std::wstring caption;
    ULONG problemCode = 0;
    uint64_t baseAddress = 0;
    bool hasBase = false;
    std::wstring baseError;
};

struct ControllerOverride {
    std::wstring hwid;
    std::optional<uint32_t> interval;
    std::optional<uint32_t> hcsparamsOffset;
    std::optional<uint32_t> rtsoff;
    std::optional<bool> enabled;
};

struct Config {
    uint32_t globalInterval = kDefaultInterval;
    uint32_t globalHcsparamsOffset = kDefaultHcsparamsOffset;
    uint32_t globalRtsoff = kDefaultRtsoff;
    std::vector<ControllerOverride> overrides;
};

struct WinIoContext {
    HANDLE driverHandle = INVALID_HANDLE_VALUE;
    bool is64BitOS = false;
    bool serviceCreated = false;
    std::wstring driverPath;
};

template <typename F>
class ScopeExit {
public:
    explicit ScopeExit(F func) : func_(std::move(func)), active_(true) {}
    ~ScopeExit() {
        if (active_) {
            func_();
        }
    }
    void Release() { active_ = false; }

private:
    F func_;
    bool active_;
};

std::wstring ToUpper(std::wstring value) {
    for (auto& ch : value) {
        ch = static_cast<wchar_t>(towupper(ch));
    }
    return value;
}

std::wstring Trim(std::wstring value) {
    const auto isSpace = [](wchar_t ch) {
        return ch == L' ' || ch == L'\t' || ch == L'\r' || ch == L'\n';
    };
    size_t start = 0;
    while (start < value.size() && isSpace(value[start])) {
        ++start;
    }
    size_t end = value.size();
    while (end > start && isSpace(value[end - 1])) {
        --end;
    }
    return value.substr(start, end - start);
}

bool ContainsInsensitive(const std::wstring& value, const std::wstring& needle) {
    return ToUpper(value).find(ToUpper(needle)) != std::wstring::npos;
}

bool EqualsInsensitive(const std::wstring& left, const std::wstring& right) {
    return ToUpper(left) == ToUpper(right);
}

bool StartsWithInsensitive(const std::wstring& value, const std::wstring& prefix) {
    if (value.size() < prefix.size()) {
        return false;
    }
    return ToUpper(value.substr(0, prefix.size())) == ToUpper(prefix);
}

std::wstring GetLastErrorMessage(DWORD error) {
    LPWSTR buffer = nullptr;
    const DWORD flags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
    const DWORD size = FormatMessageW(flags, nullptr, error, 0, reinterpret_cast<LPWSTR>(&buffer), 0, nullptr);
    std::wstring message = (size && buffer) ? buffer : L"";
    if (buffer) {
        LocalFree(buffer);
    }
    while (!message.empty() && (message.back() == L'\r' || message.back() == L'\n')) {
        message.pop_back();
    }
    return message;
}

std::wstring ToHex(uint64_t value) {
    std::wostringstream stream;
    stream << L"0x" << std::uppercase << std::hex << value;
    return stream.str();
}

std::wstring StripInlineComment(const std::wstring& value) {
    size_t hashPos = value.find(L'#');
    size_t semiPos = value.find(L';');
    size_t cut = std::wstring::npos;
    if (hashPos != std::wstring::npos) {
        cut = hashPos;
    }
    if (semiPos != std::wstring::npos) {
        cut = (cut == std::wstring::npos) ? semiPos : std::min(cut, semiPos);
    }
    if (cut == std::wstring::npos) {
        return value;
    }
    return value.substr(0, cut);
}

bool TryParseUint32(const std::wstring& text, uint32_t* value) {
    std::wstring trimmed = Trim(text);
    if (trimmed.empty()) {
        return false;
    }
    try {
        size_t index = 0;
        unsigned long long parsed = std::stoull(trimmed, &index, 0);
        if (index != trimmed.size()) {
            return false;
        }
        if (parsed > std::numeric_limits<uint32_t>::max()) {
            return false;
        }
        *value = static_cast<uint32_t>(parsed);
        return true;
    } catch (...) {
        return false;
    }
}

bool TryParseBool(const std::wstring& text, bool* value) {
    std::wstring trimmed = ToUpper(Trim(text));
    if (trimmed == L"TRUE" || trimmed == L"1") {
        *value = true;
        return true;
    }
    if (trimmed == L"FALSE" || trimmed == L"0") {
        *value = false;
        return true;
    }

    uint32_t numeric = 0;
    if (TryParseUint32(trimmed, &numeric)) {
        *value = numeric != 0;
        return true;
    }

    return false;
}

bool FileExists(const std::wstring& path) {
    const DWORD attrs = GetFileAttributesW(path.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES) {
        return false;
    }
    return (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

std::wstring GetModuleDirectory() {
    wchar_t buffer[MAX_PATH] = {};
    const DWORD length = GetModuleFileNameW(nullptr, buffer, static_cast<DWORD>(std::size(buffer)));
    if (length == 0 || length >= std::size(buffer)) {
        return L"";
    }
    std::wstring path(buffer, length);
    const size_t slashPos = path.find_last_of(L"\\/");
    if (slashPos == std::wstring::npos) {
        return L"";
    }
    return path.substr(0, slashPos);
}

std::wstring GetCurrentDirectoryPath() {
    wchar_t buffer[MAX_PATH] = {};
    const DWORD length = GetCurrentDirectoryW(static_cast<DWORD>(std::size(buffer)), buffer);
    if (length == 0 || length >= std::size(buffer)) {
        return L"";
    }
    return std::wstring(buffer, length);
}

std::wstring ParentPath(const std::wstring& path) {
    const size_t slashPos = path.find_last_of(L"\\/");
    if (slashPos == std::wstring::npos) {
        return L"";
    }
    return path.substr(0, slashPos);
}

std::wstring FindFileInCommonPaths(const wchar_t* fileName) {
    std::vector<std::wstring> candidates;

    const std::wstring cwd = GetCurrentDirectoryPath();
    if (!cwd.empty()) {
        candidates.push_back(cwd + L"\\" + fileName);
    }

    const std::wstring exeDir = GetModuleDirectory();
    if (!exeDir.empty()) {
        candidates.push_back(exeDir + L"\\" + fileName);
        const std::wstring exeParent = ParentPath(exeDir);
        if (!exeParent.empty()) {
            candidates.push_back(exeParent + L"\\" + fileName);
            const std::wstring exeGrandParent = ParentPath(exeParent);
            if (!exeGrandParent.empty()) {
                candidates.push_back(exeGrandParent + L"\\" + fileName);
            }
        }
    }

    for (const auto& candidate : candidates) {
        if (FileExists(candidate)) {
            return candidate;
        }
    }

    return L"";
}

std::wstring FindDriverPath() {
    return FindFileInCommonPaths(kDriverFileName);
}

std::wstring FindConfigPath() {
    return FindFileInCommonPaths(kConfigFileName);
}

bool IsAdmin() {
    BOOL isAdmin = FALSE;
    SID_IDENTIFIER_AUTHORITY ntAuthority = SECURITY_NT_AUTHORITY;
    PSID adminGroup = nullptr;
    if (AllocateAndInitializeSid(
            &ntAuthority, 2, SECURITY_BUILTIN_DOMAIN_RID, DOMAIN_ALIAS_RID_ADMINS,
            0, 0, 0, 0, 0, 0, &adminGroup)) {
        CheckTokenMembership(nullptr, adminGroup, &isAdmin);
        FreeSid(adminGroup);
    }
    return isAdmin == TRUE;
}

bool Is64BitOS() {
#if defined(_WIN64)
    return true;
#else
    BOOL wow64 = FALSE;
    if (!IsWow64Process(GetCurrentProcess(), &wow64)) {
        return false;
    }
    return wow64 == TRUE;
#endif
}

bool GetDevicePropertyData(HDEVINFO devInfoSet, SP_DEVINFO_DATA* devInfo, DWORD property,
                           std::vector<BYTE>* data, DWORD* regType) {
    DWORD requiredSize = 0;
    DWORD type = 0;
    if (!SetupDiGetDeviceRegistryPropertyW(devInfoSet, devInfo, property, &type, nullptr, 0, &requiredSize)) {
        const DWORD err = GetLastError();
        if (err != ERROR_INSUFFICIENT_BUFFER) {
            return false;
        }
    }

    if (requiredSize == 0) {
        return false;
    }

    data->resize(requiredSize);
    if (!SetupDiGetDeviceRegistryPropertyW(devInfoSet, devInfo, property, &type, data->data(), requiredSize, nullptr)) {
        return false;
    }

    if (regType) {
        *regType = type;
    }
    return true;
}

void ParseMultiSz(const wchar_t* data, size_t length, std::vector<std::wstring>* values) {
    size_t index = 0;
    while (index < length && data[index] != L'\0') {
        const size_t start = index;
        while (index < length && data[index] != L'\0') {
            ++index;
        }
        if (index > start) {
            values->emplace_back(data + start, index - start);
        }
        ++index;
    }
}

bool GetDeviceMultiSzProperty(HDEVINFO devInfoSet, SP_DEVINFO_DATA* devInfo, DWORD property,
                              std::vector<std::wstring>* values) {
    values->clear();
    std::vector<BYTE> data;
    DWORD type = 0;
    if (!GetDevicePropertyData(devInfoSet, devInfo, property, &data, &type)) {
        return false;
    }

    if (type != REG_MULTI_SZ && type != REG_SZ) {
        return false;
    }

    const wchar_t* text = reinterpret_cast<const wchar_t*>(data.data());
    if (!text) {
        return false;
    }

    const size_t length = data.size() / sizeof(wchar_t);
    if (length == 0) {
        return false;
    }

    if (type == REG_SZ) {
        values->emplace_back(text);
        return true;
    }

    ParseMultiSz(text, length, values);
    return !values->empty();
}

bool GetDeviceStringProperty(HDEVINFO devInfoSet, SP_DEVINFO_DATA* devInfo, DWORD property, std::wstring* value) {
    std::vector<std::wstring> values;
    if (!GetDeviceMultiSzProperty(devInfoSet, devInfo, property, &values)) {
        return false;
    }
    if (values.empty()) {
        return false;
    }
    *value = values.front();
    return true;
}

bool GetDeviceInstanceId(HDEVINFO devInfoSet, SP_DEVINFO_DATA* devInfo, std::wstring* id) {
    DWORD requiredSize = 0;
    SetupDiGetDeviceInstanceIdW(devInfoSet, devInfo, nullptr, 0, &requiredSize);
    if (GetLastError() != ERROR_INSUFFICIENT_BUFFER || requiredSize == 0) {
        return false;
    }

    std::wstring buffer(requiredSize, L'\0');
    if (!SetupDiGetDeviceInstanceIdW(devInfoSet, devInfo, buffer.data(), requiredSize, nullptr)) {
        return false;
    }

    buffer.resize(wcslen(buffer.c_str()));
    *id = buffer;
    return true;
}

std::wstring GetDeviceCaption(HDEVINFO devInfoSet, SP_DEVINFO_DATA* devInfo) {
    std::wstring caption;
    if (GetDeviceStringProperty(devInfoSet, devInfo, SPDRP_FRIENDLYNAME, &caption)) {
        return caption;
    }
    if (GetDeviceStringProperty(devInfoSet, devInfo, SPDRP_DEVICEDESC, &caption)) {
        return caption;
    }
    return L"Unknown USB Controller";
}

bool HasXhciClassCode(const std::vector<std::wstring>& ids) {
    for (const auto& id : ids) {
        if (ContainsInsensitive(id, L"CC_0C0330") || ContainsInsensitive(id, L"CLASS_0C0330")) {
            return true;
        }
    }
    return false;
}

bool IsXhciDevice(HDEVINFO devInfoSet, SP_DEVINFO_DATA* devInfo) {
    std::wstring service;
    if (GetDeviceStringProperty(devInfoSet, devInfo, SPDRP_SERVICE, &service)) {
        if (ToUpper(service) == L"USBXHCI") {
            return true;
        }
    }

    std::vector<std::wstring> ids;
    if (GetDeviceMultiSzProperty(devInfoSet, devInfo, SPDRP_HARDWAREID, &ids) && HasXhciClassCode(ids)) {
        return true;
    }
    if (GetDeviceMultiSzProperty(devInfoSet, devInfo, SPDRP_COMPATIBLEIDS, &ids) && HasXhciClassCode(ids)) {
        return true;
    }

    return false;
}

bool GetDeviceProblemCode(DEVINST devInst, ULONG* problemCode) {
    ULONG status = 0;
    ULONG problem = 0;
    const CONFIGRET cr = CM_Get_DevNode_Status(&status, &problem, devInst, 0);
    if (cr != CR_SUCCESS) {
        return false;
    }
    *problemCode = problem;
    return true;
}

bool ExtractBaseFromResource(RESOURCEID resType, const BYTE* data, size_t size, uint64_t* base) {
    if (resType == ResType_Mem) {
        if (size < sizeof(MEM_RESOURCE)) {
            return false;
        }
        const auto* mem = reinterpret_cast<const MEM_RESOURCE*>(data);
        uint64_t candidate = mem->MEM_Header.MD_Alloc_Base;
        if (candidate == 0 && mem->MEM_Header.MD_Count > 0) {
            candidate = mem->MEM_Data[0].MR_Min;
        }
        if (candidate == 0) {
            return false;
        }
        *base = candidate;
        return true;
    }

    if (resType == ResType_MemLarge) {
        if (size < sizeof(MEM_LARGE_RESOURCE)) {
            return false;
        }
        const auto* mem = reinterpret_cast<const MEM_LARGE_RESOURCE*>(data);
        uint64_t candidate = mem->MEM_LARGE_Header.MLD_Alloc_Base;
        if (candidate == 0 && mem->MEM_LARGE_Header.MLD_Count > 0) {
            candidate = mem->MEM_LARGE_Data[0].MLR_Min;
        }
        if (candidate == 0) {
            return false;
        }
        *base = candidate;
        return true;
    }

    return false;
}

bool LoadConfigFile(const std::wstring& path, Config* config, std::wstring* error) {
    std::ifstream file{std::filesystem::path(path)};
    if (!file) {
        if (error) {
            *error = L"failed to open config: " + path;
        }
        return false;
    }

    Config result;
    ControllerOverride* currentDevice = nullptr;
    bool inGlobal = true;

    std::string lineRaw;
    size_t lineNumber = 0;
    while (std::getline(file, lineRaw)) {
        ++lineNumber;
        if (!lineRaw.empty() && lineRaw.back() == '\r') {
            lineRaw.pop_back();
        }
        if (lineNumber == 1 && lineRaw.size() >= 3 &&
            static_cast<unsigned char>(lineRaw[0]) == 0xEF &&
            static_cast<unsigned char>(lineRaw[1]) == 0xBB &&
            static_cast<unsigned char>(lineRaw[2]) == 0xBF) {
            lineRaw.erase(0, 3);
        }

        std::wstring line(lineRaw.begin(), lineRaw.end());
        line = Trim(StripInlineComment(line));
        if (line.empty()) {
            continue;
        }

        if (line.front() == L'[' && line.back() == L']') {
            std::wstring section = Trim(line.substr(1, line.size() - 2));
            if (EqualsInsensitive(section, L"global")) {
                currentDevice = nullptr;
                inGlobal = true;
                continue;
            }

            if (StartsWithInsensitive(section, L"device:")) {
                std::wstring hwid = Trim(section.substr(7));
                if (!hwid.empty()) {
                    result.overrides.push_back({hwid, std::nullopt, std::nullopt, std::nullopt, std::nullopt});
                    currentDevice = &result.overrides.back();
                    inGlobal = false;
                }
                continue;
            }

            if (StartsWithInsensitive(section, L"device ")) {
                std::wstring hwid = Trim(section.substr(7));
                if (!hwid.empty()) {
                    result.overrides.push_back({hwid, std::nullopt, std::nullopt, std::nullopt, std::nullopt});
                    currentDevice = &result.overrides.back();
                    inGlobal = false;
                }
                continue;
            }

            continue;
        }

        const size_t eqPos = line.find(L'=');
        if (eqPos == std::wstring::npos) {
            continue;
        }

        std::wstring key = ToUpper(Trim(line.substr(0, eqPos)));
        std::wstring value = Trim(line.substr(eqPos + 1));

        if (key == L"ENABLED") {
            bool parsedEnabled = true;
            if (!TryParseBool(value, &parsedEnabled)) {
                if (error) {
                    *error = L"invalid enabled value at line " + std::to_wstring(lineNumber);
                }
                return false;
            }

            if (!inGlobal && currentDevice != nullptr) {
                currentDevice->enabled = parsedEnabled;
            }
            continue;
        }

        uint32_t parsed = 0;
        if (!TryParseUint32(value, &parsed)) {
            if (error) {
                *error = L"invalid number at line " + std::to_wstring(lineNumber);
            }
            return false;
        }

        auto applyValue = [&](ControllerOverride* target, const std::wstring& keyName, uint32_t val) {
            if (keyName == L"INTERVAL") {
                target->interval = val;
            } else if (keyName == L"HCSPARAMS_OFFSET" || keyName == L"HCSPARAPS_OFFSET") {
                target->hcsparamsOffset = val;
            } else if (keyName == L"RTSOFF") {
                target->rtsoff = val;
            }
        };

        if (inGlobal || currentDevice == nullptr) {
            if (key == L"INTERVAL") {
                result.globalInterval = parsed;
            } else if (key == L"HCSPARAMS_OFFSET" || key == L"HCSPARAPS_OFFSET") {
                result.globalHcsparamsOffset = parsed;
            } else if (key == L"RTSOFF") {
                result.globalRtsoff = parsed;
            }
        } else {
            applyValue(currentDevice, key, parsed);
        }
    }

    *config = std::move(result);
    return true;
}

bool GetDeviceMemoryBase(DEVINST devInst, uint64_t* base, std::wstring* error) {
    LOG_CONF logConf = 0;
    CONFIGRET cr = CM_Get_First_Log_Conf(&logConf, devInst, ALLOC_LOG_CONF);
    if (cr != CR_SUCCESS) {
        cr = CM_Get_First_Log_Conf(&logConf, devInst, BOOT_LOG_CONF);
    }
    if (cr != CR_SUCCESS) {
        if (error) {
            *error = L"failed to query logical config (CONFIGRET " + std::to_wstring(cr) + L")";
        }
        return false;
    }

    ScopeExit logConfCleanup([&]() { CM_Free_Log_Conf_Handle(logConf); });

    uint64_t minBase = 0;
    bool found = false;
    const RESOURCEID resTypes[] = {ResType_Mem, ResType_MemLarge};

    for (RESOURCEID resType : resTypes) {
        RES_DES resDes = 0;
        CONFIGRET resCr = CM_Get_Next_Res_Des(&resDes, logConf, resType, nullptr, 0);
        while (resCr == CR_SUCCESS) {
            ULONG dataSize = 0;
            const CONFIGRET sizeCr = CM_Get_Res_Des_Data_Size(&dataSize, resDes, 0);
            if (sizeCr == CR_SUCCESS && dataSize > 0) {
                std::vector<BYTE> buffer(dataSize);
                if (CM_Get_Res_Des_Data(resDes, buffer.data(), dataSize, 0) == CR_SUCCESS) {
                    uint64_t candidate = 0;
                    if (ExtractBaseFromResource(resType, buffer.data(), buffer.size(), &candidate)) {
                        if (!found || candidate < minBase) {
                            minBase = candidate;
                            found = true;
                        }
                    }
                }
            }

            RES_DES nextResDes = 0;
            const CONFIGRET nextCr = CM_Get_Next_Res_Des(&nextResDes, resDes, resType, nullptr, 0);
            CM_Free_Res_Des_Handle(resDes);
            resDes = nextResDes;
            resCr = nextCr;
        }
    }

    if (!found) {
        if (error) {
            *error = L"no memory resource found";
        }
        return false;
    }

    *base = minBase;
    return true;
}

bool EnumerateXhciControllers(std::vector<UsbControllerInfo>* out, std::wstring* error) {
    HDEVINFO devInfoSet = SetupDiGetClassDevsW(nullptr, L"PCI", nullptr, DIGCF_PRESENT | DIGCF_ALLCLASSES);
    if (devInfoSet == INVALID_HANDLE_VALUE) {
        if (error) {
            *error = L"failed to enumerate PCI devices: " + GetLastErrorMessage(GetLastError());
        }
        return false;
    }

    ScopeExit cleanup([&]() { SetupDiDestroyDeviceInfoList(devInfoSet); });

    std::vector<UsbControllerInfo> results;
    for (DWORD index = 0;; ++index) {
        SP_DEVINFO_DATA devInfo{};
        devInfo.cbSize = sizeof(devInfo);
        if (!SetupDiEnumDeviceInfo(devInfoSet, index, &devInfo)) {
            if (GetLastError() == ERROR_NO_MORE_ITEMS) {
                break;
            }
            if (error) {
                *error = L"failed to enumerate device info: " + GetLastErrorMessage(GetLastError());
            }
            return false;
        }

        if (!IsXhciDevice(devInfoSet, &devInfo)) {
            continue;
        }

        UsbControllerInfo info;
        if (!GetDeviceInstanceId(devInfoSet, &devInfo, &info.deviceId)) {
            continue;
        }

        info.caption = GetDeviceCaption(devInfoSet, &devInfo);
        GetDeviceProblemCode(devInfo.DevInst, &info.problemCode);

        std::wstring baseError;
        if (GetDeviceMemoryBase(devInfo.DevInst, &info.baseAddress, &baseError)) {
            info.hasBase = true;
        } else {
            info.baseError = baseError;
        }

        results.push_back(std::move(info));
    }

    *out = std::move(results);
    return true;
}

bool EnableDirectIo(const WinIoContext& ctx, std::wstring* error) {
    if (ctx.is64BitOS) {
        return true;
    }
    DWORD bytesReturned = 0;
    if (!DeviceIoControl(ctx.driverHandle, IoctlWinioEnableDirectIo, nullptr, 0, nullptr, 0, &bytesReturned, nullptr)) {
        if (error) {
            *error = L"failed to enable direct I/O: " + GetLastErrorMessage(GetLastError());
        }
        return false;
    }
    return true;
}

void DisableDirectIo(const WinIoContext& ctx) {
    if (ctx.is64BitOS || ctx.driverHandle == INVALID_HANDLE_VALUE) {
        return;
    }
    DWORD bytesReturned = 0;
    DeviceIoControl(ctx.driverHandle, IoctlWinioDisableDirectIo, nullptr, 0, nullptr, 0, &bytesReturned, nullptr);
}

bool MapPhysicalMemory(const WinIoContext& ctx, uint64_t address, uint64_t size, PhysStruct* out, std::wstring* error) {
    PhysStruct phys{};
    phys.physMemSizeInBytes = size;
    phys.physAddress = address;

    DWORD bytesReturned = 0;
    if (!DeviceIoControl(
            ctx.driverHandle,
            IoctlWinioMapPhysToLin,
            &phys, sizeof(phys),
            &phys, sizeof(phys),
            &bytesReturned,
            nullptr)) {
        if (error) {
            *error = L"failed to map physical memory: " + GetLastErrorMessage(GetLastError());
        }
        return false;
    }

    if (phys.physMemLin == 0) {
        if (error) {
            *error = L"failed to map physical memory: returned null linear address";
        }
        return false;
    }

    *out = phys;
    return true;
}

bool UnmapPhysicalMemory(const WinIoContext& ctx, const PhysStruct& phys, std::wstring* error) {
    PhysStruct local = phys;
    DWORD bytesReturned = 0;
    if (!DeviceIoControl(
            ctx.driverHandle,
            IoctlWinioUnmapPhysAddr,
            &local, sizeof(local),
            &local, sizeof(local),
            &bytesReturned,
            nullptr)) {
        if (error) {
            *error = L"failed to unmap physical memory: " + GetLastErrorMessage(GetLastError());
        }
        return false;
    }
    return true;
}

bool ReadPhys32(const WinIoContext& ctx, uint64_t address, uint32_t* value, std::wstring* error) {
    PhysStruct phys{};
    if (!MapPhysicalMemory(ctx, address, 4, &phys, error)) {
        return false;
    }

    const volatile uint32_t* ptr = reinterpret_cast<volatile uint32_t*>(static_cast<uintptr_t>(phys.physMemLin));
    const uint32_t result = *ptr;

    if (!UnmapPhysicalMemory(ctx, phys, error)) {
        return false;
    }

    *value = result;
    return true;
}

bool WritePhys32(const WinIoContext& ctx, uint64_t address, uint32_t value, std::wstring* error) {
    PhysStruct phys{};
    if (!MapPhysicalMemory(ctx, address, 4, &phys, error)) {
        return false;
    }

    volatile uint32_t* ptr = reinterpret_cast<volatile uint32_t*>(static_cast<uintptr_t>(phys.physMemLin));
    *ptr = value;

    if (!UnmapPhysicalMemory(ctx, phys, error)) {
        return false;
    }
    return true;
}

bool QueryServiceStatus(SC_HANDLE service, SERVICE_STATUS_PROCESS* status) {
    DWORD bytesNeeded = 0;
    return QueryServiceStatusEx(
        service, SC_STATUS_PROCESS_INFO, reinterpret_cast<LPBYTE>(status),
        sizeof(SERVICE_STATUS_PROCESS), &bytesNeeded) != FALSE;
}

bool EnsureWinIoService(WinIoContext& ctx, std::wstring* error) {
    SC_HANDLE scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_ALL_ACCESS);
    if (!scm) {
        if (error) {
            *error = L"failed to open service manager: " + GetLastErrorMessage(GetLastError());
        }
        return false;
    }

    SC_HANDLE service = OpenServiceW(scm, kWinIoServiceName, SERVICE_ALL_ACCESS);
    if (!service) {
        const DWORD lastError = GetLastError();
        if (lastError != ERROR_SERVICE_DOES_NOT_EXIST) {
            CloseServiceHandle(scm);
            if (error) {
                *error = L"failed to open WINIO service: " + GetLastErrorMessage(lastError);
            }
            return false;
        }

        service = CreateServiceW(
            scm,
            kWinIoServiceName,
            kWinIoServiceName,
            SERVICE_ALL_ACCESS,
            SERVICE_KERNEL_DRIVER,
            SERVICE_DEMAND_START,
            SERVICE_ERROR_NORMAL,
            ctx.driverPath.c_str(),
            nullptr,
            nullptr,
            nullptr,
            nullptr,
            nullptr);
        if (!service) {
            CloseServiceHandle(scm);
            if (error) {
                *error = L"failed to create WINIO service: " + GetLastErrorMessage(GetLastError());
            }
            return false;
        }
        ctx.serviceCreated = true;
    }

    SERVICE_STATUS_PROCESS status{};
    bool wasRunning = false;
    if (QueryServiceStatus(service, &status)) {
        wasRunning = status.dwCurrentState == SERVICE_RUNNING;
    }

    if (!wasRunning) {
        if (!StartServiceW(service, 0, nullptr)) {
            const DWORD lastError = GetLastError();
            if (lastError != ERROR_SERVICE_ALREADY_RUNNING) {
                CloseServiceHandle(service);
                CloseServiceHandle(scm);
                if (error) {
                    *error = L"failed to start WINIO service: " + GetLastErrorMessage(lastError);
                }
                return false;
            }
        }
    }

    CloseServiceHandle(service);
    CloseServiceHandle(scm);
    return true;
}

void StopWinIoServiceIfNeeded(const WinIoContext& ctx) {
    SC_HANDLE scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_ALL_ACCESS);
    if (!scm) {
        return;
    }

    SC_HANDLE service = OpenServiceW(scm, kWinIoServiceName, SERVICE_ALL_ACCESS);
    if (!service) {
        CloseServiceHandle(scm);
        return;
    }

    SERVICE_STATUS_PROCESS status{};
    if (QueryServiceStatus(service, &status) && status.dwCurrentState != SERVICE_STOPPED) {
        SERVICE_STATUS stopStatus{};
        ControlService(service, SERVICE_CONTROL_STOP, &stopStatus);

        for (int i = 0; i < 25; ++i) {
            SERVICE_STATUS_PROCESS check{};
            if (!QueryServiceStatus(service, &check) || check.dwCurrentState == SERVICE_STOPPED) {
                break;
            }
            Sleep(200);
        }
    }

    if (ctx.serviceCreated) {
        DeleteService(service);
    }

    CloseServiceHandle(service);
    CloseServiceHandle(scm);
}

bool InitializeWinIo(WinIoContext& ctx, std::wstring* error) {
    ctx.is64BitOS = Is64BitOS();

    if (!EnsureWinIoService(ctx, error)) {
        return false;
    }

    ctx.driverHandle = CreateFileW(
        L"\\\\.\\WINIO",
        GENERIC_READ | GENERIC_WRITE,
        FILE_SHARE_READ | FILE_SHARE_WRITE,
        nullptr,
        OPEN_EXISTING,
        FILE_ATTRIBUTE_NORMAL,
        nullptr);
    if (ctx.driverHandle == INVALID_HANDLE_VALUE) {
        if (error) {
            *error = L"failed to open \\\\.\\WINIO: " + GetLastErrorMessage(GetLastError());
        }
        return false;
    }

    if (!EnableDirectIo(ctx, error)) {
        return false;
    }

    return true;
}

void ShutdownWinIo(WinIoContext& ctx) {
    DisableDirectIo(ctx);

    if (ctx.driverHandle != INVALID_HANDLE_VALUE) {
        CloseHandle(ctx.driverHandle);
        ctx.driverHandle = INVALID_HANDLE_VALUE;
    }

    StopWinIoServiceIfNeeded(ctx);
}

}

int wmain(int argc, wchar_t* argv[]) {
    bool verbose = false;
    for (int i = 1; i < argc; ++i) {
        if (wcscmp(argv[i], L"-v") == 0 || wcscmp(argv[i], L"--verbose") == 0 || wcscmp(argv[i], L"/v") == 0) {
            verbose = true;
        }
    }

    if (!IsAdmin()) {
        std::wcout << L"error: administrator privileges required" << std::endl;
        return 1;
    }

    std::wstring driverPath = FindDriverPath();
    if (driverPath.empty()) {
        std::wcout << L"error: winio.sys not exists" << std::endl;
        return 1;
    }

    Config config;
    std::wstring configPath = FindConfigPath();
    if (!configPath.empty()) {
        std::wstring configError;
        if (!LoadConfigFile(configPath, &config, &configError)) {
            std::wcout << L"error: " << configError << std::endl;
            return 1;
        }
    }

    std::wstring enumError;
    std::vector<UsbControllerInfo> controllers;
    if (!EnumerateXhciControllers(&controllers, &enumError)) {
        std::wcout << L"error: " << enumError << std::endl;
        return 1;
    }

    WinIoContext winio;
    winio.driverPath = driverPath;
    std::wstring driverError;
    if (!InitializeWinIo(winio, &driverError)) {
        std::wcout << L"error: " << driverError << std::endl;
        ShutdownWinIo(winio);
        return 1;
    }

    ScopeExit cleanup([&]() { ShutdownWinIo(winio); });

    if (!configPath.empty()) {
        std::wcout << L"config = " << configPath << L" (overrides: " << config.overrides.size() << L")" << std::endl;
    } else {
        std::wcout << L"config = defaults (no " << kConfigFileName << L" found)" << std::endl;
    }
    std::wcout << L"winio.sys = " << driverPath << std::endl << std::endl;

    for (const auto& controller : controllers) {
        if (controller.problemCode == CM_PROB_DISABLED) {
            continue;
        }

        std::wcout << controller.caption << L" - " << controller.deviceId << std::endl;
        if (controller.problemCode != 0) {
            std::wcout << L"  problem_code = " << controller.problemCode << std::endl;
        }

        if (!controller.hasBase) {
            if (!controller.baseError.empty()) {
                std::wcout << L"  base_address = error: " << controller.baseError << std::endl << std::endl;
            } else {
                std::wcout << L"  base_address = error: could not obtain base address" << std::endl << std::endl;
            }
            continue;
        }

        uint32_t desiredInterval = config.globalInterval;
        uint32_t hcsparamsOffset = config.globalHcsparamsOffset;
        uint32_t rtsoff = config.globalRtsoff;
        bool enabled = true;
        std::wstring overrideMatch;

        for (const auto& entry : config.overrides) {
            if (ContainsInsensitive(controller.deviceId, entry.hwid)) {
                if (entry.enabled) {
                    enabled = *entry.enabled;
                }
                if (entry.interval) {
                    desiredInterval = *entry.interval;
                }
                if (entry.hcsparamsOffset) {
                    hcsparamsOffset = *entry.hcsparamsOffset;
                }
                if (entry.rtsoff) {
                    rtsoff = *entry.rtsoff;
                }
                overrideMatch = entry.hwid;
            }
        }

        const uint64_t capabilityAddress = controller.baseAddress;
        std::wcout << L"  base_address = " << ToHex(capabilityAddress) << std::endl;
        std::wcout << L"  interval = " << ToHex(desiredInterval);
        if (!overrideMatch.empty()) {
            std::wcout << L" (override: " << overrideMatch << L")";
        }
        std::wcout << std::endl;
        std::wcout << L"  hcsparams_offset = " << ToHex(hcsparamsOffset)
                   << L", rtsoff = " << ToHex(rtsoff) << std::endl;

        if (!enabled) {
            if (!overrideMatch.empty()) {
                std::wcout << L"  skipped (disabled by config: " << overrideMatch << L")" << std::endl << std::endl;
            } else {
                std::wcout << L"  skipped (disabled by config)" << std::endl << std::endl;
            }
            continue;
        }

        uint32_t hcsparamsValue = 0;
        uint32_t rtsoffValue = 0;
        std::wstring ioError;
        if (!ReadPhys32(winio, capabilityAddress + hcsparamsOffset, &hcsparamsValue, &ioError)) {
            std::wcout << L"error: failed to read XHCI registers: " << ioError << std::endl << std::endl;
            continue;
        }
        if (!ReadPhys32(winio, capabilityAddress + rtsoff, &rtsoffValue, &ioError)) {
            std::wcout << L"error: failed to read XHCI registers: " << ioError << std::endl << std::endl;
            continue;
        }

        const uint32_t maxIntrs = (hcsparamsValue >> 8) & 0xFF;
        const uint64_t runtimeAddress = capabilityAddress + rtsoffValue;

        std::wcout << L"  max_intrs = " << maxIntrs
                   << L", runtime_address = " << ToHex(runtimeAddress) << std::endl;

        if (verbose) {
            std::wcout << L"capability_address  = " << ToHex(capabilityAddress) << std::endl;
            std::wcout << L"HCSPARAMS_value     = capability_address + hcsparams_offset = "
                       << ToHex(capabilityAddress) << L" + " << ToHex(hcsparamsOffset) << L" = "
                       << ToHex(hcsparamsValue) << std::endl;
            std::wcout << L"HCSPARAMS_bitmask   = " << std::bitset<32>(hcsparamsValue) << std::endl;
            std::wcout << L"max_intrs           = " << maxIntrs << std::endl;
            std::wcout << L"RTSOFF_value        = capability_address + rtsoff = "
                       << ToHex(capabilityAddress) << L" + " << ToHex(rtsoff) << L" = "
                       << ToHex(rtsoffValue) << std::endl;
            std::wcout << L"runtime_address     = capability_address + RTSOFF_value = "
                       << ToHex(capabilityAddress) << L" + " << ToHex(rtsoffValue) << L" = "
                       << ToHex(runtimeAddress) << std::endl;
        }

        uint32_t writeFailures = 0;
        for (uint32_t i = 0; i < maxIntrs; ++i) {
            const uint64_t interrupterAddress = runtimeAddress + 0x24 + (0x20 * i);
            const std::wstring interrupterHex = ToHex(interrupterAddress);

            if (verbose) {
                std::wcout << std::endl;
                std::wcout << L"interrupter_address = runtime_address + 0x24 + (0x20 * index) = "
                           << ToHex(runtimeAddress) << L" + 0x24 + (0x20 * " << i << L") = "
                           << interrupterHex << std::endl;
                std::wcout << L"Write DWORD = " << ToHex(desiredInterval) << std::endl;
            }

            if (!WritePhys32(winio, interrupterAddress, desiredInterval, &ioError)) {
                std::wcout << L"error: failed to write IMOD interval at "
                           << interrupterHex << L": " << ioError << std::endl;
                ++writeFailures;
            }
        }

        std::wcout << L"  writes = " << maxIntrs << L", failures = " << writeFailures << std::endl;
        std::wcout << std::endl;
    }

    return 0;
}
