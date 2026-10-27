#pragma once

#define FILE_DEVICE_IMOD 0x00008010
#define IMOD_IOCTL_INDEX 0x810

#define IOCTL_IMOD_READ_PHYSICAL \
    CTL_CODE(FILE_DEVICE_IMOD, IMOD_IOCTL_INDEX + 2, METHOD_BUFFERED, FILE_ANY_ACCESS)

#define IOCTL_IMOD_WRITE_PHYSICAL \
    CTL_CODE(FILE_DEVICE_IMOD, IMOD_IOCTL_INDEX + 3, METHOD_BUFFERED, FILE_ANY_ACCESS)

// Versioned, read-only query for one exact PCI devnode. No PCI writes.
#define IOCTL_IMOD_QUERY_PCI \
    CTL_CODE(FILE_DEVICE_IMOD, IMOD_IOCTL_INDEX + 4, METHOD_BUFFERED, FILE_READ_ACCESS)
#define IMOD_PCI_QUERY_VERSION 1
#define IMOD_PCI_CONFIG_SIZE 256
#define IMOD_PCI_RESOURCE_SIZE 65536

struct ImodPciQuery {
    ULONG version;
    WCHAR pdoName[256];
    WCHAR instanceId[256];
};

struct ImodPciReply {
    ULONG version;
    NTSTATUS configStatus;
    NTSTATUS resourcesStatus;
    ULONG configBytes;
    ULONG resourcesBytes;
    UCHAR config[IMOD_PCI_CONFIG_SIZE];
    UCHAR resources[IMOD_PCI_RESOURCE_SIZE];
};

#pragma pack(push, 1)

struct tagPhysAccessStruct
{
    ULONGLONG physAddress;
    ULONG accessSizeInBytes;
    ULONG reserved;
    ULONGLONG value;
};

#pragma pack(pop)
