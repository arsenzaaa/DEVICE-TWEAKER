#pragma once

#define FILE_DEVICE_IMOD 0x00008010
#define IMOD_IOCTL_INDEX 0x810

// Indices + 2 (unrestricted reads) and + 3 (unrestricted writes) are retired.

// Index + 3 was an unrestricted physical write; it is no longer accepted.

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

// Writes one register of an exact PCI devnode. The address must lie entirely
// inside a memory range Windows assigned to that device; anything else is
// rejected with STATUS_ACCESS_DENIED. The reply returns the checked resources.
#define IOCTL_IMOD_WRITE_DEVICE \
    CTL_CODE(FILE_DEVICE_IMOD, IMOD_IOCTL_INDEX + 5, METHOD_BUFFERED, FILE_WRITE_ACCESS)
#define IMOD_DEVICE_WRITE_VERSION 1

struct ImodDeviceWrite {
    ULONG version;
    ULONG accessSize;
    ULONGLONG physAddress;
    ULONGLONG value;
    WCHAR pdoName[256];
    WCHAR instanceId[256];
};

struct ImodDeviceWriteReply {
    ULONG version;
    ULONG resourcesBytes;
    UCHAR resources[IMOD_PCI_RESOURCE_SIZE];
};

// Device-scoped MMIO reads use ImodDeviceWrite as their request layout.
#define IOCTL_IMOD_READ_DEVICE \
    CTL_CODE(FILE_DEVICE_IMOD, IMOD_IOCTL_INDEX + 6, METHOD_BUFFERED, FILE_READ_ACCESS)
struct ImodDeviceReadReply {
    ULONG version;
    ULONG reserved;
    ULONGLONG value;
};

// Controller contexts are read internally; callers cannot supply a RAM address.
#define IOCTL_IMOD_QUERY_XHCI \
    CTL_CODE(FILE_DEVICE_IMOD, IMOD_IOCTL_INDEX + 7, METHOD_BUFFERED, FILE_READ_ACCESS)
#define IMOD_XHCI_QUERY_VERSION 1
#define IMOD_XHCI_MAX_SLOTS 255
struct ImodXhciQuery {
    ULONG version;
    ULONG hcsparamsOffset;
    ULONGLONG capabilityAddress;
    ULONGLONG memoryLength;
    WCHAR pdoName[256];
    WCHAR instanceId[256];
};
struct ImodXhciSlot {
    ULONG rootPort;
    ULONG deviceAddress;
    ULONG interrupter;
};
struct ImodXhciReply {
    ULONG version;
    ULONG maxSlots;
    ULONG maxIntrs;
    ULONG contextSize;
    ULONG count;
    ULONG endpointTargets;
    ULONG slotTargets;
    struct ImodXhciSlot slots[IMOD_XHCI_MAX_SLOTS];
};
