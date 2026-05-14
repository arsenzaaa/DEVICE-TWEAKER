#include <ntddk.h>
#include <initguid.h>
#include <wdmsec.h>

#include "imod_driver.h"

extern NTKERNELAPI NTSTATUS IoCreateDriver(PUNICODE_STRING DriverName, PDRIVER_INITIALIZE InitializationFunction);

#define IMOD_MAPPED_DRIVER_NAME L"\\Driver\\DeviceTweakerImodMapped2"
#define IMOD_DEVICE_NAME L"\\Device\\DeviceTweakerImod2"
#define IMOD_DOS_DEVICE_NAME L"\\DosDevices\\DeviceTweakerImod2"
#define IMOD_MAX_MAP_SIZE PAGE_SIZE

DEFINE_GUID(GUID_DEVCLASS_DEVICE_TWEAKER_IMOD,
    0x605c1705, 0x1cf0, 0x49c7, 0xa0, 0x56, 0xd6, 0x87, 0xa3, 0x70, 0x7f, 0xd0);

static NTSTATUS ImodDriverDispatch(IN PDEVICE_OBJECT DeviceObject, IN PIRP Irp);
static void ImodDriverUnload(IN PDRIVER_OBJECT DriverObject);
static NTSTATUS ImodDriverInitialize(IN PDRIVER_OBJECT DriverObject);
static NTSTATUS ImodDriverCreateDriver(IN PDRIVER_OBJECT DriverObject, IN PUNICODE_STRING RegistryPath);
static NTSTATUS ImodMapPhysicalMemory(
    PHYSICAL_ADDRESS physicalAddress,
    SIZE_T physicalMemorySize,
    PVOID *mappedAddress,
    PVOID *mappedBase,
    SIZE_T *mappedSize);
static NTSTATUS ImodUnmapPhysicalMemory(
    PVOID mappedBase,
    SIZE_T mappedSize);
static NTSTATUS ImodReadPhysicalMemory(
    PHYSICAL_ADDRESS physicalAddress,
    ULONG accessSize,
    ULONGLONG *value);
static NTSTATUS ImodWritePhysicalMemory(
    PHYSICAL_ADDRESS physicalAddress,
    ULONG accessSize,
    ULONGLONG value);
static BOOLEAN ImodIsValidAccessSize(ULONG accessSize);

NTSTATUS DriverEntry(IN PDRIVER_OBJECT DriverObject, IN PUNICODE_STRING RegistryPath)
{
    if (DriverObject == NULL)
    {
        UNICODE_STRING driverName;

        UNREFERENCED_PARAMETER(RegistryPath);

        RtlInitUnicodeString(&driverName, IMOD_MAPPED_DRIVER_NAME);
        return IoCreateDriver(&driverName, ImodDriverCreateDriver);
    }

    UNREFERENCED_PARAMETER(RegistryPath);

    return ImodDriverInitialize(DriverObject);
}

static NTSTATUS ImodDriverCreateDriver(IN PDRIVER_OBJECT DriverObject, IN PUNICODE_STRING RegistryPath)
{
    UNREFERENCED_PARAMETER(RegistryPath);

    return ImodDriverInitialize(DriverObject);
}

static NTSTATUS ImodDriverInitialize(IN PDRIVER_OBJECT DriverObject)
{
    UNICODE_STRING deviceName;
    UNICODE_STRING deviceLink;
    UNICODE_STRING deviceSecurity;
    PDEVICE_OBJECT deviceObject = NULL;
    NTSTATUS status;

    KdPrint(("IMOD: DriverEntry"));

    RtlInitUnicodeString(&deviceName, IMOD_DEVICE_NAME);
    RtlInitUnicodeString(&deviceSecurity, L"D:P(A;;GA;;;SY)(A;;GA;;;BA)");

    status = IoCreateDeviceSecure(
        DriverObject,
        0,
        &deviceName,
        FILE_DEVICE_IMOD,
        FILE_DEVICE_SECURE_OPEN,
        FALSE,
        &deviceSecurity,
        &GUID_DEVCLASS_DEVICE_TWEAKER_IMOD,
        &deviceObject);

    if (!NT_SUCCESS(status))
    {
        KdPrint(("IMOD: IoCreateDeviceSecure failed: 0x%08X", status));
        return status;
    }

    DriverObject->MajorFunction[IRP_MJ_CREATE] =
        DriverObject->MajorFunction[IRP_MJ_CLOSE] =
        DriverObject->MajorFunction[IRP_MJ_DEVICE_CONTROL] = ImodDriverDispatch;
    DriverObject->DriverUnload = ImodDriverUnload;

    RtlInitUnicodeString(&deviceLink, IMOD_DOS_DEVICE_NAME);

    status = IoCreateSymbolicLink(&deviceLink, &deviceName);
    if (!NT_SUCCESS(status))
    {
        KdPrint(("IMOD: IoCreateSymbolicLink failed: 0x%08X", status));
        IoDeleteDevice(deviceObject);
        return status;
    }

    deviceObject->Flags &= ~DO_DEVICE_INITIALIZING;

    return STATUS_SUCCESS;
}

static NTSTATUS ImodDriverDispatch(IN PDEVICE_OBJECT DeviceObject, IN PIRP Irp)
{
    PIO_STACK_LOCATION irpStack = IoGetCurrentIrpStackLocation(Irp);
    ULONG inputLength;
    ULONG outputLength;
    ULONG ioControlCode;
    PVOID ioBuffer = Irp->AssociatedIrp.SystemBuffer;
    NTSTATUS status = STATUS_SUCCESS;
    struct tagPhysStruct phys;
    struct tagPhysAccessStruct access;
    PHYSICAL_ADDRESS physicalAddress;

    UNREFERENCED_PARAMETER(DeviceObject);

    Irp->IoStatus.Status = STATUS_SUCCESS;
    Irp->IoStatus.Information = 0;

    switch (irpStack->MajorFunction)
    {
    case IRP_MJ_CREATE:
    case IRP_MJ_CLOSE:
        break;

    case IRP_MJ_DEVICE_CONTROL:
        inputLength = irpStack->Parameters.DeviceIoControl.InputBufferLength;
        outputLength = irpStack->Parameters.DeviceIoControl.OutputBufferLength;
        ioControlCode = irpStack->Parameters.DeviceIoControl.IoControlCode;

        switch (ioControlCode)
        {
        case IOCTL_IMOD_MAP_PHYSICAL:
            if (ioBuffer == NULL || inputLength < sizeof(phys) || outputLength < sizeof(phys))
            {
                status = STATUS_INVALID_PARAMETER;
                break;
            }

            PVOID mappedAddress = NULL;
            PVOID mappedBase = NULL;
            SIZE_T mappedSize = 0;

            RtlCopyMemory(&phys, ioBuffer, sizeof(phys));

            physicalAddress.QuadPart = (LONGLONG)(ULONG_PTR)phys.physAddress;
            status = ImodMapPhysicalMemory(
                physicalAddress,
                (SIZE_T)phys.physMemSizeInBytes,
                &mappedAddress,
                &mappedBase,
                &mappedSize);

            if (NT_SUCCESS(status))
            {
                phys.physMemLin = (ULONGLONG)(ULONG_PTR)mappedAddress;
                phys.physicalMemoryHandle = (ULONGLONG)(ULONG_PTR)mappedBase;
                phys.physSection = (ULONGLONG)mappedSize;
                RtlCopyMemory(ioBuffer, &phys, sizeof(phys));
                Irp->IoStatus.Information = sizeof(phys);
            }

            break;

        case IOCTL_IMOD_UNMAP_PHYSICAL:
            if (ioBuffer == NULL || inputLength < sizeof(phys))
            {
                status = STATUS_INVALID_PARAMETER;
                break;
            }

            RtlCopyMemory(&phys, ioBuffer, sizeof(phys));

            if (phys.physicalMemoryHandle == 0 || phys.physSection == 0)
            {
                status = STATUS_INVALID_PARAMETER;
                break;
            }

            status = ImodUnmapPhysicalMemory(
                (PVOID)(ULONG_PTR)phys.physicalMemoryHandle,
                (SIZE_T)phys.physSection);

            break;

        case IOCTL_IMOD_READ_PHYSICAL:
            if (ioBuffer == NULL || inputLength < sizeof(access) || outputLength < sizeof(access))
            {
                status = STATUS_INVALID_PARAMETER;
                break;
            }

            RtlCopyMemory(&access, ioBuffer, sizeof(access));

            physicalAddress.QuadPart = (LONGLONG)(ULONG_PTR)access.physAddress;
            status = ImodReadPhysicalMemory(
                physicalAddress,
                access.accessSizeInBytes,
                &access.value);

            if (NT_SUCCESS(status))
            {
                RtlCopyMemory(ioBuffer, &access, sizeof(access));
                Irp->IoStatus.Information = sizeof(access);
            }

            break;

        case IOCTL_IMOD_WRITE_PHYSICAL:
            if (ioBuffer == NULL || inputLength < sizeof(access) || outputLength < sizeof(access))
            {
                status = STATUS_INVALID_PARAMETER;
                break;
            }

            RtlCopyMemory(&access, ioBuffer, sizeof(access));

            physicalAddress.QuadPart = (LONGLONG)(ULONG_PTR)access.physAddress;
            status = ImodWritePhysicalMemory(
                physicalAddress,
                access.accessSizeInBytes,
                access.value);

            if (NT_SUCCESS(status))
            {
                Irp->IoStatus.Information = sizeof(access);
            }

            break;

        default:
            status = STATUS_INVALID_DEVICE_REQUEST;
            break;
        }

        break;

    default:
        status = STATUS_INVALID_DEVICE_REQUEST;
        break;
    }

    Irp->IoStatus.Status = status;
    IoCompleteRequest(Irp, IO_NO_INCREMENT);

    return status;
}

static void ImodDriverUnload(IN PDRIVER_OBJECT DriverObject)
{
    UNICODE_STRING deviceLink;
    NTSTATUS status;

    RtlInitUnicodeString(&deviceLink, IMOD_DOS_DEVICE_NAME);

    status = IoDeleteSymbolicLink(&deviceLink);
    if (!NT_SUCCESS(status))
    {
        KdPrint(("IMOD: IoDeleteSymbolicLink failed: 0x%08X", status));
    }

    if (DriverObject->DeviceObject != NULL)
    {
        IoDeleteDevice(DriverObject->DeviceObject);
    }
}

static BOOLEAN ImodIsValidAccessSize(ULONG accessSize)
{
    return accessSize == sizeof(UCHAR) ||
        accessSize == sizeof(USHORT) ||
        accessSize == sizeof(ULONG);
}

static NTSTATUS ImodMapPhysicalMemory(
    PHYSICAL_ADDRESS physicalAddress,
    SIZE_T physicalMemorySize,
    PVOID *mappedAddress,
    PVOID *mappedBase,
    SIZE_T *mappedSize)
{
    PHYSICAL_ADDRESS alignedPhysicalAddress;
    SIZE_T pageMask;
    SIZE_T offset;
    SIZE_T totalSize;
    SIZE_T alignedSize;
    PVOID baseAddress;

    *mappedAddress = NULL;
    *mappedBase = NULL;
    *mappedSize = 0;

    if (physicalMemorySize == 0 || physicalMemorySize > IMOD_MAX_MAP_SIZE)
    {
        return STATUS_INVALID_PARAMETER;
    }

    pageMask = PAGE_SIZE - 1;
    offset = (SIZE_T)(physicalAddress.QuadPart & (LONGLONG)pageMask);
    if (physicalMemorySize > (SIZE_T)-1 - offset)
    {
        return STATUS_INVALID_PARAMETER;
    }

    totalSize = physicalMemorySize + offset;
    alignedSize = (totalSize + pageMask) & ~pageMask;
    alignedPhysicalAddress.QuadPart = physicalAddress.QuadPart & ~((LONGLONG)pageMask);

    baseAddress = MmMapIoSpace(alignedPhysicalAddress, alignedSize, MmNonCached);
    if (baseAddress == NULL)
    {
        KdPrint(("IMOD: MmMapIoSpace failed"));
        return STATUS_INSUFFICIENT_RESOURCES;
    }

    *mappedBase = baseAddress;
    *mappedSize = alignedSize;
    *mappedAddress = (PVOID)((PUCHAR)baseAddress + offset);

    return STATUS_SUCCESS;
}

static NTSTATUS ImodUnmapPhysicalMemory(
    PVOID mappedBase,
    SIZE_T mappedSize)
{
    if (mappedBase == NULL || mappedSize == 0)
    {
        return STATUS_INVALID_PARAMETER;
    }

    MmUnmapIoSpace(mappedBase, mappedSize);
    return STATUS_SUCCESS;
}

static NTSTATUS ImodReadPhysicalMemory(
    PHYSICAL_ADDRESS physicalAddress,
    ULONG accessSize,
    ULONGLONG *value)
{
    PVOID mappedAddress = NULL;
    PVOID mappedBase = NULL;
    SIZE_T mappedSize = 0;
    NTSTATUS status;

    if (value == NULL || !ImodIsValidAccessSize(accessSize))
    {
        return STATUS_INVALID_PARAMETER;
    }

    status = ImodMapPhysicalMemory(
        physicalAddress,
        accessSize,
        &mappedAddress,
        &mappedBase,
        &mappedSize);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    __try
    {
        switch (accessSize)
        {
        case sizeof(UCHAR):
            *value = READ_REGISTER_UCHAR((volatile UCHAR *)mappedAddress);
            break;
        case sizeof(USHORT):
            *value = READ_REGISTER_USHORT((volatile USHORT *)mappedAddress);
            break;
        case sizeof(ULONG):
            *value = READ_REGISTER_ULONG((volatile ULONG *)mappedAddress);
            break;
        default:
            status = STATUS_INVALID_PARAMETER;
            break;
        }
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        status = GetExceptionCode();
    }

    (void)ImodUnmapPhysicalMemory(mappedBase, mappedSize);
    return status;
}

static NTSTATUS ImodWritePhysicalMemory(
    PHYSICAL_ADDRESS physicalAddress,
    ULONG accessSize,
    ULONGLONG value)
{
    PVOID mappedAddress = NULL;
    PVOID mappedBase = NULL;
    SIZE_T mappedSize = 0;
    NTSTATUS status;

    if (!ImodIsValidAccessSize(accessSize))
    {
        return STATUS_INVALID_PARAMETER;
    }

    status = ImodMapPhysicalMemory(
        physicalAddress,
        accessSize,
        &mappedAddress,
        &mappedBase,
        &mappedSize);
    if (!NT_SUCCESS(status))
    {
        return status;
    }

    __try
    {
        switch (accessSize)
        {
        case sizeof(UCHAR):
            WRITE_REGISTER_UCHAR((volatile UCHAR *)mappedAddress, (UCHAR)value);
            break;
        case sizeof(USHORT):
            WRITE_REGISTER_USHORT((volatile USHORT *)mappedAddress, (USHORT)value);
            break;
        case sizeof(ULONG):
            WRITE_REGISTER_ULONG((volatile ULONG *)mappedAddress, (ULONG)value);
            break;
        default:
            status = STATUS_INVALID_PARAMETER;
            break;
        }
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        status = GetExceptionCode();
    }

    (void)ImodUnmapPhysicalMemory(mappedBase, mappedSize);
    return status;
}
