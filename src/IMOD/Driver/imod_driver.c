#include <ntifs.h>
#include <initguid.h>
#include <wdmsec.h>
#include <wdmguid.h>
#include <devpkey.h>

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
static NTSTATUS ImodQueryPci(struct ImodPciQuery *query, struct ImodPciReply *reply);
static IO_WORKITEM_ROUTINE_EX ImodQueryPciWorker;
static NTSTATUS ImodOpenPciDevice(const WCHAR *pdoName, const WCHAR *instanceId,
    PFILE_OBJECT *fileObject, PDEVICE_OBJECT *pdo);
static BOOLEAN ImodResourcesContainRange(const UCHAR *list, ULONG bytes, ULONGLONG address, ULONG size);
static NTSTATUS ImodWriteDevice(struct ImodDeviceWrite *request, struct ImodDeviceWriteReply *reply);
static IO_WORKITEM_ROUTINE_EX ImodWriteDeviceWorker;
C_ASSERT(sizeof(CM_PARTIAL_RESOURCE_DESCRIPTOR) == 20);
C_ASSERT(FIELD_OFFSET(struct ImodPciReply, resources) == 276);
// ImodResourcesContainRange walks CM_RESOURCE_LIST bytes with these offsets.
C_ASSERT(FIELD_OFFSET(CM_RESOURCE_LIST, List) == 4);
C_ASSERT(FIELD_OFFSET(CM_FULL_RESOURCE_DESCRIPTOR, PartialResourceList.Count) == 12);
C_ASSERT(FIELD_OFFSET(CM_FULL_RESOURCE_DESCRIPTOR, PartialResourceList.PartialDescriptors) == 16);
C_ASSERT(FIELD_OFFSET(CM_PARTIAL_RESOURCE_DESCRIPTOR, Flags) == 2);
C_ASSERT(FIELD_OFFSET(CM_PARTIAL_RESOURCE_DESCRIPTOR, u.Memory.Start) == 4);
C_ASSERT(FIELD_OFFSET(CM_PARTIAL_RESOURCE_DESCRIPTOR, u.Memory.Length) == 12);
C_ASSERT(FIELD_OFFSET(CM_PARTIAL_RESOURCE_DESCRIPTOR, u.Memory40.Length40) == 12);

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
        case IOCTL_IMOD_QUERY_PCI:
            if (ioBuffer == NULL || inputLength != sizeof(struct ImodPciQuery)
                || outputLength < sizeof(struct ImodPciReply))
            {
                status = STATUS_INVALID_PARAMETER;
                break;
            }
            {
                PIO_WORKITEM workItem = IoAllocateWorkItem(DeviceObject);
                if (workItem == NULL) { status = STATUS_INSUFFICIENT_RESOURCES; break; }
                // IoGetDevicePropertyData requires a system thread, not the IOCTL caller's context.
                IoMarkIrpPending(Irp);
                IoQueueWorkItemEx(workItem, ImodQueryPciWorker, DelayedWorkQueue, Irp);
                return STATUS_PENDING;
            }
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

        case IOCTL_IMOD_WRITE_DEVICE:
            if (ioBuffer == NULL || inputLength != sizeof(struct ImodDeviceWrite)
                || outputLength < sizeof(struct ImodDeviceWriteReply))
            {
                status = STATUS_INVALID_PARAMETER;
                break;
            }
            {
                PIO_WORKITEM workItem = IoAllocateWorkItem(DeviceObject);
                if (workItem == NULL) { status = STATUS_INSUFFICIENT_RESOURCES; break; }
                // Device properties are read from a system thread, as for the PCI query.
                IoMarkIrpPending(Irp);
                IoQueueWorkItemEx(workItem, ImodWriteDeviceWorker, DelayedWorkQueue, Irp);
                return STATUS_PENDING;
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

static VOID ImodQueryPciWorker(PVOID IoObject, PVOID Context, PIO_WORKITEM WorkItem)
{
    PIRP irp = (PIRP)Context;
    struct ImodPciQuery query;
    struct ImodPciReply *reply = (struct ImodPciReply *)irp->AssociatedIrp.SystemBuffer;
    NTSTATUS status;
    UNREFERENCED_PARAMETER(IoObject);
    // METHOD_BUFFERED shares input/output; retain the request before clearing the reply.
    RtlCopyMemory(&query, reply, sizeof(query));
    RtlZeroMemory(reply, sizeof(*reply));
    status = irp->Cancel ? STATUS_CANCELLED : ImodQueryPci(&query, reply);
    irp->IoStatus.Status = status;
    irp->IoStatus.Information = NT_SUCCESS(status)
        ? FIELD_OFFSET(struct ImodPciReply, resources) + reply->resourcesBytes : 0;
    IoFreeWorkItem(WorkItem);
    IoCompleteRequest(irp, IO_NO_INCREMENT);
}

static NTSTATUS ImodQueryPci(struct ImodPciQuery *query, struct ImodPciReply *reply)
{
    PFILE_OBJECT fileObject = NULL;
    PDEVICE_OBJECT pdo = NULL, target = NULL;
    ULONG bytes = 0;
    BUS_INTERFACE_STANDARD bus;
    NTSTATUS status;
    KEVENT event;
    IO_STATUS_BLOCK ioStatus;
    PIRP irp;
    PIO_STACK_LOCATION stack;

    if (KeGetCurrentIrql() != PASSIVE_LEVEL || query->version != IMOD_PCI_QUERY_VERSION)
        return STATUS_INVALID_PARAMETER;
    status = ImodOpenPciDevice(query->pdoName, query->instanceId, &fileObject, &pdo);
    if (!NT_SUCCESS(status)) return status;

    reply->version = IMOD_PCI_QUERY_VERSION;
    reply->resourcesStatus = IoGetDeviceProperty(pdo, DevicePropertyAllocatedResources,
        sizeof(reply->resources), reply->resources, &bytes);
    if (NT_SUCCESS(reply->resourcesStatus) && bytes <= sizeof(reply->resources))
        reply->resourcesBytes = bytes;
    else if (NT_SUCCESS(reply->resourcesStatus)) reply->resourcesStatus = STATUS_DATA_ERROR;

    RtlZeroMemory(&bus, sizeof(bus));
    KeInitializeEvent(&event, NotificationEvent, FALSE);
    target = IoGetAttachedDeviceReference(pdo);
    irp = IoBuildSynchronousFsdRequest(IRP_MJ_PNP, target, NULL, 0, NULL, &event, &ioStatus);
    if (irp == NULL) { reply->configStatus = STATUS_INSUFFICIENT_RESOURCES; goto Success; }
    stack = IoGetNextIrpStackLocation(irp);
    stack->MinorFunction = IRP_MN_QUERY_INTERFACE;
    stack->Parameters.QueryInterface.InterfaceType = (LPGUID)&GUID_BUS_INTERFACE_STANDARD;
    stack->Parameters.QueryInterface.Size = sizeof(bus);
    stack->Parameters.QueryInterface.Version = 1;
    stack->Parameters.QueryInterface.Interface = (PINTERFACE)&bus;
    stack->Parameters.QueryInterface.InterfaceSpecificData = NULL;
    irp->IoStatus.Status = STATUS_NOT_SUPPORTED;
    status = IoCallDriver(target, irp);
    if (status == STATUS_PENDING)
    {
        KeWaitForSingleObject(&event, Executive, KernelMode, FALSE, NULL);
        status = ioStatus.Status;
    }
    reply->configStatus = status;
    if (NT_SUCCESS(status))
    {
        if (bus.Size < sizeof(bus) || bus.Version != 1 || bus.GetBusData == NULL
            || bus.InterfaceDereference == NULL)
            reply->configStatus = STATUS_NOT_SUPPORTED;
        else
        {
            reply->configBytes = bus.GetBusData(bus.Context, PCI_WHICHSPACE_CONFIG,
                reply->config, 0, sizeof(reply->config));
            if (reply->configBytes != sizeof(reply->config)) reply->configStatus = STATUS_DEVICE_DATA_ERROR;
        }
        if (bus.InterfaceDereference != NULL) bus.InterfaceDereference(bus.Context);
    }
Success:
    status = STATUS_SUCCESS; // Per-source statuses remain in the reply.
    if (target != NULL) ObDereferenceObject(target);
    if (pdo != NULL) ObDereferenceObject(pdo);
    if (fileObject != NULL) ObDereferenceObject(fileObject);
    return status;
}

// Opens an exact PCI devnode: the PDO name selects it and the instance ID,
// read back from the PDO in kernel, must match in full. Callers own both references.
static NTSTATUS ImodOpenPciDevice(const WCHAR *pdoName, const WCHAR *instanceId,
    PFILE_OBJECT *fileObject, PDEVICE_OBJECT *pdo)
{
    UNICODE_STRING name, expectedId, actualId;
    PDEVICE_OBJECT deviceObject = NULL;
    WCHAR id[256];
    ULONG bytes = 0;
    DEVPROPTYPE type = 0;
    NTSTATUS status;

    *fileObject = NULL;
    *pdo = NULL;
    if (pdoName[255] != 0 || instanceId[255] != 0)
        return STATUS_INVALID_PARAMETER;
    RtlInitUnicodeString(&name, pdoName);
    RtlInitUnicodeString(&expectedId, instanceId);
    if (name.Length < 16 || expectedId.Length < 8
        || _wcsnicmp(pdoName, L"\\Device\\", 8) != 0
        || _wcsnicmp(instanceId, L"PCI\\", 4) != 0)
        return STATUS_INVALID_PARAMETER;

    status = IoGetDeviceObjectPointer(&name, FILE_READ_ATTRIBUTES, fileObject, &deviceObject);
    if (!NT_SUCCESS(status)) return status;
    *pdo = IoGetDeviceAttachmentBaseRef(deviceObject);
    status = IoGetDevicePropertyData(*pdo, &DEVPKEY_Device_InstanceId, LOCALE_NEUTRAL, 0,
        sizeof(id), id, &bytes, &type);
    if (NT_SUCCESS(status)
        && (type != DEVPROP_TYPE_STRING || bytes < sizeof(WCHAR) || bytes > sizeof(id)
            || (bytes % sizeof(WCHAR)) != 0 || id[bytes / sizeof(WCHAR) - 1] != 0))
        status = STATUS_DATA_ERROR;
    if (NT_SUCCESS(status))
    {
        RtlInitUnicodeString(&actualId, id);
        if (!RtlEqualUnicodeString(&actualId, &expectedId, TRUE))
            status = STATUS_OBJECT_TYPE_MISMATCH;
    }

    if (!NT_SUCCESS(status))
    {
        ObDereferenceObject(*pdo);
        ObDereferenceObject(*fileObject);
        *pdo = NULL;
        *fileObject = NULL;
    }

    return status;
}

// Walks a CM_RESOURCE_LIST byte by byte and reports whether [address, address + size)
// lies inside one memory range. Every count is checked against the buffer, so a
// malformed list allows nothing.
static BOOLEAN ImodResourcesContainRange(const UCHAR *list, ULONG bytes, ULONGLONG address, ULONG size)
{
    ULONG fullCount, offset, full, partial;

    if (list == NULL || size == 0 || bytes < sizeof(ULONG) || address > MAXULONGLONG - size)
        return FALSE;
    RtlCopyMemory(&fullCount, list, sizeof(ULONG));
    offset = sizeof(ULONG);
    for (full = 0; full < fullCount; full++)
    {
        ULONG partialCount;

        // InterfaceType, BusNumber, Version, Revision, Count.
        if (bytes - offset < 16) return FALSE;
        RtlCopyMemory(&partialCount, list + offset + 12, sizeof(ULONG));
        offset += 16;
        if (partialCount > (bytes - offset) / 20) return FALSE;
        for (partial = 0; partial < partialCount; partial++, offset += 20)
        {
            const UCHAR *descriptor = list + offset;
            USHORT flags;
            ULONGLONG start, length;
            ULONG length32;

            RtlCopyMemory(&flags, descriptor + 2, sizeof(USHORT));
            RtlCopyMemory(&start, descriptor + 4, sizeof(ULONGLONG));
            RtlCopyMemory(&length32, descriptor + 12, sizeof(ULONG));
            if (descriptor[0] == CmResourceTypeMemory)
                length = length32;
            else if (descriptor[0] == CmResourceTypeMemoryLarge && (flags & CM_RESOURCE_MEMORY_LARGE_40))
                length = (ULONGLONG)length32 << 8;
            else if (descriptor[0] == CmResourceTypeMemoryLarge && (flags & CM_RESOURCE_MEMORY_LARGE_48))
                length = (ULONGLONG)length32 << 16;
            else if (descriptor[0] == CmResourceTypeMemoryLarge && (flags & CM_RESOURCE_MEMORY_LARGE_64))
                length = (ULONGLONG)length32 << 32;
            else
                continue;

            if (length >= size && address >= start && address - start <= length - size)
                return TRUE;
        }
    }

    return FALSE;
}

static VOID ImodWriteDeviceWorker(PVOID IoObject, PVOID Context, PIO_WORKITEM WorkItem)
{
    PIRP irp = (PIRP)Context;
    struct ImodDeviceWrite request;
    struct ImodDeviceWriteReply *reply = (struct ImodDeviceWriteReply *)irp->AssociatedIrp.SystemBuffer;
    NTSTATUS status;
    UNREFERENCED_PARAMETER(IoObject);
    // METHOD_BUFFERED shares input/output; retain the request before clearing the reply.
    RtlCopyMemory(&request, reply, sizeof(request));
    RtlZeroMemory(reply, sizeof(*reply));
    status = irp->Cancel ? STATUS_CANCELLED : ImodWriteDevice(&request, reply);
    irp->IoStatus.Status = status;
    irp->IoStatus.Information = NT_SUCCESS(status)
        ? FIELD_OFFSET(struct ImodDeviceWriteReply, resources) + reply->resourcesBytes : 0;
    IoFreeWorkItem(WorkItem);
    IoCompleteRequest(irp, IO_NO_INCREMENT);
}

static NTSTATUS ImodWriteDevice(struct ImodDeviceWrite *request, struct ImodDeviceWriteReply *reply)
{
    PFILE_OBJECT fileObject = NULL;
    PDEVICE_OBJECT pdo = NULL;
    PHYSICAL_ADDRESS address;
    ULONG bytes = 0;
    NTSTATUS status;

    if (KeGetCurrentIrql() != PASSIVE_LEVEL || request->version != IMOD_DEVICE_WRITE_VERSION
        || !ImodIsValidAccessSize(request->accessSize))
        return STATUS_INVALID_PARAMETER;
    status = ImodOpenPciDevice(request->pdoName, request->instanceId, &fileObject, &pdo);
    if (!NT_SUCCESS(status)) return status;

    // The device stays referenced, so its resources cannot change during the write.
    reply->version = IMOD_DEVICE_WRITE_VERSION;
    status = IoGetDeviceProperty(pdo, DevicePropertyAllocatedResources,
        sizeof(reply->resources), reply->resources, &bytes);
    if (NT_SUCCESS(status) && bytes > sizeof(reply->resources))
        status = STATUS_DATA_ERROR;
    if (NT_SUCCESS(status))
    {
        reply->resourcesBytes = bytes;
        if (!ImodResourcesContainRange(reply->resources, bytes, request->physAddress, request->accessSize))
        {
            KdPrint(("IMOD: write outside device memory rejected"));
            status = STATUS_ACCESS_DENIED;
        }
        else
        {
            address.QuadPart = (LONGLONG)request->physAddress;
            status = ImodWritePhysicalMemory(address, request->accessSize, request->value);
        }
    }

    ObDereferenceObject(pdo);
    ObDereferenceObject(fileObject);
    return status;
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
