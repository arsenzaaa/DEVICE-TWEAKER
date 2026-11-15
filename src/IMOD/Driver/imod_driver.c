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
static NTSTATUS ImodReadPciConfig(PDEVICE_OBJECT pdo, UCHAR *config);
static NTSTATUS ImodReadPciClass(PDEVICE_OBJECT pdo, ULONG *classCode);
static BOOLEAN ImodIsAllowedClass(ULONG classCode);
static NTSTATUS ImodReadDevice(struct ImodDeviceWrite *request, struct ImodDeviceReadReply *reply);
static NTSTATUS ImodQueryXhci(struct ImodXhciQuery *request, struct ImodXhciReply *reply);
static IO_WORKITEM_ROUTINE_EX ImodReadDeviceWorker;
static IO_WORKITEM_ROUTINE_EX ImodQueryXhciWorker;
C_ASSERT(sizeof(struct ImodDeviceWrite) == 1048);
C_ASSERT(sizeof(struct ImodXhciQuery) == 1048);
C_ASSERT(sizeof(struct ImodDeviceReadReply) == 16);
C_ASSERT(FIELD_OFFSET(struct ImodXhciReply, slots) == 28);
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

        case IOCTL_IMOD_READ_DEVICE:
        case IOCTL_IMOD_QUERY_XHCI:
            if (ioBuffer == NULL
                || inputLength != sizeof(struct ImodDeviceWrite)
                || outputLength < (ioControlCode == IOCTL_IMOD_READ_DEVICE
                    ? sizeof(struct ImodDeviceReadReply) : sizeof(struct ImodXhciReply)))
            {
                status = STATUS_INVALID_PARAMETER;
                break;
            }
            {
                PIO_WORKITEM workItem = IoAllocateWorkItem(DeviceObject);
                if (workItem == NULL) { status = STATUS_INSUFFICIENT_RESOURCES; break; }
                IoMarkIrpPending(Irp);
                IoQueueWorkItemEx(workItem, ioControlCode == IOCTL_IMOD_READ_DEVICE
                    ? ImodReadDeviceWorker : ImodQueryXhciWorker, DelayedWorkQueue, Irp);
                return STATUS_PENDING;
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

static NTSTATUS ImodReadPciConfig(PDEVICE_OBJECT pdo, UCHAR *config)
{
    PDEVICE_OBJECT target = NULL;
    BUS_INTERFACE_STANDARD bus;
    NTSTATUS status;
    ULONG configBytes = 0;
    KEVENT event;
    IO_STATUS_BLOCK ioStatus;
    PIRP irp;
    PIO_STACK_LOCATION stack;
    RtlZeroMemory(&bus, sizeof(bus));
    KeInitializeEvent(&event, NotificationEvent, FALSE);
    target = IoGetAttachedDeviceReference(pdo);
    irp = IoBuildSynchronousFsdRequest(IRP_MJ_PNP, target, NULL, 0, NULL, &event, &ioStatus);
    if (irp == NULL) { status = STATUS_INSUFFICIENT_RESOURCES; goto Done; }
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
    if (NT_SUCCESS(status))
    {
        if (bus.Size < sizeof(bus) || bus.Version != 1 || bus.GetBusData == NULL
            || bus.InterfaceDereference == NULL)
            status = STATUS_NOT_SUPPORTED;
        else
        {
            configBytes = bus.GetBusData(bus.Context, PCI_WHICHSPACE_CONFIG,
                config, 0, IMOD_PCI_CONFIG_SIZE);
            if (configBytes != IMOD_PCI_CONFIG_SIZE) status = STATUS_DEVICE_DATA_ERROR;
        }
        if (bus.InterfaceDereference != NULL) bus.InterfaceDereference(bus.Context);
    }
Done:
    if (target != NULL) ObDereferenceObject(target);
    return status;
}

static NTSTATUS ImodQueryPci(struct ImodPciQuery *query, struct ImodPciReply *reply)
{
    PFILE_OBJECT fileObject = NULL;
    PDEVICE_OBJECT pdo = NULL;
    ULONG bytes = 0;
    NTSTATUS status;
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
    reply->configStatus = ImodReadPciConfig(pdo, reply->config);
    if (NT_SUCCESS(reply->configStatus)) reply->configBytes = IMOD_PCI_CONFIG_SIZE;
    ObDereferenceObject(pdo);
    ObDereferenceObject(fileObject);
    return STATUS_SUCCESS;
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
    BOOLEAN matched = FALSE;

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

            if (descriptor[0] == CmResourceTypeDeviceSpecific)
            {
                ULONG dataSize;
                RtlCopyMemory(&dataSize, descriptor + 4, sizeof(ULONG));
                // A device-specific payload follows the last partial descriptor.
                if (partial + 1 != partialCount || dataSize > bytes - offset - 20) return FALSE;
                offset += dataSize;
                continue;
            }
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
                matched = TRUE;
        }
    }

    return matched;
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

    // References keep the objects alive, not their PnP resources. This standalone
    // driver does not synchronize surprise removal/power transitions with MMIO.
    {
        ULONG classCode = 0;
        status = ImodReadPciClass(pdo, &classCode);
        if (NT_SUCCESS(status) && !ImodIsAllowedClass(classCode)) status = STATUS_ACCESS_DENIED;
        if (!NT_SUCCESS(status))
        {
            ObDereferenceObject(pdo);
            ObDereferenceObject(fileObject);
            return status;
        }
    }
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

// PCI class is read from the bus, not from a caller-supplied hardware ID.
static NTSTATUS ImodReadPciClass(PDEVICE_OBJECT pdo, ULONG *classCode)
{
    UCHAR config[IMOD_PCI_CONFIG_SIZE];
    NTSTATUS status = ImodReadPciConfig(pdo, config);
    if (!NT_SUCCESS(status)) return status;
    if ((config[0] == 0xFF && config[1] == 0xFF) || (config[0] == 0 && config[1] == 0))
        return STATUS_DEVICE_NOT_CONNECTED;
    *classCode = ((ULONG)config[11] << 16) | ((ULONG)config[10] << 8) | config[9];
    return STATUS_SUCCESS;
}

static BOOLEAN ImodIsAllowedClass(ULONG classCode)
{
    return classCode == 0x0C0330 || (classCode >> 8) == 0x0200;
}

static NTSTATUS ImodOpenMemoryDevice(const WCHAR *name, const WCHAR *id,
    PFILE_OBJECT *file, PDEVICE_OBJECT *pdo, UCHAR **resources, ULONG *bytes, ULONG *classCode)
{
    NTSTATUS status;
    *resources = NULL;
    status = ImodOpenPciDevice(name, id, file, pdo);
    if (!NT_SUCCESS(status)) return status;
    status = ImodReadPciClass(*pdo, classCode);
    if (NT_SUCCESS(status) && !ImodIsAllowedClass(*classCode)) status = STATUS_ACCESS_DENIED;
    if (NT_SUCCESS(status))
    {
        *resources = ExAllocatePool2(POOL_FLAG_PAGED, IMOD_PCI_RESOURCE_SIZE, 'rImD');
        if (*resources == NULL) status = STATUS_INSUFFICIENT_RESOURCES;
    }
    if (NT_SUCCESS(status))
    {
        status = IoGetDeviceProperty(*pdo, DevicePropertyAllocatedResources,
            IMOD_PCI_RESOURCE_SIZE, *resources, bytes);
        if (NT_SUCCESS(status) && *bytes > IMOD_PCI_RESOURCE_SIZE) status = STATUS_DATA_ERROR;
    }
    if (!NT_SUCCESS(status))
    {
        if (*resources != NULL) ExFreePoolWithTag(*resources, 'rImD');
        *resources = NULL;
        ObDereferenceObject(*pdo);
        ObDereferenceObject(*file);
        *pdo = NULL;
        *file = NULL;
    }
    return status;
}

static VOID ImodReadDeviceWorker(PVOID IoObject, PVOID Context, PIO_WORKITEM WorkItem)
{
    PIRP irp = Context;
    struct ImodDeviceWrite request;
    struct ImodDeviceReadReply *reply = irp->AssociatedIrp.SystemBuffer;
    UNREFERENCED_PARAMETER(IoObject);
    RtlCopyMemory(&request, reply, sizeof(request));
    RtlZeroMemory(reply, sizeof(*reply));
    irp->IoStatus.Status = irp->Cancel ? STATUS_CANCELLED : ImodReadDevice(&request, reply);
    irp->IoStatus.Information = NT_SUCCESS(irp->IoStatus.Status) ? sizeof(*reply) : 0;
    IoFreeWorkItem(WorkItem);
    IoCompleteRequest(irp, IO_NO_INCREMENT);
}

static NTSTATUS ImodReadDevice(struct ImodDeviceWrite *request, struct ImodDeviceReadReply *reply)
{
    PFILE_OBJECT file = NULL;
    PDEVICE_OBJECT pdo = NULL;
    UCHAR *resources = NULL;
    ULONG bytes = 0, classCode = 0;
    PHYSICAL_ADDRESS address;
    NTSTATUS status;
    if (KeGetCurrentIrql() != PASSIVE_LEVEL || request->version != IMOD_DEVICE_WRITE_VERSION
        || !ImodIsValidAccessSize(request->accessSize)) return STATUS_INVALID_PARAMETER;
    status = ImodOpenMemoryDevice(request->pdoName, request->instanceId, &file, &pdo,
        &resources, &bytes, &classCode);
    if (!NT_SUCCESS(status)) return status;
    if (!ImodResourcesContainRange(resources, bytes, request->physAddress, request->accessSize))
        status = STATUS_ACCESS_DENIED;
    else
    {
        address.QuadPart = (LONGLONG)request->physAddress;
        status = ImodReadPhysicalMemory(address, request->accessSize, &reply->value);
        if (NT_SUCCESS(status)) reply->version = IMOD_DEVICE_WRITE_VERSION;
    }
    ExFreePoolWithTag(resources, 'rImD');
    ObDereferenceObject(pdo);
    ObDereferenceObject(file);
    return status;
}

// RAM is never mapped as device memory. Only the controller's bounded context
// walk uses this internal helper; no IOCTL accepts a RAM address from user mode.
static NTSTATUS ImodCopyContext(ULONGLONG address, PVOID target, SIZE_T size)
{
    MM_COPY_ADDRESS source;
    SIZE_T copied = 0;
    NTSTATUS status;
    if (address == 0 || address > MAXLONGLONG || size > (ULONGLONG)MAXLONGLONG - address)
        return STATUS_INVALID_PARAMETER;
    source.PhysicalAddress.QuadPart = (LONGLONG)address;
    status = MmCopyMemory(target, source, size, MM_COPY_MEMORY_PHYSICAL, &copied);
    return NT_SUCCESS(status) && copied != size ? STATUS_PARTIAL_COPY : status;
}

static NTSTATUS ImodReadController32(const UCHAR *resources, ULONG bytes,
    ULONGLONG base, ULONGLONG length, ULONGLONG offset, ULONG *value)
{
    ULONGLONG raw = 0;
    PHYSICAL_ADDRESS address;
    NTSTATUS status;
    if (offset > length || sizeof(ULONG) > length - offset || base > MAXULONGLONG - offset
        || !ImodResourcesContainRange(resources, bytes, base + offset, sizeof(ULONG)))
        return STATUS_ACCESS_DENIED;
    address.QuadPart = (LONGLONG)(base + offset);
    status = ImodReadPhysicalMemory(address, sizeof(ULONG), &raw);
    *value = (ULONG)raw;
    return status;
}

static VOID ImodQueryXhciWorker(PVOID IoObject, PVOID Context, PIO_WORKITEM WorkItem)
{
    PIRP irp = Context;
    struct ImodXhciQuery request;
    struct ImodXhciReply *reply = irp->AssociatedIrp.SystemBuffer;
    UNREFERENCED_PARAMETER(IoObject);
    RtlCopyMemory(&request, reply, sizeof(request));
    RtlZeroMemory(reply, sizeof(*reply));
    irp->IoStatus.Status = irp->Cancel ? STATUS_CANCELLED : ImodQueryXhci(&request, reply);
    irp->IoStatus.Information = NT_SUCCESS(irp->IoStatus.Status)
        ? FIELD_OFFSET(struct ImodXhciReply, slots) + reply->count * sizeof(struct ImodXhciSlot) : 0;
    IoFreeWorkItem(WorkItem);
    IoCompleteRequest(irp, IO_NO_INCREMENT);
}

// Returns only root-port/address/interrupter tuples, never raw context memory.
// The walk has the same slot/endpoint/TRB selection rules as the managed model.
static NTSTATUS ImodQueryXhci(struct ImodXhciQuery *request, struct ImodXhciReply *reply)
{
    PFILE_OBJECT file = NULL;
    PDEVICE_OBJECT pdo = NULL;
    UCHAR *resources = NULL;
    ULONG bytes = 0, classCode = 0, cap = 0, hcs = 0, hcc = 0, low = 0, high = 0, slot;
    ULONGLONG dcbaa;
    NTSTATUS status;
    if (KeGetCurrentIrql() != PASSIVE_LEVEL || request->version != IMOD_XHCI_QUERY_VERSION
        || request->capabilityAddress == 0 || request->memoryLength < 0x40) return STATUS_INVALID_PARAMETER;
    status = ImodOpenMemoryDevice(request->pdoName, request->instanceId, &file, &pdo,
        &resources, &bytes, &classCode);
    if (!NT_SUCCESS(status)) return status;
    if (classCode != 0x0C0330) { status = STATUS_ACCESS_DENIED; goto Done; }
#define READ_CAP(off, val) ImodReadController32(resources, bytes, request->capabilityAddress, request->memoryLength, (off), &(val))
    status = READ_CAP(0, cap);
    if (!NT_SUCCESS(status)) goto Done;
    if ((cap & 0xFF) < 0x20 || (cap >> 16) < 0x0090 || (cap >> 16) > 0x0120)
        { status = STATUS_DEVICE_DATA_ERROR; goto Done; }
    status = READ_CAP(request->hcsparamsOffset, hcs);
    if (!NT_SUCCESS(status)) goto Done;
    status = READ_CAP(0x10, hcc);
    if (!NT_SUCCESS(status)) goto Done;
    reply->maxSlots = hcs & 0xFF;
    reply->maxIntrs = (hcs >> 8) & 0x7FF;
    reply->contextSize = (hcc & 4) ? 64 : 32;
    if (reply->maxSlots == 0 || reply->maxIntrs == 0)
        { status = STATUS_DEVICE_DATA_ERROR; goto Done; }
    status = READ_CAP((cap & 0xFF) + 0x30, low);
    if (!NT_SUCCESS(status)) goto Done;
    status = READ_CAP((cap & 0xFF) + 0x34, high);
    if (!NT_SUCCESS(status)) goto Done;
    dcbaa = (((ULONGLONG)high << 32) | low) & ~63ULL;
    if (dcbaa == 0) { status = STATUS_DEVICE_DATA_ERROR; goto Done; }
    for (slot = 1; slot <= reply->maxSlots; slot++)
    {
        ULONGLONG context = 0;
        ULONG words[4], entries, state, port, deviceAddress, target, endpoint;
        BOOLEAN endpointFound = FALSE;
        if (dcbaa > MAXULONGLONG - (ULONGLONG)slot * 8
            || !NT_SUCCESS(ImodCopyContext(dcbaa + (ULONGLONG)slot * 8, &context, sizeof(context)))) continue;
        context &= ~63ULL;
        if (!NT_SUCCESS(ImodCopyContext(context, words, sizeof(words)))) continue;
        entries = (words[0] >> 27) & 31;
        state = (words[3] >> 27) & 31;
        port = (words[1] >> 16) & 255;
        deviceAddress = words[3] & 255;
        target = (words[2] >> 22) & 1023;
        if (state < 2 || (words[0] & (1UL << 26)) || port == 0) continue;
        for (endpoint = 2; endpoint <= entries; endpoint++)
        {
            ULONG ep[4], trb[4], type, candidate;
            ULONGLONG ring, offset = (ULONGLONG)endpoint * reply->contextSize;
            if (context > MAXULONGLONG - offset
                || !NT_SUCCESS(ImodCopyContext(context + offset, ep, sizeof(ep)))) continue;
            type = (ep[1] >> 3) & 7;
            if ((ep[0] & 7) == 0 || (type != 3 && type != 5 && type != 7)) continue;
            ring = (((ULONGLONG)ep[3] << 32) | ep[2]) & ~15ULL;
            if (!NT_SUCCESS(ImodCopyContext(ring, trb, sizeof(trb)))) continue;
            type = (trb[3] >> 10) & 63;
            candidate = (trb[2] >> 22) & 1023;
            if ((type == 1 || type == 3 || type == 5) && candidate < reply->maxIntrs)
                { target = candidate; endpointFound = TRUE; break; }
        }
        if (endpointFound) reply->endpointTargets++; else reply->slotTargets++;
        if (target < reply->maxIntrs)
        {
            struct ImodXhciSlot *entry = &reply->slots[reply->count++];
            entry->rootPort = port;
            entry->deviceAddress = deviceAddress;
            entry->interrupter = target;
        }
    }
    reply->version = IMOD_XHCI_QUERY_VERSION;
    status = STATUS_SUCCESS;
Done:
#undef READ_CAP
    ExFreePoolWithTag(resources, 'rImD');
    ObDereferenceObject(pdo);
    ObDereferenceObject(file);
    return status;
}
