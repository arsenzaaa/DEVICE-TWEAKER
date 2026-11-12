/* Executed against extracted production functions and user-mode fake APIs.
 * Successful access to a permissive fake mapping does NOT predict the outcome
 * of accessing physically absent hardware. The observation is the access attempt. */
static void LifecycleResources(ULONGLONG start)
{
    ULONG n = put32(assignedResources, 0, 1);
    n = putFull(assignedResources, n, 1);
    assignedBytes = putPartial(assignedResources, n, CmResourceTypeMemory, 0, start, 0x1000);
}

static NTSTATUS LifecycleReadRequest(ULONGLONG address)
{
    IRP irp = {0};
    struct tagPhysAccessStruct access = {0};
    NTSTATUS status;
    access.physAddress = address; access.accessSizeInBytes = 4;
    irp.stack.MajorFunction = IRP_MJ_DEVICE_CONTROL;
    irp.AssociatedIrp.SystemBuffer = &access;
    irp.stack.Parameters.DeviceIoControl.InputBufferLength = sizeof(access);
    irp.stack.Parameters.DeviceIoControl.OutputBufferLength = sizeof(access);
    irp.stack.Parameters.DeviceIoControl.IoControlCode = IOCTL_IMOD_READ_PHYSICAL;
    status = ImodDriverDispatch(&device, &irp);
    assert(completions == 1);
    if (NT_SUCCESS(status)) {
        assert(irp.IoStatus.Information == sizeof(access));
        if (lifecycleInvalid && lifecycleChange == LifecycleReassigned
            && lifecyclePoint != LifecycleAfterAccess) assert(access.value == 0xA5A5A5A5ull);
    } else assert(irp.IoStatus.Information == 0);
    return status;
}

static void LifecycleReport(const char *name, NTSTATUS status, int expectedAccesses,
    int expectedInvalidAccesses, int expectedEvent)
{
    int invalidAccesses = lifecycleStaleReads + lifecycleStaleWrites;
    assert(reads + writes == expectedAccesses);
    assert(invalidAccesses == expectedInvalidAccesses && lifecycleFired == expectedEvent);
    assert(unmapCalls == mapCalls - (failMap ? 1 : 0));
    assert(dereferences == (openCalls && !failOpen ? 2 : 0));
    if (invalidAccesses) assert(lifecycleEventSequence < lifecycleAccessSequence);
    if (lifecycleOwnerWrites) {
        assert(memory[4] == 0x78 && memory[7] == 0x12 && memory[8] == 0xA5);
    }
    printf("LIFECYCLE_CASE={\"name\":\"%s\",\"operationStatus\":%d,"
        "\"reads\":%d,\"writes\":%d,\"maps\":%d,\"unmaps\":%d,"
        "\"dereferences\":%d,\"injectedEvent\":%d,\"accessesAfterInvalidation\":%d,"
        "\"mapsAfterInvalidation\":%d,\"newOwnerWrites\":%d,"
        "\"eventSequence\":%d,\"accessSequence\":%d}\n",
        name, status, reads, writes, mapCalls, unmapCalls, dereferences, lifecycleFired,
        invalidAccesses, lifecycleStaleMaps, lifecycleOwnerWrites,
        lifecycleEventSequence, lifecycleAccessSequence);
}

static void LifecycleBegin(void)
{
    reset(); LifecycleResources(0x1000); lifecycleEnabled = 1;
}

static void deviceLifecycleTests(void)
{
    NTSTATUS status;
    int point, change;
    char name[96];
    const char *changes[] = { "", "removed", "reassigned", "stopped" };
    const char *points[] = { "", "after-query", "before-map", "after-map", "after-access" };

    LifecycleBegin(); status = dispatchDeviceWrite(0x1004, 4, IMOD_DEVICE_WRITE_VERSION, 0);
    assert(status == STATUS_SUCCESS); LifecycleReport("write-steady", status, 1, 0, 0);
    LifecycleBegin(); status = LifecycleReadRequest(0x1004);
    assert(status == STATUS_SUCCESS); LifecycleReport("read-steady", status, 1, 0, 0);
    LifecycleBegin(); failOpen = 1;
    status = dispatchDeviceWrite(0x1004, 4, IMOD_DEVICE_WRITE_VERSION, 0);
    assert(!NT_SUCCESS(status) && propertyCalls == 0);
    LifecycleReport("write-open-failed", status, 0, 0, 0);
    LifecycleBegin(); failProperty = 1;
    status = dispatchDeviceWrite(0x1004, 4, IMOD_DEVICE_WRITE_VERSION, 0);
    assert(status == STATUS_DATA_ERROR); LifecycleReport("write-query-failed", status, 0, 0, 0);
    LifecycleBegin(); failMap = 1;
    status = dispatchDeviceWrite(0x1004, 4, IMOD_DEVICE_WRITE_VERSION, 0);
    assert(status == STATUS_INSUFFICIENT_RESOURCES); LifecycleReport("write-map-failed", status, 0, 0, 0);
    LifecycleBegin(); failMap = 1; status = LifecycleReadRequest(0x1004);
    assert(status == STATUS_INSUFFICIENT_RESOURCES); LifecycleReport("read-map-failed", status, 0, 0, 0);
    LifecycleBegin(); status = dispatchDeviceWrite(0x3000, 4, IMOD_DEVICE_WRITE_VERSION, 0);
    assert(status == STATUS_ACCESS_DENIED); LifecycleReport("write-outside-range", status, 0, 0, 0);
    LifecycleBegin(); status = dispatchDeviceWrite(0x1004, 4, IMOD_DEVICE_WRITE_VERSION, 1);
    assert(status == STATUS_CANCELLED); LifecycleReport("write-cancelled", status, 0, 0, 0);

    /* Resource query on a later request sees the new range. This checks the
     * actual range guard; it does not invent stop/start handling in DTIMOD. */
    LifecycleBegin(); LifecycleResources(0x2000);
    status = dispatchDeviceWrite(0x1004, 4, IMOD_DEVICE_WRITE_VERSION, 0);
    assert(status == STATUS_ACCESS_DENIED); LifecycleReport("write-old-range-rejected", status, 0, 0, 0);
    LifecycleBegin(); LifecycleResources(0x2000);
    status = dispatchDeviceWrite(0x2004, 4, IMOD_DEVICE_WRITE_VERSION, 0);
    assert(status == STATUS_SUCCESS); LifecycleReport("write-new-range-accepted", status, 1, 0, 0);

    for (change = LifecycleRemoved; change <= LifecycleStopped; change++) {
        for (point = LifecycleAfterQuery; point <= LifecycleAfterMap; point++) {
            LifecycleBegin(); lifecyclePoint = point; lifecycleChange = change;
            status = dispatchDeviceWrite(0x1004, 4, IMOD_DEVICE_WRITE_VERSION, 0);
            sprintf_s(name, sizeof(name), "write-%s-%s", changes[change], points[point]);
            assert(status == STATUS_SUCCESS); LifecycleReport(name, status, 1, 1, 1);
        }
        /* The raw physical-read IOCTL does not query a device's resources. */
        for (point = LifecycleBeforeMap; point <= LifecycleAfterMap; point++) {
            LifecycleBegin(); lifecyclePoint = point; lifecycleChange = change;
            status = LifecycleReadRequest(0x1004);
            sprintf_s(name, sizeof(name), "read-%s-%s", changes[change], points[point]);
            assert(status == STATUS_SUCCESS && openCalls == 0 && propertyCalls == 0);
            LifecycleReport(name, status, 1, 1, 1);
        }
    }
    /* Negative timing controls: a change after access cannot make that earlier
     * access stale. They catch a detector that flags every injected event. */
    LifecycleBegin(); lifecyclePoint = LifecycleAfterAccess; lifecycleChange = LifecycleRemoved;
    status = dispatchDeviceWrite(0x1004, 4, IMOD_DEVICE_WRITE_VERSION, 0);
    assert(status == STATUS_SUCCESS && lifecycleAccessSequence < lifecycleEventSequence);
    LifecycleReport("write-removed-after-access", status, 1, 0, 1);
    LifecycleBegin(); lifecyclePoint = LifecycleAfterAccess; lifecycleChange = LifecycleStopped;
    status = LifecycleReadRequest(0x1004);
    assert(status == STATUS_SUCCESS && lifecycleAccessSequence < lifecycleEventSequence);
    LifecycleReport("read-stopped-after-access", status, 1, 0, 1);
    LifecycleReset();
}
