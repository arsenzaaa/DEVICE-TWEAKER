/* User-mode diagnostic hooks only. These events are injected assumptions, not
 * an implementation of the Windows PnP manager or a production driver guard.
 * Included by Test-NativeDriverModel.ps1 only with -ProbeDeviceLifecycle. */
enum LifecyclePoint {
    LifecycleAfterQuery = 1,
    LifecycleBeforeMap,
    LifecycleAfterMap,
    LifecycleAfterAccess
};
enum LifecycleChange { LifecycleRemoved = 1, LifecycleReassigned, LifecycleStopped };
static int lifecycleEnabled, lifecyclePoint, lifecycleChange, lifecycleFired;
static int lifecycleInvalid, lifecycleStaleReads, lifecycleStaleWrites;
static int lifecycleStaleMaps, lifecycleOwnerWrites, lifecycleSequence;
static int lifecycleEventSequence, lifecycleAccessSequence;

static void LifecycleReset(void)
{
    lifecycleEnabled = lifecyclePoint = lifecycleChange = lifecycleFired = 0;
    lifecycleInvalid = lifecycleStaleReads = lifecycleStaleWrites = 0;
    lifecycleStaleMaps = lifecycleOwnerWrites = lifecycleSequence = 0;
    lifecycleEventSequence = lifecycleAccessSequence = 0;
}

static void LifecycleEvent(int point)
{
    if (!lifecycleEnabled || lifecycleFired || lifecyclePoint != point) return;
#ifndef LIFECYCLE_DROP_EVENT
    lifecycleFired = lifecycleInvalid = 1;
    lifecycleEventSequence = ++lifecycleSequence;
    /* One physical-address buffer, whose modeled owner changes. A mapping
     * made before the change still points to the same bytes afterwards. */
    if (lifecycleChange == LifecycleReassigned) memset(memory, 0xA5, sizeof(memory));
#endif
}

static void LifecycleBeforeAccess(int writing)
{
    if (!lifecycleEnabled) return;
    lifecycleAccessSequence = ++lifecycleSequence;
    if (!lifecycleInvalid) return;
#ifndef LIFECYCLE_HIDE_ACCESS
    if (writing) lifecycleStaleWrites++;
    else lifecycleStaleReads++;
#endif
    if (writing && lifecycleChange == LifecycleReassigned) lifecycleOwnerWrites++;
}

static void *LifecycleMap(void)
{
    LifecycleEvent(LifecycleBeforeMap);
    mapCalls++;
    if (lifecycleEnabled && lifecycleInvalid) lifecycleStaleMaps++;
    if (failMap) return NULL;
    LifecycleEvent(LifecycleAfterMap);
    return memory;
}

#define LIFECYCLE_REGISTER_FUNCTIONS(type, suffix) \
static type LifecycleRead##suffix(volatile type *p) { \
    type value; LifecycleBeforeAccess(0); reads++; value = *p; \
    LifecycleEvent(LifecycleAfterAccess); return value; } \
static void LifecycleWrite##suffix(volatile type *p, type value) { \
    LifecycleBeforeAccess(1); writes++; *p = value; LifecycleEvent(LifecycleAfterAccess); }
LIFECYCLE_REGISTER_FUNCTIONS(UCHAR, Uchar)
LIFECYCLE_REGISTER_FUNCTIONS(USHORT, Ushort)
LIFECYCLE_REGISTER_FUNCTIONS(ULONG, Ulong)
