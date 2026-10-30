# Compile the actual DTIMOD entry/dispatch/MMIO functions against user-mode fake kernel APIs.
# No SYS is executed, no driver/service is loaded, no physical memory is accessed.
[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$OutputDirectory,
    [ValidateSet('unload', 'write-past-device-end')][string]$InjectRegression)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$driver = Join-Path $root 'src\IMOD\Driver\imod_driver.c'
$header = Join-Path $root 'src\IMOD\Driver\imod_driver.h'
$text = Get-Content -LiteralPath $driver -Raw
$functions = @('ImodDriverUnload','DriverEntry','ImodDriverCreateDriver','ImodDriverInitialize','ImodDriverDispatch',
    'ImodIsValidAccessSize','ImodMapPhysicalMemory','ImodUnmapPhysicalMemory','ImodReadPhysicalMemory','ImodWritePhysicalMemory',
    'ImodResourcesContainRange')
$bodies = foreach ($name in $functions) {
    $match = [regex]::Match($text, '(?m)^(?:static )?(?:NTSTATUS|BOOLEAN|void) '+$name+'\([^;]+?\)\s*\{')
    if (-not $match.Success) { throw "Native function missing: $name" }
    $opening = $text.IndexOf('{',$match.Index); $end = $opening+1; $depth = 1
    while ($depth -gt 0) { if ($text[$end] -eq '{') { $depth++ }; if ($text[$end] -eq '}') { $depth-- }; $end++ }
    $text.Substring($match.Index,$end-$match.Index)
}
# Deliberate regressions in the actual driver code; the model must fail with an assertion.
$regressions = @{
    'unload' = @{ old = 'DriverObject->DriverUnload = ImodDriverUnload;'; new = 'DriverObject->DriverUnload = NULL;';
        detected = 'missing unload assignment in actual driver initialization' }
    'write-past-device-end' = @{ old = 'address - start <= length - size'; new = 'address - start <= length';
        detected = 'device write accepted although it crosses the end of the memory range' }
}
if ($InjectRegression) {
    $regression = $regressions[$InjectRegression]
    $joined = $bodies -join "`n`n"
    if ([regex]::Matches($joined, [regex]::Escape($regression.old)).Count -ne 1) { throw "Mutation anchor must occur once: $InjectRegression" }
    $bodies = @($joined.Replace($regression.old, $regression.new))
}
$stub = @'
#include <assert.h>
#include <excpt.h>
#include <stdint.h>
#include <stddef.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
#define IN
#define TRUE 1
#define FALSE 0
#define UNREFERENCED_PARAMETER(x) (void)(x)
#define KdPrint(x) ((void)0)
#define NT_SUCCESS(x) ((x)>=0)
#define STATUS_SUCCESS 0
#define STATUS_INVALID_PARAMETER (-1)
#define STATUS_INSUFFICIENT_RESOURCES (-2)
#define STATUS_INVALID_DEVICE_REQUEST (-3)
#define STATUS_PENDING 1
#define FILE_DEVICE_SECURE_OPEN 0
#define FILE_READ_ACCESS 1
#define FILE_ANY_ACCESS 0
#define METHOD_BUFFERED 0
#define CTL_CODE(d,f,m,a) (((unsigned long)(d)<<16)|((unsigned long)(a)<<14)|((unsigned long)(f)<<2)|(m))
#define DO_DEVICE_INITIALIZING 0x80
#define IRP_MJ_CREATE 0
#define IRP_MJ_CLOSE 1
#define IRP_MJ_DEVICE_CONTROL 2
#define IO_NO_INCREMENT 0
#define PAGE_SIZE 4096
#define IMOD_MAX_MAP_SIZE PAGE_SIZE
#define MAXLONGLONG INT64_MAX
#define MAXULONGLONG UINT64_MAX
#define FILE_WRITE_ACCESS 2
#define CmResourceTypeMemory 3
#define CmResourceTypeMemoryLarge 7
#define CM_RESOURCE_MEMORY_LARGE_40 0x0200
#define CM_RESOURCE_MEMORY_LARGE_48 0x0400
#define CM_RESOURCE_MEMORY_LARGE_64 0x0800
#define EXCEPTION_EXECUTE_HANDLER 1
#define MmNonCached 0
#define DelayedWorkQueue 0
#define IMOD_MAPPED_DRIVER_NAME L"\\Driver\\DeviceTweakerImodMapped2"
#define IMOD_DEVICE_NAME L"\\Device\\DeviceTweakerImod2"
#define IMOD_DOS_DEVICE_NAME L"\\DosDevices\\DeviceTweakerImod2"
typedef int NTSTATUS, BOOLEAN;
typedef unsigned long ULONG;
typedef unsigned short USHORT, WCHAR;
typedef unsigned char UCHAR, *PUCHAR;
typedef uint64_t ULONGLONG;
typedef int64_t LONGLONG;
typedef uintptr_t ULONG_PTR;
typedef size_t SIZE_T;
typedef void *PVOID;
typedef struct { LONGLONG QuadPart; } PHYSICAL_ADDRESS;
typedef struct { const wchar_t *text; } UNICODE_STRING, *PUNICODE_STRING;
typedef struct DEVICE { unsigned Flags; } DEVICE, *PDEVICE_OBJECT;
typedef struct STACK { int MajorFunction; struct { struct { ULONG InputBufferLength, OutputBufferLength, IoControlCode; } DeviceIoControl; } Parameters; } IO_STACK_LOCATION,*PIO_STACK_LOCATION;
typedef struct IRP { struct { PVOID SystemBuffer; } AssociatedIrp; struct { NTSTATUS Status; SIZE_T Information; } IoStatus; IO_STACK_LOCATION stack; } IRP,*PIRP;
typedef struct DRIVER { PDEVICE_OBJECT DeviceObject; NTSTATUS (*MajorFunction[3])(PDEVICE_OBJECT,PIRP); void (*DriverUnload)(struct DRIVER *); } DRIVER,*PDRIVER_OBJECT;
typedef void *PIO_WORKITEM;

static int GUID_DEVCLASS_DEVICE_TWEAKER_IMOD, mapCalls, unmapCalls, reads, writes, failMap, failCreate, failLink, deleted;

static unsigned char memory[4096];
static DEVICE device;
static DRIVER created;
static void RtlInitUnicodeString(PUNICODE_STRING s,const wchar_t *t) { s->text=t; }
#define RtlCopyMemory memcpy
#define RtlZeroMemory(p,n) memset((p),0,(n))
static NTSTATUS IoCreateDeviceSecure(PDRIVER_OBJECT d,int e,PUNICODE_STRING n,int t,int o,int x,PUNICODE_STRING s,int *g,PDEVICE_OBJECT *out) {
    (void)e;(void)n;(void)t;(void)o;(void)x;(void)s;(void)g;if(failCreate)return -2;d->DeviceObject=&device;*out=&device;return 0; }
static NTSTATUS IoCreateSymbolicLink(PUNICODE_STRING a,PUNICODE_STRING b) { (void)a;(void)b;return failLink ? -2:0; }
static NTSTATUS IoDeleteSymbolicLink(PUNICODE_STRING n) { (void)n; return 0; }
static void ImodDriverUnload(PDRIVER_OBJECT d);
static void IoDeleteDevice(PDEVICE_OBJECT d) { (void)d;deleted++; }
static PIO_STACK_LOCATION IoGetCurrentIrpStackLocation(PIRP irp) { return &irp->stack; }
static void IoCompleteRequest(PIRP irp,int increment) { (void)irp;(void)increment; }
static PIO_WORKITEM IoAllocateWorkItem(PDEVICE_OBJECT d) { (void)d;return NULL; }
static void IoMarkIrpPending(PIRP i) { (void)i; }
static void ImodQueryPciWorker(void *a,void *b,void *c) { (void)a;(void)b;(void)c; }
static void ImodWriteDeviceWorker(void *a,void *b,void *c) { (void)a;(void)b;(void)c; }
static void IoQueueWorkItemEx(PIO_WORKITEM w,void (*fn)(void*,void*,void*),int q,void *c) { (void)w;(void)fn;(void)q;(void)c; }
static void *MmMapIoSpace(PHYSICAL_ADDRESS a,SIZE_T n,int cache) { (void)a;(void)cache;assert(n==4096);mapCalls++;return failMap?NULL:memory; }
static void MmUnmapIoSpace(void *p,SIZE_T n) { assert(p==memory && n==4096);unmapCalls++; }
#define READ_REGISTER_UCHAR(p) (reads++,*(p))
#define READ_REGISTER_USHORT(p) (reads++,*(p))
#define READ_REGISTER_ULONG(p) (reads++,*(p))
#define WRITE_REGISTER_UCHAR(p,v) (writes++,*(p)=(v))
#define WRITE_REGISTER_USHORT(p,v) (writes++,*(p)=(v))
#define WRITE_REGISTER_ULONG(p,v) (writes++,*(p)=(v))
static NTSTATUS ImodDriverInitialize(PDRIVER_OBJECT d);
static NTSTATUS ImodDriverCreateDriver(PDRIVER_OBJECT d,PUNICODE_STRING p);
static NTSTATUS ImodDriverDispatch(PDEVICE_OBJECT d,PIRP i);
static BOOLEAN ImodIsValidAccessSize(ULONG s);
static NTSTATUS ImodMapPhysicalMemory(PHYSICAL_ADDRESS a,SIZE_T s,PVOID *m,PVOID *b,SIZE_T *n);
static NTSTATUS ImodUnmapPhysicalMemory(PVOID b,SIZE_T s);
static NTSTATUS ImodReadPhysicalMemory(PHYSICAL_ADDRESS a,ULONG s,ULONGLONG *v);
static NTSTATUS ImodWritePhysicalMemory(PHYSICAL_ADDRESS a,ULONG s,ULONGLONG v);
static BOOLEAN ImodResourcesContainRange(const UCHAR *list,ULONG bytes,ULONGLONG address,ULONG size);
static NTSTATUS IoCreateDriver(PUNICODE_STRING n,NTSTATUS(*fn)(PDRIVER_OBJECT,PUNICODE_STRING)) { (void)n;return fn(&created,NULL); }
'@
$main = @'
static void reset(void) { mapCalls=unmapCalls=reads=writes=failMap=0; }
static void initTests(void) {

    DRIVER normal={0},mapped={0};UNICODE_STRING path={L"\\REGISTRY\\MACHINE\\SYSTEM\\CurrentControlSet\\Services\\DTIMOD"};
    assert(DriverEntry(&normal,&path)==0 && normal.DriverUnload==ImodDriverUnload);
    path.text=L"\\REGISTRY\\MACHINE\\SYSTEM\\CurrentControlSet\\Services\\DTIMODMapped";
    assert(DriverEntry(&mapped,&path)==0 && mapped.DriverUnload==ImodDriverUnload);
    assert(DriverEntry(NULL,NULL)==0 && created.DriverUnload==ImodDriverUnload);
    failCreate=1; assert(DriverEntry(&normal,&path)<0); failCreate=0;
    failLink=1; deleted=0; assert(DriverEntry(&normal,&path)<0 && deleted==1);failLink=0;
}
static void mmioTests(void) {
    PHYSICAL_ADDRESS a={0x1000}; ULONGLONG value=0; unsigned widths[]={1,2,4};
    for(unsigned i=0;i<3;i++) {
        reset();assert(ImodWritePhysicalMemory(a,widths[i],0x12345678)==0);
        assert(ImodReadPhysicalMemory(a,widths[i],&value)==0);
        assert(value==(0x12345678ull & (widths[i]==1?0xffull:widths[i]==2?0xffffull:0xffffffffull)));
        assert(mapCalls==2 && unmapCalls==2 && reads==1 && writes==1);
    }
    reset(); assert(ImodReadPhysicalMemory(a,8,&value)<0 && mapCalls==0);
    assert(ImodWritePhysicalMemory(a,3,0)<0 && mapCalls==0);
    assert(ImodReadPhysicalMemory(a,4,NULL)<0 && mapCalls==0);
    PVOID mapped=NULL,base=NULL;SIZE_T size=0;
    assert(ImodMapPhysicalMemory(a,0,&mapped,&base,&size)==STATUS_INVALID_PARAMETER);
    assert(ImodMapPhysicalMemory(a,PAGE_SIZE+1,&mapped,&base,&size)==STATUS_INVALID_PARAMETER);
    a.QuadPart=0x1000;failMap=1;assert(ImodReadPhysicalMemory(a,4,&value)==STATUS_INSUFFICIENT_RESOURCES);
    assert(reads==0 && writes==0 && unmapCalls==0);reset();
    a.QuadPart=0x1ffc; assert(ImodReadPhysicalMemory(a,4,&value)==0 && unmapCalls==1);
}
static void dispatchTests(void) {
    IRP irp={0}; struct tagPhysAccessStruct access={0x1000,4,0,0};
    assert(sizeof(access)==24 && offsetof(struct tagPhysAccessStruct,value)==16);
    irp.stack.MajorFunction=IRP_MJ_DEVICE_CONTROL;
    irp.AssociatedIrp.SystemBuffer=&access;
    irp.stack.Parameters.DeviceIoControl.InputBufferLength=sizeof(access);
    irp.stack.Parameters.DeviceIoControl.OutputBufferLength=sizeof(access);
    irp.stack.Parameters.DeviceIoControl.IoControlCode=IOCTL_IMOD_READ_PHYSICAL;
    reset();assert(ImodDriverDispatch(&device,&irp)==0 && irp.IoStatus.Information==sizeof(access) && reads==1);
    irp.stack.Parameters.DeviceIoControl.OutputBufferLength=0;
    reset();assert(ImodDriverDispatch(&device,&irp)<0 && mapCalls==0);
    irp.stack.Parameters.DeviceIoControl.IoControlCode=IOCTL_IMOD_QUERY_PCI;
    assert(ImodDriverDispatch(&device,&irp)==STATUS_INVALID_PARAMETER);
    struct ImodPciReply reply={0};irp.AssociatedIrp.SystemBuffer=&reply;
    irp.stack.Parameters.DeviceIoControl.InputBufferLength=sizeof(struct ImodPciQuery);
    irp.stack.Parameters.DeviceIoControl.OutputBufferLength=sizeof(reply);
    assert(ImodDriverDispatch(&device,&irp)==STATUS_INSUFFICIENT_RESOURCES);
    irp.stack.Parameters.DeviceIoControl.IoControlCode=0;
    assert(ImodDriverDispatch(&device,&irp)==STATUS_INVALID_DEVICE_REQUEST);
    /* The former unrestricted physical write (index + 3) no longer exists. */
    irp.AssociatedIrp.SystemBuffer=&access;
    irp.stack.Parameters.DeviceIoControl.InputBufferLength=sizeof(access);
    irp.stack.Parameters.DeviceIoControl.OutputBufferLength=sizeof(access);
    irp.stack.Parameters.DeviceIoControl.IoControlCode=CTL_CODE(FILE_DEVICE_IMOD,IMOD_IOCTL_INDEX+3,METHOD_BUFFERED,FILE_ANY_ACCESS);
    reset();assert(ImodDriverDispatch(&device,&irp)==STATUS_INVALID_DEVICE_REQUEST && mapCalls==0 && writes==0);
    /* Device writes need the full request and reply buffers, then go to a work item. */
    irp.stack.Parameters.DeviceIoControl.IoControlCode=IOCTL_IMOD_WRITE_DEVICE;
    reset();assert(ImodDriverDispatch(&device,&irp)==STATUS_INVALID_PARAMETER && mapCalls==0);
    static struct ImodDeviceWriteReply writeReply;
    irp.AssociatedIrp.SystemBuffer=&writeReply;
    irp.stack.Parameters.DeviceIoControl.InputBufferLength=sizeof(struct ImodDeviceWrite);
    irp.stack.Parameters.DeviceIoControl.OutputBufferLength=sizeof(writeReply);
    assert(sizeof(struct ImodDeviceWrite)==1048 && offsetof(struct ImodDeviceWrite,instanceId)==536);
    reset();assert(ImodDriverDispatch(&device,&irp)==STATUS_INSUFFICIENT_RESOURCES && mapCalls==0 && writes==0);
}
// CM_RESOURCE_LIST bytes: Count, then per full descriptor InterfaceType, BusNumber,
// Version, Revision, Count and 20-byte partial descriptors.
static ULONG put32(UCHAR *b,ULONG o,ULONG v) { memcpy(b+o,&v,4); return o+4; }
static ULONG putPartial(UCHAR *b,ULONG o,UCHAR type,USHORT flags,ULONGLONG start,ULONG length) {
    memset(b+o,0,20); b[o]=type; memcpy(b+o+2,&flags,2); memcpy(b+o+4,&start,8); memcpy(b+o+12,&length,4); return o+20; }
static ULONG putFull(UCHAR *b,ULONG o,ULONG partials) { o=put32(b,o,5); o=put32(b,o,0); memset(b+o,0,4); o+=4; return put32(b,o,partials); }
static void rangeTests(void) {
    UCHAR list[256]; ULONG n;
    n=put32(list,0,1); n=putFull(list,n,2);
    n=putPartial(list,n,1,0,0x3000,0x100);                       /* I/O port: never writable */
    n=putPartial(list,n,CmResourceTypeMemory,0,0xF7000000ull,0x10000);
    assert(ImodResourcesContainRange(list,n,0xF7000024ull,4));
    assert(ImodResourcesContainRange(list,n,0xF700FFFCull,4));
    assert(!ImodResourcesContainRange(list,n,0xF700FFFDull,4));   /* crosses the end */
    assert(!ImodResourcesContainRange(list,n,0xF6FFFFFCull,4));   /* before the BAR */
    assert(!ImodResourcesContainRange(list,n,0x3000,1));          /* port range */
    assert(!ImodResourcesContainRange(list,n,0xF7000024ull,0));
    assert(!ImodResourcesContainRange(list,n,UINT64_MAX-1,4));     /* address overflow */
    assert(!ImodResourcesContainRange(NULL,n,0xF7000024ull,4));
    assert(!ImodResourcesContainRange(list,n-1,0xF700FFFCull,4));  /* truncated descriptor */
    assert(!ImodResourcesContainRange(list,3,0xF7000024ull,4));
    put32(list,16,1000); assert(!ImodResourcesContainRange(list,n,0xF7000024ull,4)); /* count beyond buffer */
    n=put32(list,0,2); n=putFull(list,n,1); n=putPartial(list,n,CmResourceTypeMemory,0,0x1000,0x1000);
    n=putFull(list,n,1); n=putPartial(list,n,CmResourceTypeMemoryLarge,CM_RESOURCE_MEMORY_LARGE_64,0x4000000000ull,1);
    assert(ImodResourcesContainRange(list,n,0x40FFFFFFFCull,4));    /* second list, 4 GiB large range */
    assert(!ImodResourcesContainRange(list,n,0x4100000000ull,4));
    n=put32(list,0,1); n=putFull(list,n,1); n=putPartial(list,n,CmResourceTypeMemoryLarge,0,0x5000,0x10);
    assert(!ImodResourcesContainRange(list,n,0x5000,4));           /* large range without size flag */
}
int main(void) { _set_error_mode(_OUT_TO_STDERR);_set_abort_behavior(0,_WRITE_ABORT_MSG|_CALL_REPORTFAULT);initTests();mmioTests();rangeTests();dispatchTests();puts("PASS: actual driver entry, unload assignment, MMIO width/map/unmap, device memory range check and dispatch; fake kernel only; PCI and write workers stubbed");return 0; }
'@
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
$OutputDirectory = (Resolve-Path $OutputDirectory).Path
$source = Join-Path $OutputDirectory 'driver-model.c'
# The exact production header is included, not a parallel copy of the wire protocol.
($stub + "`n" + '#include "' + $header.Replace('\','/') + '"' + "`n" + ($bodies -join "`n`n") + "`n" + $main) | Set-Content -LiteralPath $source -Encoding UTF8
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$vs = & $vswhere -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
$tools = Get-ChildItem (Join-Path $vs 'VC\Tools\MSVC') -Directory | Sort-Object Name -Descending | Select-Object -First 1
$sdk = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10'
$sdkVersion = Get-ChildItem (Join-Path $sdk 'Include') -Directory | Where-Object { Test-Path (Join-Path $_.FullName 'ucrt') } | Sort-Object Name -Descending | Select-Object -First 1
$cl = Join-Path $tools.FullName 'bin\Hostx64\x64\cl.exe'
$exe = Join-Path $OutputDirectory 'driver-model.exe'
& $cl /nologo /W4 /WX /TC $source "/Fe:$exe" "/Fo:$(Join-Path $OutputDirectory 'driver-model.obj')" "/I$(Join-Path $tools.FullName 'include')" "/I$(Join-Path $sdkVersion.FullName 'ucrt')" /link "/LIBPATH:$(Join-Path $tools.FullName 'lib\x64')" "/LIBPATH:$(Join-Path $sdk ('Lib\'+$sdkVersion.Name+'\ucrt\x64'))" "/LIBPATH:$(Join-Path $sdk ('Lib\'+$sdkVersion.Name+'\um\x64'))" *> (Join-Path $OutputDirectory 'build.log')
if ($LASTEXITCODE -ne 0) { throw "Native model compilation failed: $OutputDirectory\build.log" }
$savedPreference = $ErrorActionPreference
try { $ErrorActionPreference = 'Continue'; & $exe *> (Join-Path $OutputDirectory 'run.log'); $runCode = $LASTEXITCODE }
finally { $ErrorActionPreference = $savedPreference }
if ($InjectRegression) {
    if ($runCode -eq 0) { throw "Native regression escaped the model: $InjectRegression" }
    # An assertion from the actual driver code is required, not a compiler failure.
    if (-not ((Get-Content (Join-Path $OutputDirectory 'run.log') -Raw) -match 'Assertion failed')) { throw "Expected assertion was not observed: $InjectRegression" }
    [ordered]@{ status='DETECTED'; regression=$InjectRegression; expected=$regression.detected; kernelExecuted=$false } |
        ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'mutation.json') -Encoding UTF8
    Write-Output ('DETECTED: ' + $regression.detected)
    exit 0
}
if ($runCode -ne 0) { throw "Native model failed: $OutputDirectory\run.log" }
[ordered]@{ status='PASS'; driverSourceSha256=(Get-FileHash $driver).Hash; driverHeaderSha256=(Get-FileHash $header).Hash;
    modelSourceSha256=(Get-FileHash $source).Hash; functions=$functions; kernelExecuted=$false; hardwareAccessed=$false } |
    ConvertTo-Json -Depth 4 | Set-Content (Join-Path $OutputDirectory 'result.json') -Encoding UTF8
Get-Content (Join-Path $OutputDirectory 'run.log')
