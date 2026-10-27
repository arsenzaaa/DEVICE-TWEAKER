# Compile the actual DTIMOD entry/dispatch/MMIO functions against user-mode fake kernel APIs.
# No SYS is executed, no driver/service is loaded, no physical memory is accessed.
[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$OutputDirectory, [switch]$InjectUnloadRegression)
$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$driver = Join-Path $root 'DEVICE TWEAKER\IMOD\Driver\imod_driver.c'
$header = Join-Path $root 'DEVICE TWEAKER\IMOD\Driver\imod_driver.h'
$text = Get-Content -LiteralPath $driver -Raw
$functions = @('ImodDriverUnload','DriverEntry','ImodDriverCreateDriver','ImodDriverInitialize','ImodDriverDispatch',
    'ImodIsValidAccessSize','ImodMapPhysicalMemory','ImodUnmapPhysicalMemory','ImodReadPhysicalMemory','ImodWritePhysicalMemory')
$bodies = foreach ($name in $functions) {
    $match = [regex]::Match($text, '(?m)^(?:static )?(?:NTSTATUS|BOOLEAN|void) '+$name+'\([^;]+?\)\s*\{')
    if (-not $match.Success) { throw "Native function missing: $name" }
    $opening = $text.IndexOf('{',$match.Index); $end = $opening+1; $depth = 1
    while ($depth -gt 0) { if ($text[$end] -eq '{') { $depth++ }; if ($text[$end] -eq '}') { $depth-- }; $end++ }
    $text.Substring($match.Index,$end-$match.Index)
}
if ($InjectUnloadRegression) {
    $joined = $bodies -join "`n`n"
    if (-not $joined.Contains('DriverObject->DriverUnload = ImodDriverUnload;')) { throw 'Unload mutation anchor missing' }
    $bodies = @($joined.Replace('DriverObject->DriverUnload = ImodDriverUnload;', 'DriverObject->DriverUnload = NULL;'))
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
}
int main(void) { _set_error_mode(_OUT_TO_STDERR);_set_abort_behavior(0,_WRITE_ABORT_MSG|_CALL_REPORTFAULT);initTests();mmioTests();dispatchTests();puts("PASS: actual driver entry, unload assignment, MMIO width/map/unmap and dispatch; fake kernel only; PCI worker stubbed");return 0; }
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
if ($InjectUnloadRegression) {
    if ($runCode -eq 0) { throw 'Native unload regression escaped the model' }
    # Assertion from actual initialization is required, not a compiler failure.
    if (-not ((Get-Content (Join-Path $OutputDirectory 'run.log') -Raw) -match 'Assertion failed')) { throw 'Expected initialization assertion was not observed' }
    [ordered]@{ status='DETECTED'; expected='resident initialization assertion'; kernelExecuted=$false } |
        ConvertTo-Json | Set-Content (Join-Path $OutputDirectory 'mutation.json') -Encoding UTF8
    Write-Output 'DETECTED: missing unload assignment in actual driver initialization'
    exit 0
}
if ($runCode -ne 0) { throw "Native model failed: $OutputDirectory\run.log" }
[ordered]@{ status='PASS'; driverSourceSha256=(Get-FileHash $driver).Hash; driverHeaderSha256=(Get-FileHash $header).Hash;
    modelSourceSha256=(Get-FileHash $source).Hash; functions=$functions; kernelExecuted=$false; hardwareAccessed=$false } |
    ConvertTo-Json -Depth 4 | Set-Content (Join-Path $OutputDirectory 'result.json') -Encoding UTF8
Get-Content (Join-Path $OutputDirectory 'run.log')
