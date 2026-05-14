/*******************************************************************************
*
*  (C) COPYRIGHT AUTHORS, 2020 - 2026
*
*  TITLE:       HVDETECT.CPP
*
*  VERSION:     1.00
*
*  DATE:        25 Mar 2026
*
*  Hypervisor detection support.
*
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
* ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED
* TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
* PARTICULAR PURPOSE.
*
*******************************************************************************/

#include "global.h"
#include "hvdetect.h"

typedef struct _KDU_HYPERVISOR_VENDOR_DESC {
    LPCSTR VendorString;
    KDU_HYPERVISOR_VENDOR VendorId;
    LPCSTR FriendlyName;
} KDU_HYPERVISOR_VENDOR_DESC, * PKDU_HYPERVISOR_VENDOR_DESC;

static const KDU_HYPERVISOR_VENDOR_DESC g_KduHypervisorVendors[] = {
    { "Microsoft Hv",   HypervisorVendorMicrosoft,  "Microsoft Hyper-V" },
    { "Hv#1",           HypervisorVendorMicrosoft,  "Microsoft Hyper-V" },
    { "VMwareVMware",   HypervisorVendorVMware,     "VMware" },
    { "KVMKVMKVM",      HypervisorVendorKVM,        "KVM" },
    { "XenVMMXenVMM",   HypervisorVendorXen,        "Xen" },
    { "VBoxVBoxVBox",   HypervisorVendorVirtualBox, "Oracle VirtualBox" },
    { "prl hyperv  ",   HypervisorVendorParallels,  "Parallels" },
    { "bhyve bhyve ",   HypervisorVendorBhyve,      "bhyve" },
    { "ACRNACRNACRN",   HypervisorVendorAcrn,       "Project ACRN" }
};

/*
* KDUTrimVendorString
*
* Purpose:
*
* Trim trailing whitespace characters from hypervisor vendor string.
*
*/
VOID KDUTrimVendorString(
    _Inout_updates_(BufferSize) PCHAR VendorName,
    _In_ ULONG BufferSize
)
{
    LONG i;

    if (VendorName == NULL || BufferSize == 0)
        return;

    VendorName[BufferSize - 1] = 0;

    for (i = (LONG)_strlen_a(VendorName) - 1; i >= 0; i--) {
        if (VendorName[i] != ' ' &&
            VendorName[i] != '\t' &&
            VendorName[i] != '\r' &&
            VendorName[i] != '\n')
        {
            break;
        }
        VendorName[i] = 0;
    }
}

/*
* KDUIsPrintableVendorString
*
* Purpose:
*
* Verify that hypervisor vendor string contains only printable characters.
*
*/
BOOL KDUIsPrintableVendorString(
    _In_ LPCSTR VendorName,
    _In_ ULONG Length
)
{
    ULONG i;
    CHAR ch;

    if (VendorName == NULL || Length == 0)
        return FALSE;

    for (i = 0; i < Length; i++) {
        ch = VendorName[i];

        if (ch == 0)
            break;

        if (ch < 0x20 || ch > 0x7E)
            return FALSE;
    }

    return TRUE;
}

/*
* KDUClassifyHypervisorVendor
*
* Purpose:
*
* Classify hypervisor vendor string and assign friendly vendor name.
*
*/
VOID KDUClassifyHypervisorVendor(
    _Inout_ PKDU_HYPERVISOR_INFO HypervisorInfo
)
{
    ULONG i;

    HypervisorInfo->VendorId = HypervisorVendorUnknown;
    HypervisorInfo->FriendlyName[0] = 0;

    if (!HypervisorInfo->VendorKnown)
        return;

    KDUTrimVendorString(HypervisorInfo->VendorName, RTL_NUMBER_OF(HypervisorInfo->VendorName));

    for (i = 0; i < RTL_NUMBER_OF(g_KduHypervisorVendors); i++) {
        if (_strcmpi_a(HypervisorInfo->VendorName, g_KduHypervisorVendors[i].VendorString) == 0) {
            HypervisorInfo->VendorId = g_KduHypervisorVendors[i].VendorId;
            _strcpy_a(HypervisorInfo->FriendlyName, g_KduHypervisorVendors[i].FriendlyName);
            return;
        }
    }

    _strcpy_a(HypervisorInfo->FriendlyName, HypervisorInfo->VendorName);
}

/*
* KDUCaptureHypervisorVendorFromCpuid
*
* Purpose:
*
* Query hypervisor vendor string and maximum leaf through CPUID hypervisor interface.
*
*/
VOID KDUCaptureHypervisorVendorFromCpuid(
    _Inout_ PKDU_HYPERVISOR_INFO HypervisorInfo
)
{
    INT cpuInfo[4];

    cpuInfo[0] = -1;
    cpuInfo[1] = -1;
    cpuInfo[2] = -1;
    cpuInfo[3] = -1;

    __cpuid(cpuInfo, 0x40000000);

    HypervisorInfo->QueriedByCpuid = TRUE;
    HypervisorInfo->MaxLeaf = (ULONG)cpuInfo[0];

    RtlSecureZeroMemory(HypervisorInfo->VendorName, sizeof(HypervisorInfo->VendorName));

    RtlCopyMemory(&HypervisorInfo->VendorName[0], &cpuInfo[1], sizeof(INT));
    RtlCopyMemory(&HypervisorInfo->VendorName[4], &cpuInfo[2], sizeof(INT));
    RtlCopyMemory(&HypervisorInfo->VendorName[8], &cpuInfo[3], sizeof(INT));
    HypervisorInfo->VendorName[12] = 0;

    if (KDUIsPrintableVendorString(HypervisorInfo->VendorName, 12)) {
        HypervisorInfo->VendorKnown = (HypervisorInfo->VendorName[0] != 0);
    }
    else {
        RtlSecureZeroMemory(HypervisorInfo->VendorName, sizeof(HypervisorInfo->VendorName));
        HypervisorInfo->VendorKnown = FALSE;
    }
}

/*
* KDUQueryHypervisorInformation
*
* Purpose:
*
* Query hypervisor presence and vendor information using kernel interface first
* and CPUID fallback when needed.
*
*/
VOID KDUQueryHypervisorInformation(
    _Out_ PKDU_HYPERVISOR_INFO HypervisorInfo
)
{
    ULONG returnLength;
    NTSTATUS ntStatus;
    INT cpuInfo[4];
    SYSTEM_HYPERVISOR_DETAIL_INFORMATION hdi;
    PHV_VENDOR_AND_MAX_FUNCTION pvi;

    RtlSecureZeroMemory(HypervisorInfo, sizeof(KDU_HYPERVISOR_INFO));
    RtlSecureZeroMemory(&hdi, sizeof(hdi));

    returnLength = 0;
    ntStatus = NtQuerySystemInformation(SystemHypervisorDetailInformation,
        &hdi, sizeof(hdi), &returnLength);

    if (NT_SUCCESS(ntStatus)) {

        pvi = (PHV_VENDOR_AND_MAX_FUNCTION)&hdi.HvVendorAndMaxFunction.Data;

        HypervisorInfo->Present = TRUE;
        HypervisorInfo->QueriedByKernel = TRUE;
        HypervisorInfo->MaxLeaf = pvi->MaxFunction;

        RtlSecureZeroMemory(HypervisorInfo->VendorName, sizeof(HypervisorInfo->VendorName));
        RtlCopyMemory(HypervisorInfo->VendorName, pvi->VendorName, 12);
        HypervisorInfo->VendorName[12] = 0;

        if (KDUIsPrintableVendorString(HypervisorInfo->VendorName, 12) &&
            HypervisorInfo->VendorName[0] != 0)
        {
            HypervisorInfo->VendorKnown = TRUE;
        }
        else {
            RtlSecureZeroMemory(HypervisorInfo->VendorName, sizeof(HypervisorInfo->VendorName));
            HypervisorInfo->VendorKnown = FALSE;
        }

        KDUClassifyHypervisorVendor(HypervisorInfo);
        return;
    }

    cpuInfo[0] = -1;
    cpuInfo[1] = -1;
    cpuInfo[2] = -1;
    cpuInfo[3] = -1;

    __cpuid(cpuInfo, 1);

    if (((cpuInfo[2] >> 31) & 1) == 0) {
        return;
    }

    HypervisorInfo->Present = TRUE;

    KDUCaptureHypervisorVendorFromCpuid(HypervisorInfo);
    KDUClassifyHypervisorVendor(HypervisorInfo);
}

/*
* KDUDetectHypervisor
*
* Purpose:
*
* Query hypervisor information and print diagnostic message when hypervisor is present.
*
*/
VOID KDUDetectHypervisor(
    VOID
)
{
    KDU_HYPERVISOR_INFO hypervisorInfo;

    KDUQueryHypervisorInformation(&hypervisorInfo);

    if (!hypervisorInfo.Present)
        return;

    if (hypervisorInfo.VendorKnown) {

        if (hypervisorInfo.FriendlyName[0] != 0) {
            if (_strcmpi_a(hypervisorInfo.FriendlyName, hypervisorInfo.VendorName) == 0) {
                supPrintfEvent(kduEventInformation,
                    "[+] Hypervisor present: \"%s\"\r\n",
                    hypervisorInfo.FriendlyName);
            }
            else {
                supPrintfEvent(kduEventInformation,
                    "[+] Hypervisor present: \"%s\" (%s)\r\n",
                    hypervisorInfo.FriendlyName,
                    hypervisorInfo.VendorName);
            }
        }
        else {
            supPrintfEvent(kduEventInformation,
                "[+] Hypervisor present: \"%s\"\r\n",
                hypervisorInfo.VendorName);
        }

    }
    else {
        supPrintfEvent(kduEventInformation,
            "[+] Hypervisor present, vendor could not be determined\r\n");
    }
}
