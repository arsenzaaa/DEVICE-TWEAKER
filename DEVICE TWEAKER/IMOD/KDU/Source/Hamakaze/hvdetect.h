/*******************************************************************************
*
*  (C) COPYRIGHT AUTHORS, 2020 - 2026
*
*  TITLE:       HVDETECT.H
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

#pragma once

typedef enum _KDU_HYPERVISOR_VENDOR {
    HypervisorVendorUnknown = 0,
    HypervisorVendorMicrosoft,
    HypervisorVendorVMware,
    HypervisorVendorKVM,
    HypervisorVendorXen,
    HypervisorVendorVirtualBox,
    HypervisorVendorParallels,
    HypervisorVendorBhyve,
    HypervisorVendorAcrn
} KDU_HYPERVISOR_VENDOR;

typedef struct _KDU_HYPERVISOR_INFO {
    BOOL Present;
    BOOL VendorKnown;
    BOOL QueriedByKernel;
    BOOL QueriedByCpuid;
    ULONG MaxLeaf;
    KDU_HYPERVISOR_VENDOR VendorId;
    CHAR VendorName[32];
    CHAR FriendlyName[64];
} KDU_HYPERVISOR_INFO, * PKDU_HYPERVISOR_INFO;

VOID KDUDetectHypervisor(
    VOID);

VOID KDUQueryHypervisorInformation(
    _Out_ PKDU_HYPERVISOR_INFO HypervisorInfo); 
