/*******************************************************************************
*
*  (C) COPYRIGHT AUTHORS, 2026
*
*  TITLE:       ENVDETECT.H
*
*  VERSION:     1.00
*
*  DATE:        01 Apr 2026
*
*  Environment diagnostics support.
*
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
* ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED
* TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
* PARTICULAR PURPOSE.
*
*******************************************************************************/

#pragma once

typedef struct _KDU_ENVIRONMENT_INFO {
    BOOL VmArtifactsPresent;
    BOOL SandboxArtifactsPresent;
    BOOL EmulationArtifactsPresent;
    ULONG IndicatorCount;
    WCHAR Description[512];
} KDU_ENVIRONMENT_INFO, * PKDU_ENVIRONMENT_INFO;

VOID KDUDetectEnvironment(
    VOID);
