/*******************************************************************************
*
*  (C) COPYRIGHT AUTHORS, 2020 - 2026
*
*  TITLE:       PROVDB.CPP
*
*  VERSION:     1.48
*
*  DATE:        01 Apr 2026
*
*  Embedded providers database definitions.
*
* THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
* ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED
* TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
* PARTICULAR PURPOSE.
*
*******************************************************************************/

#include "global.h"

//
// Empty, optional use. Should replicate Tanikaze db layout.
//
KDU_DB_ENTRY gEmbeddedProvEntry[] = { {} };

extern "C" {

    KDU_DB gProvTableEmbedded = {
        RTL_NUMBER_OF(gEmbeddedProvEntry),
        gEmbeddedProvEntry
    };

    KDU_DB_VERSION gVersionEmbedded = {
        KDU_VERSION_MAJOR,
        KDU_VERSION_MINOR,
        KDU_VERSION_REVISION,
        KDU_VERSION_BUILD
    };

}
