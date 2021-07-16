Attribute VB_Name = "mDBAPI_BLIP"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2012, David Briant. All rights reserved.
'    Licensed under BSD 3-Clause License - see https://github.com/DangerMouseB
'
'************************************************************************************************************************************************

Option Explicit
Option Private Module

Declare Function apiVariantFromBuffer Lib "BLIP.dll" Alias "uVariantFromBuffer" (ByVal schemaID As Long, ByVal pBuffer As Long, ByVal length As Long, oVar As Variant) As Long
Declare Function apiVariantToBuffer Lib "BLIP.dll" Alias "uVariantToBuffer" (ByVal schemaID As Long, var As Variant, ByVal pBuffer As Long, ByVal length As Long) As Long
Declare Function apiBufferSizeForVariant Lib "BLIP.dll" Alias "uBufferSizeForVariant" (ByVal schemaID As Long, var As Variant, oLength As Long) As Long

