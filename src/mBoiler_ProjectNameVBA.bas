Attribute VB_Name = "mBoiler_ProjectNameVBA"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2011, David Briant. All rights reserved.
'    Licensed under BSD 3-Clause License - see https://github.com/DangerMouseB
'
'************************************************************************************************************************************************

Option Explicit
Option Private Module

Function GLOBAL_PROJECT_NAME() As String
    GLOBAL_PROJECT_NAME = Application.ThisWorkbook.name
End Function

