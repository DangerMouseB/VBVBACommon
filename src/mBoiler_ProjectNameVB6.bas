Attribute VB_Name = "mBoiler_ProjectNameVB6"
'************************************************************************************************************************************************
'
'    Copyright (c) 2009-2011, David Briant. All rights reserved.
'    Licensed under BSD 3-Clause License - see https://github.com/DangerMouseB
'
'************************************************************************************************************************************************

Option Explicit

Function GLOBAL_PROJECT_NAME() As String
    GLOBAL_PROJECT_NAME = App.FileDescription
End Function
