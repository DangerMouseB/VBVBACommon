Attribute VB_Name = "mBoiler_ProjectNameVBA"
'*************************************************************************************************************************************************************************************************************************************************
'
' Copyright (c) David Briant 2009-2011 - All rights reserved
'
'*************************************************************************************************************************************************************************************************************************************************

Option Explicit
Option Private Module

Function GLOBAL_PROJECT_NAME() As String
    GLOBAL_PROJECT_NAME = Application.ThisWorkbook.name
End Function
