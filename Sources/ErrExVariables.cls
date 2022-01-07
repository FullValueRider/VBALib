VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrExVariables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' *****************************************************************************
' *  VCOM library: vbWatchdog  http://www.everythingaccess.com/vbwatchdog.htm *
' *  3.9.2 (Ultimate Edition TRI built May 26 2021                            *
' *---------------------------------------------------------------------------*
' *  This is a demonstration version of the library for TESTING purposes only.*
' *                                                                           *
' *  This software is provided "as-is," without any express or implied        *
' *  warranty.  In no event shall the author be held liable for any           *
' *  consequential, incidental, direct, indirect, special, punitive, or other *
' *  damages arising from the use of this software (including without         *
' *  limitation, damages for loss of business profits, business interruption, *
' *  loss of information or other pecuniary loss).                            *
' *****************************************************************************

Option Explicit

Enum VarScope
    LocalVariable = &H0
    ParameterVariable = &H1
    ModuleVariable = &H2
    StaticVariable = &H3
End Enum

Public VCOMObject As Object      ' DO NOT USE!!! THIS IS FOR INTERNAL USE ONLY!!!
Attribute VCOMObject.VB_VarUserMemId = -4

Sub FirstVar()
    Call VCOMObject.FirstVar
End Sub

Sub NextVar()
    Call VCOMObject.NextVar
End Sub

Property Get IsEnd() As Boolean
    IsEnd = VCOMObject.IsEnd
End Property

Property Get Name() As String
    Name = VCOMObject.Name
End Property

Property Get VarPtr() As Long
    VarPtr = VCOMObject.VarPtr
End Property

Property Get Scope() As VarScope
    Scope = VCOMObject.Scope
End Property

Property Get ScopeDesc() As String
    ScopeDesc = VCOMObject.ScopeDesc
End Property

Property Get TypeDesc() As String
    TypeDesc = VCOMObject.TypeDesc
End Property

Property Get Value() As Variant
    Call VCOMObject.AssignVar(Value, VCOMObject.Value)
End Property

Property Get ValueDesc() As String
    ValueDesc = VCOMObject.ValueDesc
End Property

Function DumpAll(Optional ByVal SplitPos As Long = &H32) As String
    DumpAll = VCOMObject.DumpAll(SplitPos)
End Function

