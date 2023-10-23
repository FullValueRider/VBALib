Attribute VB_Name = "HelpClasses"
'@Folder("Helpers")
'@ModuleDescription("A module for helpers for classes")
Option Explicit

'A factory method for TextFormatter class
Public Function Fmt(Optional ByVal ipWithTypes As e_WithTypes = e_WithTypes.m_None) As Format
    Set Fmt = Format(ipWithTypes)
End Function

'Public Function Num(ByRef ipNumber As Variant) As Number
'    Dim myNum As Number
'    Set myNum = New Number
'    myNum = ipNumber
'
'
'    Set Num = myNum
'
'End Function
