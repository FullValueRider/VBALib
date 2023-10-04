Attribute VB_Name = "HelpClasses"
'@Folder("Helpers")
'@ModuleDescription("A module for helpers for classes
Option Explicit




'Public Function IsNothing(ByRef ipItem As Variant) As Boolean
'
'    If Not VBA.IsObject(ipItem) Then
'        IsNothing = False
'        Exit Function
'    End If
'
'    IsNothing = ipItem Is Nothing
'
'End Function
'
'Public Function IsNotNothing(ByRef ipItem As Object) As Boolean
'    IsNotNothing = Not (ipItem Is Nothing)
'End Function



Public Function Fmt() As Format
    Set Fmt = Format.Deb
End Function

