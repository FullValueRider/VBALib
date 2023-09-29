VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A location for global helper methods"
'@PredeclaredId
'@Exposed
'@Folder("Helpers")
'@IgnoreModule
Option Explicit


'@ModuleDescription("A location for global helper methods")

Public Function IsNothing(ByRef ipItem As Variant) As Boolean

    If VBA.IsObject(ipItem) Then
        IsNothing = ipItem Is Nothing
    Else
        IsNothing = False
    End If
    
End Function


Public Function IsNotNothing(ByRef ipItem As Variant) As Boolean
    IsNotNothing = Not IsNothing(ipItem)
End Function
