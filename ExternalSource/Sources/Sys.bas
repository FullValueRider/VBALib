Attribute VB_Name = "Sys"
Attribute VB_Description = "A place for useful VBA things not explicitly provided by VBA"
    
'@ModuleDescription("A place for useful VBA things not explicitly provided by VBA")
Option Explicit
'@Folder("VBALib")

Public Enum Extent

    IsFirstIndex = 0
    IsLbound = 0
    IsLastIndex = 1
    IsUbound = 1
    IsCount = 2
    IsSpanFirstIndex = 3
    
End Enum

Public Enum StartRun
    
    StartIndex = 0
    IsRun = 1
    
End Enum

'@Ignore ConstantNotUsed
Public Const MaxLong                    As Long = &H7FFFFFFF
Public Const MinLong                    As Long = &HFFFFFFFF

Public Const Failed                     As Long = -1

'Arrays supppoert
'Public Const ParamArrayIsEmpty          As Long = -1


'TypesIterables Support
'Public Const ThisLib                    As String = "VBALib"

Public Function AsOneItem(ByVal ipIterable As Variant) As Variant
    AsOneItem = Array(ipIterable)
End Function

