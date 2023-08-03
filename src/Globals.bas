Attribute VB_Name = "Globals"
'@IgnoreModule
'@Folder("Constants")
Option Explicit

Public Const minLongLong                    As LongLong = &HFFFFFFFFFFFFFFFF^
Public Const maxLonglong                    As LongLong = &H7FFFFFFFFFFFFFFF^

Public Const maxLong                        As Long = &H7FFFFFFF
Public Const minLong                        As Long = &HFFFFFFFF


'Public Const maxDecimal@ = 7.92281625142643E+28
'Public Const MinDecimal@ = -7.92281625142643E+28
Public Function maxDecimal() As Variant
    maxDecimal = VBA.CDec(7.92281625142643E+28)
End Function


Public Function MinDecimal() As Variant
    MinDecimal = VBA.CDec(-7.92281625142643E+28)
End Function


