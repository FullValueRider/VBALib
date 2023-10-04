Attribute VB_Name = "HelpVBA"
'@IgnoreModule
'@Folder("Constants")
'@ModuleDescription("A location for helpers which enhance the functionality of VBA")
Option Explicit

Public Const minLongLong                    As LongLong = &HFFFFFFFFFFFFFFFF^
Public Const maxLonglong                    As LongLong = &H7FFFFFFFFFFFFFFF^

Public Const maxLong                        As Long = &H7FFFFFFF
Public Const minLong                        As Long = &HFFFFFFFF

'@ is the symbol for currecy not decimal
Public Const maxDecimal As Variant = 7.92281625142643E+28
Public Const MinDecimal As Variant = -7.92281625142643E+28

'Public Function maxDecimal() As Variant
'    maxDecimal = VBA.CDec(7.92281625142643E+28)
'End Function
'
'
'Public Function MinDecimal() As Variant
'    MinDecimal = VBA.CDec(-7.92281625142643E+28)
'End Function


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

Public Sub Swap(ByRef ipLHS As Variant, ByRef ipRhs As Variant)

    Dim myTemp As Variant
    
    If VBA.IsObject(ipLHS) Then
        Set myTemp = ipLHS
    Else
        myTemp = ipLHS
    End If
    
    If VBA.IsObject(ipRhs) Then
        Set ipLHS = ipRhs
    Else
        ipLHS = ipRhs
    End If
    
    If VBA.IsObject(myTemp) Then
        Set ipRhs = myTemp
    Else
        ipRhs = myTemp
    End If
    
End Sub
