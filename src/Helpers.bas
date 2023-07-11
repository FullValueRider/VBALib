Attribute VB_Name = "Helpers"
Option Explicit

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
        Set ipLHS = myTemp
    Else
        ipLHS = myTemp
    End If
    
End Sub
