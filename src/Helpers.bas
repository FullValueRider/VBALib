Attribute VB_Name = "Helpers"
Option Explicit

Public Sub Swap(ByRef ipLHS As Variant, ByRef ipRHS As Variant)

    Dim myTemp As Variant
    
    If VBA.IsObject(ipLHS) Then
        Set myTemp = ipLHS
    Else
        myTemp = ipLHS
    End If
    
    If VBA.IsObject(ipRHS) Then
        Set ipLHS = ipRHS
    Else
        ipLHS = ipRHS
    End If
    
    If VBA.IsObject(myTemp) Then
        Set ipLHS = myTemp
    Else
        ipLHS = myTemp
    End If
    
End Sub
