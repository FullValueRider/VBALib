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
        ipRhs = myTemp
    End If
    
End Sub

Public Function LineariseArray(ByRef ipArray As Variant) As Variant

    If ArrayOp.Ranks(ipArray) = 1 Then
        LineariseArray = ipArray
        Exit Function
    End If
    
    Dim mySize As Long
    mySize = ArrayOp.Count(ipArray)
    
    Dim myA As Variant
    ReDim myA(1 To mySize)
    
    Dim myIndex As Long: myIndex = 1
    Dim myItem As Variant
    For Each myItem In ipArray
        If VBA.IsObject(myItem) Then
            Set myA(myIndex) = myItem
        Else
            myA(myIndex) = myItem
        End If
        myIndex = myIndex + 1
    Next
    
    LineariseArray = myA
    
End Function
