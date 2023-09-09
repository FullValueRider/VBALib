Attribute VB_Name = "Helpers"
'@Folder("Helpers")
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
        ipRHS = myTemp
    End If
    
End Sub

Public Function IsNothing(ByRef ipItem As Variant) As Boolean

    If Not VBA.IsObject(ipItem) Then
        IsNothing = False
        Exit Function
    End If
    
    IsNothing = ipItem Is Nothing
    
End Function

Public Function IsNotNothing(ByRef ipItem As Object) As Boolean
    IsNotNothing = Not (ipItem Is Nothing)
End Function




