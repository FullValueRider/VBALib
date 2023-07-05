Attribute VB_Name = "Pootle"
Option Explicit
'@IgnoreModule
Sub TestStack()
    Dim myS As Stack
    Set myS = New Stack
    With myS
        .Push 1
        .Push 2
        .Push 3
        .Push 4
    
    
    End With
    
    Dim myItem As Variant
    For Each myItem In myS
        Debug.Print myItem
    Next
    
    Dim myArray As Variant
    myArray = myS.ToArray
    For Each myItem In myArray
        Debug.Print myItem
    Next
    
End Sub

Sub TestEmpty()
    Debug.Print TypeName(Empty)
End Sub


Sub TestRemove()

    Dim myC As Collection
    Set myC = New Collection
    If myC.Count > 0 Then
    myC.Remove 1
    End If
End Sub

