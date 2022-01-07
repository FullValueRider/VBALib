Attribute VB_Name = "pootle"
Option Explicit

Sub testarraylist()

    Dim myAL As Lyst
    Set myAL = Lyst.Deb
    
    myAL.Add 1
    myAL.Add 2
    myAL.Add 3
    myAL.Add 4
    
    Dim myitem As Variant
    For Each myitem In myAL
    Debug.Print myitem
    Next
End Sub

