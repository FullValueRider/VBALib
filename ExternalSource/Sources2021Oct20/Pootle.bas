Attribute VB_Name = "Pootle"
'@IgnoreModule
'@Folder("VBALib")
Public Sub TestArrayInLyst()
    
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    
    Dim myArray1 As Lyst
    Set myArray1 = Lyst.Deb.AddValidatedIterable(Array(1, 2, 3, 4, 5, 6))
    
    Dim myArray2 As Lyst
    Set myArray2 = Lyst.Deb.AddValidatedIterable(Array(2, 3, 4, 5, 6, 7))
    
    Dim myArray3 As Lyst
    Set myArray3 = Lyst.Deb.AddValidatedIterable(Array(3, 4, 5, 6, 7, 8))
    
    myLyst.Add myArray1
    myLyst.Add myArray2
    myLyst.Add myArray3
    
    Debug.Print myLyst.IndexOf(myArray2)
End Sub
    
