Attribute VB_Name = "Pootle"
Option Explicit
'@IgnoreModule
Sub testingnulls()
    Debug.Print TypeName(vbNullString) ' string
    Debug.Print TypeName(Null)         ' Null
    Debug.Print TypeName(Nothing)      ' Nothing
    Debug.Print Nothing Is Nothing     ' True
    Debug.Print GroupInfo.IsBoolean(Null) ' False
    'Debug.Print Nothing > 1            ' invalid use of object
    'Debug.Print 1 > Nothing            ' Invalid use of object
    Debug.Print Empty < 1               ' True
    Debug.Print 1 < Empty               ' false
    Debug.Print -1 < Empty              ' True
    Debug.Print IsNothing(Empty)       ' False - library defined result is False
End Sub
Sub TestTreap()
    Dim myT As SeqT: Set myT = SeqT.Deb
    Debug.Print myT.Add("Hello")
    Debug.Print myT.Add("There")
    Debug.Print myT.Add("World")
    Debug.Print myT.Add("Its")
    Debug.Print myT.Add("A")
    Debug.Print myT.Add("Nice")
    Debug.Print myT.Add("Day")
    Debug.Print myT.Count
    myT.PrintByPriority
    Debug.Print
    Debug.Print
    myT.PrintByOrder
    myT.RemoveAt myT.Lastindex
    Debug.Print
    Debug.Print
    myT.PrintByOrder
    
    Dim myT1 As SeqT
    Set myT1 = myT.Sort
    Debug.Print
    Debug.Print

    myT1.PrintByPriority
    myT1.PrintByOrder
    myT1.Remove "World"
    Debug.Print
    Debug.Print
    
    myT1.PrintByPriority
    
End Sub


Sub TestTreap2()
    Dim myT As SeqT: Set myT = SeqT.Deb
    myT.Add 300
    myT.Add 300
    myT.Add 300
    myT.Add 300
    Dim myItem As Long
    For myItem = 100 To 10000 Step 10
        myT.Add myItem
    Next
    myT.Add 400
    myT.Add 400
    myT.Add 400
    myT.Add 400
    myT.Add 300
    myT.Add 300
    myT.Add 300
    myT.Add 300
    myT.Add 300
    myT.Add 300
    For myItem = 100 To 10000 Step 10
        myT.Add myItem
    Next

    Debug.Print myT.Count, myT.Count(300)
    Debug.Print myT.Count, myT.Count(400)
    Debug.Print myT.Count, myT.Count(300)

End Sub
'@Folder("Pootle")
Sub TestSeqHAdd()
    Dim myS As SeqHL
    Set myS = SeqHL(10, 40, 30, 40, 40, 80, 40, 90, 100)
'    myS.PrintByHash
'    myS.PrintByOrder
End Sub


Sub TestStrConv()
    Dim myS As String
    myS = "Hello"
    Dim myA As Variant
    myA = StrConv(myS, vbNarrow)
End Sub


Sub TestmyKvpC()

    Dim myK As KvpC
    Set myK = KvpC.Deb
    
    With myK
        .Add "One", 10
        .Add "Two", 20
        .Add "Three", 30
        .Add "Four", 40
    End With
    
        Dim myKeys As Variant: Set myKeys = myK.KeysAsSeq
        Dim myIndex As Long
        For myIndex = myKeys.FirstIndex To myKeys.Lastindex
            Debug.Print myK.Item(myKeys.Item(myIndex))
        Next
    
End Sub


Sub TestPerm()
    Dim myP As Collection
    Set myP = Maths.Permutations(Array(1, 2, 3, 4), 4)
    Fmt.Dbg "{0}", myP
End Sub



Sub TestSeqT()
    Dim myT As SeqT: Set myT = SeqT.Deb
    
    Dim myValue As Long
    For myValue = 10 To 200 Step 10
        myT.Add myValue
    Next
    
    
    Debug.Print
    Debug.Print "Inorder traversal of the given tree"
    'myT.PrintByOrder
     
    Debug.Print
    Debug.Print "Delete 20"
    myT.Remove 20
    Debug.Print "Inorder traversal of the modified tree"
    'myT.PrintByOrder
    
    Debug.Print "Delete 300"
    myT.Remove 3000
    Debug.Print "Inorder traversal of the modified tree"
    'myT.PrintByOrder
    Debug.Print
    Debug.Print "Delete 500"
    Debug.Print
    
    myT.Remove 500
    Debug.Print "Inorder traversal of the modified tree"
    'myT.PrintByOrder
    Debug.Print
    Debug.Print "Remove 510"
    myT.Remove 510
    Debug.Print
    Debug.Print "Remove 730"
    myT.Remove 730
 
    Debug.Print
    Debug.Print "Seeking 500"
    If myT.HoldsItem(500) Then
        Debug.Print "500 found"
    Else
        Debug.Print "500 Not Found"
    End If
    Debug.Print
    Debug.Print "Order collection"
    'myT.PrintByColl
End Sub

Sub TestSeqTAt()
    Dim myT As SeqT: Set myT = SeqT.Deb
    
    Dim myA As Variant: myA = Array(90, 20, 150, 30, 120, 40, 70, 90, 100, 110, 50, 130, 140, 60, 160, 80, 170, 190, 200, 90, 10, 90, 90, 90)
    Dim myItem As Variant
    For Each myItem In myA
        myT.Add myItem
'        Debug.Print
'        myT.PrintByOrder
'        Debug.Print
'        myT.PrintByColl
    Next
    
    
    Debug.Print
    Debug.Print "Inorder traversal of the given tree"
    myT.PrintByOrder
     
    Debug.Print
    Debug.Print "Delete at 2"
    myT.RemoveAt 2
    Debug.Print "Inorder traversal of the modified tree"
    myT.PrintByOrder
    
    Debug.Print "Delete at 8"
    myT.RemoveAt 8
    Debug.Print "Inorder traversal of the modified tree"
    myT.PrintByOrder
    Debug.Print
    Debug.Print "Delete At 14"
    
    myT.RemoveAt 14
    Debug.Print "Inorder traversal of the modified tree"
    myT.PrintByOrder
    
    
 
    Debug.Print
    Debug.Print "Seeking 50"
    If myT.HoldsItem(50) Then
        Debug.Print "50 found"
    Else
        Debug.Print "50 Not Found"
    End If
    Debug.Print
    Debug.Print "Order collection"
    myT.PrintByColl
End Sub
