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
    Dim myT1 As SeqT
    Set myT1 = myT.Sort
    Debug.Print
    Debug.Print

    myT1.PrintByPriority
    myT.RemoveAt 3
    
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


Sub testfmtdbgarray()

    Fmt.Dbg "{0},{1},{2},{3}", 10, 20, 30, 40
    Dim myD As KvpA
    Set myD = KvpA.Deb
    myD.AddPairs SeqA("One", "Two", "Three", "Four"), SeqA(1, 2, 3, 4)
    Fmt.Dbg "{0}", myD

    Fmt.Dbg "{0}", Array(Array(1, 2, 3, 4))
End Sub


Sub testredim()

    Dim myNodes() As KvpHNode
    ReDim myNodes(1 To 100)
    Dim myNode As KvpHNode: Set myNode = KvpHNode(1, 1, 1)
    Set myNodes(1) = myNode
    
    ReDim myNodes(1 To 50)
End Sub


Sub TestHTub()

    Dim myK As KvpHA
    Set myK = KvpHA.Deb
    myK.Reinit 1000
    
    
End Sub

Sub TestAssert()

' #If Not twinBasic Then
'     On Error GoTo Boom
'     Set Assert = CreateObject("Rubberduck.AssertClass")

'     Dim myExpected As Variant: myExpected = Array(1, 2, 3)
'     Dim myResult As Variant: myResult = Array(1, 2, "4")

'     Assert.SequenceEquals myExpected, myResult, "Whoops"
    
    
'     Dim myNum1 As Variant: myNum1 = 10
'     Dim myNum2 As Variant: myNum2 = 20
    
'     Assert.AreEqual myNum1, myNum2, "Whoops"
    
'     Debug.Print "Why didn't it go boom"
'     Exit Sub
    
' Boom:
'     Debug.Print "It went boom"
' #End If
End Sub


'Sub testShakersort()
'
'    Dim myA As Variant = Array(20, 10, 60, 100, 50, 40, 90, 30, 70, 80)
'    ' Fmt.Dbg "{0}", myA
'    ' Sorters.ShakerSortArray myA
'    'Debug.Print Stringifier.StringifyArray(myA)
'    'Fmt.Dbg "{0}", myA
'    Dim mySeq As SeqH = SeqH(myA)
'    'Debug.Print Stringifier.StringifyItemByIndex(mySeq)
'    Fmt.Dbg("{0}", mySeq)
'    ' Dim myIndex As Long
'    ' For myIndex = mySeq.FirstIndex To mySeq.LastIndex
'    '     Debug.Print mySeq.Item(myIndex),
'    ' Next
'    ' Debug.Print
'    Set mySeq = mySeq.Sorted
'    ' Debug.Print Stringifier.StringifyItemByIndex(mySeq)
'    ' For myIndex = mySeq.FirstIndex To mySeq.LastIndex
'    '     Debug.Print mySeq.Item(myIndex),
'    ' Next
'    Fmt.Dbg("{0}", mySeq)
'
'
'End Sub
