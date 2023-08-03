Attribute VB_Name = "Pootle"
Option Explicit
'@IgnoreModule


'@Folder("Pootle")
Sub TestSeqHAdd()
    Dim myS As SeqH
    Set myS = SeqH(10, 40, 30, 40, 40, 80, 40, 90, 100)
    myS.PrintByHash
    myS.PrintByOrder
End Sub


Sub TestStrConv()
    Dim myS As String
    myS = "Hello"
    Dim myA As Variant
    myA = StrConv(myS, vbNarrow)
End Sub


Sub TestKvp()

    Dim myS As SeqA: Set myS = SeqA()
    Dim myK As KvpC
    Set myK = KvpC.Deb
    
    With myK
        .Add "One", 10
        .Add "Two", 20
        .Add "Three", 30
        .Add "Four", 40
    
    
        Dim myItems As IterItems: Set myItems = IterItems(myK)
        Do
            Debug.Print myItems.CurItem(0), myItems.CurKey(0), myItems.CurOffset(0)
            Debug.Print myItems.CurItem(1), myItems.CurKey(1), myItems.CurOffset(1)
            Debug.Print myItems.CurItem(2), myItems.CurKey(2), myItems.CurOffset(2)
            Debug.Print myItems.CurItem(3), myItems.CurKey(3), myItems.CurOffset(3)
        
        Loop Until myItems.MoveNext
    
    End With
    
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

    Fmt.Dbg "{0}", Array(1, 2, 3, 4)
End Sub


Sub testredim()

    Dim myNodes() As KvpHNode
    ReDim myNodes(1 To 100)
    Dim myNode As KvpHNode: Set myNode = KvpHNode(Nothing, Nothing, 1, 1, 1)
    Set myNodes(1) = myNode
    
    ReDim myNodes(1 To 50)
End Sub


Sub TestHTub()

    Dim myK As KvpH
    Set myK = KvpH.Deb
    myK.SetSize 1000
    
    
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
