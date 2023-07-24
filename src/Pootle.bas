Attribute VB_Name = "Pootle"
Option Explicit
'@IgnoreModule


Sub TestSeqLAdd()
    Dim myS As SeqL
    Set myS = SeqL.Deb
    Debug.Print myS.Add(42)
    Debug.Print myS.Add(43)
    Debug.Print myS.Add(44)
End Sub

Sub TestStrConv()
    Dim myS As String
    myS = "Hello"
    Dim myA As Variant
    myA = StrConv(myS, vbNarrow)
End Sub

Sub TestKvp()
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
