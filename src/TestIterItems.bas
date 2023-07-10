Attribute VB_Name = "TestIterItems"
'@TestModule
'@Folder("Tests")
'@IgnoreModule

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub


'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub


'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'@TestMethod("IterItems")
Private Sub Test01a_IsObjectAndName()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, "IterItems", "IterItems")
    
    Dim myresult As Variant
    ReDim myresult(0 To 2)
    
    Dim myI As IterItems
    Set myI = IterItems(SeqC(1, 2, 3, 4, 5))
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult(0) = VBA.IsObject(myI)
    myresult(1) = "IterItems"
    myresult(2) = "IterItems"
   
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test02a_GetItem0Seq()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 1
    
    Dim myresult As Variant
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(1, 2, 3, 4, 5))
       
    myresult = myI.CurItem(0)
   
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test02b_GetItem0SeqAfterThreeMovenext()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, True, True, 40)
    
    Dim myresult As Variant
    ReDim myresult(0 To 3)
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50))
    myresult(0) = myI.MoveNext
    myresult(1) = myI.MoveNext
    myresult(2) = myI.MoveNext
    myresult(3) = myI.CurItem(0)
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test02c_GetItem0SeqAfterThreeMoveNextTwoMovePrev()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, True, True, True, True, 20)
    
    Dim myresult As Variant
    ReDim myresult(0 To 5)
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50))
    myresult(0) = myI.MoveNext
    myresult(1) = myI.MoveNext
    myresult(2) = myI.MoveNext
    myresult(3) = myI.MovePrev
    myresult(4) = myI.MovePrev
    myresult(5) = myI.CurItem(0)
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test03a_GetItemSeqAtOffset3()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 80
    
    Dim myresult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50, 60, 70, 80, 90))
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myresult = myI.CurItem(3)
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test03b_GetItemSeqAtOffsetMinus3()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 20
    
    Dim myresult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50, 60, 70, 80, 90))
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myresult = myI.CurItem(-3)
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test03c_GetItemSeqIndexGreaterThanSize()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = True
    
    Dim myresult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50, 60, 70, 80, 90))
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myresult = myI.CurItem(5)
    
    'Assert:
    Assert.AreEqual myExpected, VBA.IsNull(myresult)
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test03d_GetItemSeqIndexDeforeIndex1()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = True
    
    Dim myresult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50, 60, 70, 80, 90))
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myresult = myI.CurItem(-5)
    
    'Assert:
    Assert.AreEqual myExpected, VBA.IsNull(myresult)
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test04a_GetKeySeq()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myresult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50, 60, 70, 80, 90))
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myresult = myI.CurKey(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test04b_GetIndexSeq()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myresult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50, 60, 70, 80, 90))
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myresult = myI.CurOffset(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test05a_GetItemArray()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 50

    Dim myArray As Variant
    myArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    ReDim Preserve myArray(-4 To 4)
    
    Dim myresult As Variant
    
    'Act:
    Dim myI As IterItems
    Set myI = IterItems(myArray)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myresult = myI.CurItem(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test05b_GetKeyArray()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 0&
    
    Dim myArray As Variant
    myArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    ReDim Preserve myArray(-4 To 4)
    
    Dim myresult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(myArray)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myresult = myI.CurKey(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test05c_GetIndexArray()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myArray As Variant
    myArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    ReDim Preserve myArray(-4 To 4)
    
    Dim myresult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(myArray)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myresult = myI.CurOffset(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test06a_GetItemCollection()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 50

    Dim myC As Collection
    Set myC = New Collection
    With myC
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    Dim myresult As Variant
    
    'Act:
    myresult = myI.CurItem(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test06b_GetKeyCollection()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myC As Collection
    Set myC = New Collection
    With myC
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myresult As Variant
    
    
    'Act:
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myresult = myI.CurKey(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test06c_GetIndexCollection()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myC As Collection
    Set myC = New Collection
    With myC
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    Dim myresult As Variant
    
    
    'Act:
    myresult = myI.CurOffset(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test07a_GetItemArrayList()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 50

    Dim myC As ArrayList
    Set myC = New ArrayList
    With myC
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    Dim myresult As Variant
    
    'Act:
    myresult = myI.CurItem(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test07b_GetKeyArrayList()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 4&
    
    Dim myC As ArrayList
    Set myC = New ArrayList
    With myC
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myresult As Variant
    
    
    'Act:
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myresult = myI.CurKey(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test07c_GetIndexArrayList()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myC As ArrayList
    Set myC = New ArrayList
    With myC
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    Dim myresult As Variant
    
    
    'Act:
    myresult = myI.CurOffset(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test08a_GetItemDictionary()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 50

    Dim myK As KvpC
    Set myK = KvpC.Deb
    With myK
        .Add "Ten", 10
        .Add "Twenty", 20
        .Add "Thirty", 30
        .Add "Forty", 40
        .Add "Fifty", 50
        .Add "Sixty", 60
        .Add "Seventy", 70
        .Add "Eighty", 80
        .Add "Ninety", 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myK)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    Dim myresult As Variant
    
    'Act:
    myresult = myI.CurItem(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test08b_GetKeyDictionary()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = "Fifty"
    
    Dim myK As KvpC
    Set myK = KvpC.Deb
    With myK
        .Add "Ten", 10
        .Add "Twenty", 20
        .Add "Thirty", 30
        .Add "Forty", 40
        .Add "Fifty", 50
        .Add "Sixty", 60
        .Add "Seventy", 70
        .Add "Eighty", 80
        .Add "Ninety", 90
    End With
    
    Dim myresult As Variant
    
    
    'Act:
    Dim myI As IterItems
    Set myI = IterItems(myK)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myresult = myI.CurKey(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IterItems")
Private Sub Test08c_GetIndexDIctionary()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myK As KvpC
    Set myK = KvpC.Deb
    With myK
        .Add "Ten", 10
        .Add "Twenty", 20
        .Add "Thirty", 30
        .Add "Forty", 40
        .Add "Fifty", 50
        .Add "Sixty", 60
        .Add "Seventy", 70
        .Add "Eighty", 80
        .Add "Ninety", 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myK)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    Dim myresult As Variant
    
    
    'Act:
    myresult = myI.CurOffset(0)
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
'

'@TestMethod("IterItems")
Private Sub Test09a_GetIndexDIctionary()

    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    ReDim Preserve myExpected(1 To 9)
    
    Dim myK As KvpC
    Set myK = KvpC.Deb
    With myK
        .Add "Ten", 10
        .Add "Twenty", 20
        .Add "Thirty", 30
        .Add "Forty", 40
        .Add "Fifty", 50
        .Add "Sixty", 60
        .Add "Seventy", 70
        .Add "Eighty", 80
        .Add "Ninety", 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myK)
    
    
    
    Dim myresult As Variant
    ReDim myresult(1 To 9)
    
    'Act:
        Do
            Debug.Print myI.CurOffset(0), myI.CurItem(0)
            myresult(myI.CurOffset(0)) = VBA.CVar(myI.CurItem(0))
        
        Loop While myI.MoveNext
    
    
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
