Attribute VB_Name = "TestKvpA"
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

'@TestMethod("KvpA")
Private Sub Test01_ObjAndName()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb
    Dim myExpected As Variant
    myExpected = Array(True, "KvpA", "KvpA")
    
    Dim myresult(0 To 2) As Variant
    
    'Act:
    myresult(0) = VBA.IsObject(myK)
    myresult(1) = VBA.TypeName(myK)
    myresult(2) = myK.TypeName
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


'@TestMethod("KvpA")
Private Sub Test02_Add_ThreeItems()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb
    
    Dim myItemsExpected As Variant
    myItemsExpected = Array(3, "Hello", True)
    ReDim Preserve myItemsExpected(1 To 3)
    
    Dim myKeysExpected As Variant
    myKeysExpected = Array(1, 2, 3)
    ReDim Preserve myKeysExpected(1 To 3)
    
    
    Dim myItemsResult As Variant
    Dim myKeysResult As Variant
    
    'Act:
    myK.Add 1, 3
    myK.Add 2, "Hello"
    myK.Add 3, True
    
    myItemsResult = myK.Items.ToArray
    myKeysResult = myK.Keys.ToArray
    'Assert:
    Assert.SequenceEquals myItemsExpected, myItemsResult
    Assert.SequenceEquals myKeysExpected, myKeysResult
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpA")
Private Sub Test03_Add_Pairs()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb
    
    Dim myItemsExpected As Variant
    myItemsExpected = Array(3, "Hello", True)
    ReDim Preserve myItemsExpected(1 To 3)
    
    Dim myKeysExpected As Variant
    myKeysExpected = Array(1, 2, 3)
    ReDim Preserve myKeysExpected(1 To 3)
    
    
    Dim myItemsResult As Variant
    Dim myKeysResult As Variant
    
    'Act:
    myK.AddPairs SeqA(1, 2, 3), SeqA(3, "Hello", True)
   
    myItemsResult = myK.Items.ToArray
    myKeysResult = myK.Keys.ToArray
    'Assert:
    Assert.SequenceEquals myItemsExpected, myItemsResult
    Assert.SequenceEquals myKeysExpected, myKeysResult
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("KvpA")
Private Sub Test04a_GetItem()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&), SeqA(3, "Hello", True))
   
    Dim myExpected As String
    myExpected = "Hello"
    
    Dim myresult As String
    
    'Act:
    myresult = myK.Item(2&)
    
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

'@TestMethod("KvpA")
Private Sub Test04b_LetItem()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&), SeqA(3, "Hello", True))
   
    Dim myExpected As String
    myExpected = "World"
    
    Dim myresult As String
    
    'Act:
    myK.Item(2) = "World"
    myresult = myK.Item(2&)
    
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


'@TestMethod("KvpA")
Private Sub Test04c_SetItem()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&), SeqA(3, "Hello", True))
   
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&)
    ReDim Preserve myExpected(1 To 3)
    
    Dim myresult As Variant
    
    'Act:
    Set myK.Item(2) = SeqA(1&, 2&, 3&)
    myresult = myK.Item(2&).ToArray
    
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


'@TestMethod("KvpA")
Private Sub Test05a_Remove()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, True, 1&, 2&, 3&, 4&)
    ReDim Preserve myExpected(1 To 6)
    
    Dim myresult As Variant
    
    'Act:
    myK.Remove 2&
    myresult = myK.Items.ToArray
    
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


'@TestMethod("KvpA")
Private Sub Test05b_Remove()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, True, 2&, 4&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant
    
    'Act:
    myK.Remove 2&, 4&, 6&
    myresult = myK.Items.ToArray
    
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

'@TestMethod("KvpA")
Private Sub Test06_RemoveAfter()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, "Hello", 3&, 4&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.RemoveAfter(2&, 3).Items.ToArray
    
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


'@TestMethod("KvpA")
Private Sub Test07_RemoveBefore()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, "Hello", 3&, 4&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.RemoveBefore(6&, 3).Items.ToArray
    
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

'@TestMethod("KvpA")
Private Sub Test08a_RemoveAll()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Long
    myExpected = 0
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.RemoveAll.Count
    
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

'@TestMethod("KvpA")
Private Sub Test08b_Clear()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Long
    myExpected = 0
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.Clear.Count
    
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

'@TestMethod("KvpA")
Private Sub Test08c_Reset()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Long
    myExpected = 0
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.Clear.Count
    
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


'@TestMethod("KvpA")
Private Sub Test09_Clone()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb
    
    Dim myItemsExpected As Variant
    myItemsExpected = SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&).ToArray
    
    Dim myKeysExpected As Variant
    myKeysExpected = SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&).ToArray
    
    Dim myItemsResult As Variant
    Dim myKeysResult As Variant
    
    'Act:
    Dim myT As KvpA
    
    Set myT = myK.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&)).Clone
   
    myItemsResult = myT.Items.ToArray
    myKeysResult = myT.Keys.ToArray
    
    'Assert:
    Assert.SequenceEquals myItemsExpected, myItemsResult
    Assert.SequenceEquals myKeysExpected, myKeysResult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("KvpA")
Private Sub Test10_Hold_Lacks_FilledSeqA()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(True, False, True, True, False, False, False, False, True, True, True, True, False, False, False, False, True, True)
    ReDim Preserve myExpected(1 To 18)
    
    Dim myresult As Variant
    ReDim myresult(1 To 18)
    'Act:
    myresult(1) = myK.HoldsItems
    myresult(2) = myK.LacksItems
    
    myresult(3) = myK.HoldsItem("Hello")
    myresult(4) = myK.HoldsItem(4&)
    myresult(5) = myK.HoldsItem(42&)
    myresult(6) = myK.HoldsItem("World")
    
    myresult(7) = myK.LacksItem("Hello")
    myresult(8) = myK.LacksItem(4&)
    myresult(9) = myK.LacksItem(42&)
    myresult(10) = myK.LacksItem("World")
    
    myresult(11) = myK.HoldsKey(2&)
    myresult(12) = myK.HoldsKey(6&)
    myresult(13) = myK.HoldsKey(42&)
    myresult(14) = myK.HoldsKey("Hello")
    
    myresult(15) = myK.LacksKey(2&)
    myresult(16) = myK.LacksKey(6&)
    myresult(17) = myK.LacksKey(42&)
    myresult(18) = myK.LacksKey("Hello")
    
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


'@TestMethod("KvpA")
Private Sub Test11_MappedIt()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(4&, "Hellp", True, 2&, 3&, 4&, 5&)
    ReDim Preserve myExpected(1 To 7)
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.MappedIt(mpInc.Deb).Items.ToArray
    
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

'@TestMethod("KvpA")
Private Sub Test12_MapIt()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myOrigExpected As Variant
    myOrigExpected = Array(3&, "Hello", True, 1&, 2&, 3&, 4&)
    ReDim Preserve myOrigExpected(1 To 7)
    
    Dim myMapExpected As Variant
    myMapExpected = Array(4&, "Hellp", True, 2&, 3&, 4&, 5&)
    ReDim Preserve myMapExpected(1 To 7)
    
    Dim myOrigResult As Variant
    Dim myMapresult As KvpA
    
    'Act:
    myOrigResult = myK.Items.ToArray
    Set myMapresult = myK.MapIt(mpInc.Deb)
    
    'Assert:
    Assert.SequenceEquals myOrigExpected, myOrigResult
    Assert.SequenceEquals myMapExpected, myMapresult.Items.ToArray
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("KvpA")
Private Sub Test13_FilterIt()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, 3&, 4&)
    ReDim Preserve myExpected(1 To 3)
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.FilterIt(cmpMT(2)).Items.ToArray
    
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

'@TestMethod("KvpA")
Private Sub Test14_ReduceIt()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As LongLong
    myExpected = VBA.CLngLng(1 + 2 + 3 + 4 + 3)
    
    Dim myresult As LongLong
    
    'Act:
    myresult = myK.ReduceIt(rdSum)
    
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

'@TestMethod("KvpA")
Private Sub Test15a_KeyByIndex()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 30&
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.KeyByIndex(3)
    
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


'@TestMethod("KvpA")
Private Sub Test15b_KeyOf()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 20&
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.KeyOf("Hello")
    
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

'@TestMethod("KvpA")
Private Sub Test16a_GetFirst()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 3&
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.First
    
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

'@TestMethod("KvpA")
Private Sub Test16b_LetFirst()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 42&
    
    Dim myresult As Variant
    
    'Act:
    myK.First = 42&
    myresult = myK.First
    
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

'@TestMethod("KvpA")
Private Sub Test16c_GetLast()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 4&
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.Last
    
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

'@TestMethod("KvpA")
Private Sub Test16d_LetLast()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 42&
    
    Dim myresult As Variant
    
    'Act:
    myK.Last = 42&
    myresult = myK.Last
    
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

'@TestMethod("KvpA")
Private Sub Test16e_GetFirstKey()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 10&
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.FirstKey
    
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

'@TestMethod("KvpA")
Private Sub Test16f_LastKey()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpA
    Set myK = KvpA.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 70&
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.LastKey
    
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

