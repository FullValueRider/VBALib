Attribute VB_Name = "TestSeqC"
'@IgnoreModule
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Object
'Private Fakes As Object


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    'Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub


'@TestInitialize
'Private Sub TestInitialize()
'    'This method runs before every test in the module..
'End Sub


'@TestCleanup
'Private Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub


'@TestMethod("SeqC")
Private Sub Test01_SeqObj()

    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Set mySeq = SeqC.Deb
    Dim myExpected As Variant
    myExpected = Array(True, "SeqC", "SeqC")
    
    Dim myresult(0 To 2) As Variant
    
    'Act:
    myresult(0) = VBA.IsObject(mySeq)
    myresult(1) = VBA.TypeName(mySeq)
    myresult(2) = mySeq.TypeName
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


'@TestMethod("SeqC")
Private Sub Test02a_InitByLong_10FirstIndex_LastIndex()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Set mySeq = SeqC.Deb(10)
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 10)
    Dim myresult As Variant
    
    'Act:
    myresult = mySeq.ToArray
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    Assert.AreEqual 1&, mySeq.FirstIndex
    Assert.AreEqual 10&, mySeq.LastIndex
    Assert.AreEqual 10&, mySeq.Count
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqC")
Private Sub Test02b_InitByString()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array("H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 5)
    
    Dim myresult As Variant
    
    Dim mySeq As SeqC
    
    'Act:
    Set mySeq = SeqC.Deb("Hello")
    
    myresult = mySeq.ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test02c_InitByForEachArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myArray(1 To 3, 1 To 3) As Variant
    Dim myCount As Long
    myCount = 1
    Dim myFirst As Long
    For myFirst = 1 To 3
    
        Dim mySecond As Long
        For mySecond = 1 To 3
            myArray(myFirst, mySecond) = myCount
            myCount = myCount + 1
        Next
    Next
        
    Dim myExpected As Variant
    myExpected = Array(1&, 4&, 7&, 2&, 5&, 8&, 3&, 6&, 9&)
    ReDim Preserve myExpected(1 To 9)
    
    Dim myresult As Variant
    
    Dim mySeq As SeqC
    
    'Act:
    Set mySeq = SeqC.Deb(myArray)
    
    myresult = mySeq.ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test02d_InitByForEachArrayList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myAL As Object
    Set myAL = CreateObject("System.Collections.Arraylist")
    
    With myAL
        .Add 1
        .Add 4
        .Add 7
        .Add 2
        .Add 5
        .Add 8
        .Add 3
        .Add 6
        .Add 9
    End With
    
    Dim myExpected As Variant
    myExpected = Array(1, 4, 7, 2, 5, 8, 3, 6, 9)
    ReDim Preserve myExpected(1 To 9)
    
    Dim myresult As Variant
    
    Dim mySeq As SeqC
    
    'Act:
    Set mySeq = SeqC.Deb(myAL)
    
    myresult = mySeq.ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test02e_InitByForEachCollection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myC As Collection
    Set myC = New Collection
    
    With myC
        .Add 1
        .Add 4
        .Add 7
        .Add 2
        .Add 5
        .Add 8
        .Add 3
        .Add 6
        .Add 9
    End With
    
    Dim myExpected As Variant
    myExpected = Array(1, 4, 7, 2, 5, 8, 3, 6, 9)
    ReDim Preserve myExpected(1 To 9)
    
    Dim myresult As Variant
    
    Dim mySeq As SeqC
    
    'Act:
    Set mySeq = SeqC.Deb(myC)
    
    myresult = mySeq.ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test02f_InitByDictionary()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myD As Scripting.Dictionary
    Set myD = New Scripting.Dictionary
    
    With myD
        .Add "Hello", "World"
        .Add "Ten", 10&
        .Add "Thing", 3.142
        
    End With
    
    Dim myExpected As Variant
    myExpected = Array("Hello", "World", "Ten", 10&, "Thing", 3.142)
    ReDim Preserve myExpected(1 To 6)
    
    Dim myresult As Variant
    ReDim myresult(1 To 6)
    
    Dim mySeq As SeqC
    
    'Act:
    Set mySeq = SeqC.Deb(myD)
    Dim myTmp As Variant
    myTmp = mySeq.ToArray
    
    myresult(1) = myTmp(1)(0)
    myresult(2) = myTmp(1)(1)
    myresult(3) = myTmp(2)(0)
    myresult(4) = myTmp(2)(1)
    myresult(5) = myTmp(3)(0)
    myresult(6) = myTmp(3)(1)
        
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

'@TestMethod("SeqC")
Private Sub Test03a_WriteItem()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Set mySeq = SeqC.Deb(10)
    mySeq.Item(1) = 42
    mySeq.Item(2) = "Hello"
    mySeq.Item(3) = 3.142
    Dim myExpected As Variant
    myExpected = Array(True, True, True)
    
    Dim myresult As Variant
    ReDim myresult(0 To 2)
    'Act:
    myresult(0) = mySeq.Item(1) = 42
    myresult(1) = mySeq.Item(2) = "Hello"
    myresult(2) = mySeq.Item(3) = "3.142"
   
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

'@TestMethod("SeqC")
Private Sub Test04a_Add_MultipleItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, 42, "Hello", 3.142)
    ReDim Preserve myExpected(1 To 8)
      
    Dim myresult As Variant
   
    'Act:
    Set mySeq = SeqC.Deb(5)
    myresult = mySeq.AddItems(42, "Hello", 3.142).ToArray
    
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



''@TestMethod("SeqC")
'Private Sub Test05a_InvalidRangeItem()
'    On Error GoTo TestFail
'
'    'Arrange:
''    Dim mySeq As SeqC
''    Set mySeq = SeqC.Deb
'    Dim myExpected As Variant
'    myExpected = Array(True, True, True, False, False, False)
'
'    Dim myresult As Variant
'    ReDim myresult(0 To 5)
'
'    'Act:
'    myresult(0) = GroupInfo.InvalidRangeItem(vbNullString)
'    myresult(1) = GroupInfo.InvalidRangeItem(Array())
'    myresult(2) = GroupInfo.InvalidRangeItem(New Collection)
'    myresult(3) = GroupInfo.InvalidRangeItem("Hello")
'    myresult(4) = GroupInfo.InvalidRangeItem(Array(1, 2, 3, 4, 5))
'    Dim myC As Collection
'    Set myC = New Collection
'    myC.Add 1
'    myC.Add 2
'    myresult(5) = GroupInfo.InvalidRangeItem(myC)
'
'    'Assert:
'    Assert.SequenceEquals myExpected, myresult
'
'TestExit:
'    '@Ignore UnhandledOnErrorResumeNext
'    On Error Resume Next
'
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub


'@TestMethod("SeqC")
Private Sub Test06a_AddRange_String()

    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 10)
   
    Dim myresult As Variant
   
    'Act:
    Set mySeq = SeqC.Deb(5)
    myresult = mySeq.AddRange("Hello").ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test06b_AddRange_Array()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 10)
   
    Dim myresult As Variant
   
    'Act:
    Set mySeq = SeqC.Deb(5)
    myresult = mySeq.AddRange(Array("H", "e", "l", "l", "o")).ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test06c_AddRange_Collection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 10)
      
    Dim myresult As Variant
   
    Dim myC As Collection
    Set myC = New Collection
    With myC
        .Add "H"
        .Add "e"
        .Add "l"
        .Add "l"
        .Add "o"
    End With
    
    'Act:
    Set mySeq = SeqC(5)
    myresult = mySeq.AddRange(myC).ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test06d_AddRange_ArrayList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 10)
   
    Dim myresult As Variant
   
    Dim myAL As Object
    Set myAL = CreateObject("System.Collections.Arraylist")
    With myAL
        .Add "H"
        .Add "e"
        .Add "l"
        .Add "l"
        .Add "o"
    End With
    
    'Act:
    Set mySeq = SeqC.Deb(5)
    myresult = mySeq.AddRange(myAL).ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test06e_AddRange_Dictionary()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "Hello1", "There2", "World3")
    ReDim Preserve myExpected(1 To 8)
   
    Dim myresult As Variant
   
    Dim myD As KvpC
    Set myD = KvpC.Deb
    With myD
        .Add "Hello", 1
        .Add "There", 2
        .Add "World", 3
    End With
    
    'Act:
    Set mySeq = SeqC(5)
    myresult = mySeq.AddRange(myD).ToArray
    myresult(6) = myresult(6)(0) & VBA.CStr(myresult(6)(1))
    myresult(7) = myresult(7)(0) & VBA.CStr(myresult(7)(1))
    myresult(8) = myresult(8)(0) & VBA.CStr(myresult(8)(1))
    
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


'@TestMethod("SeqC")
Private Sub Test07a_Insert_SingleItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", Empty, 42&, Empty, 3.142, Empty)
    ReDim Preserve myExpected(1 To 8)
    Dim myExpected2 As Variant
    myExpected2 = Array(3&, 5&, 7&)
      
    Dim myresult As Variant
    Dim myResult2 As Variant
    ReDim myResult2(0 To 2)
    
    'Act:
    Set mySeq = SeqC.Deb(5)
    myResult2(0) = mySeq.Insert(3, "Hello")
    myResult2(1) = mySeq.Insert(5, 42&)
    myResult2(2) = mySeq.Insert(7, 3.142)
    
    myresult = mySeq.ToArray
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    Assert.SequenceEquals myExpected2, myResult2
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SeqC")
Private Sub Test07b_Insert_MultipleItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)
   
    Dim myresult As Variant
   
    'Act:
    Set mySeq = SeqC(5)
    mySeq.InsertItems 3, "Hello", 42&, 3.142

    myresult = mySeq.ToArray
    
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

''@TestMethod("SeqC")
'Private Sub Test09a_InsertItems()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim mySeq As SeqC
'    Dim myExpected As Variant
'    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
'    ReDim Preserve myExpected(1 To 8)
'    Dim myExpected2 As Variant
'    myExpected2 = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
'    ReDim Preserve myExpected2(1 To 8)
'    Dim myResult As Variant
'    Dim myResult2 As Variant
'
'
'    'Act:
'
'    Set mySeq = SeqC.Deb(5)
'    myResult2 = mySeq.InsertItems(3, "Hello", 42&, 3.142).ToArray
'
'
'    myResult = mySeq.ToArray
'    'Assert:
'    Assert.SequenceEquals myExpected, myResult
'    Assert.SequenceEquals myExpected2, myResult2
'
'TestExit:
'    '@Ignore UnhandledOnErrorResumeNext
'    On Error Resume Next
'
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub


'@TestMethod("SeqC")
Private Sub Test08a_InsertRange_String()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "H", "e", "l", "l", "o", Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
  
    'Act:
    Set mySeq = SeqC.Deb(5)
    mySeq.InsertRange 3, "Hello"
   
    myresult = mySeq.ToArray
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

'@TestMethod("SeqC")
Private Sub Test08b_InsertRange_Array()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)
   
    Dim myresult As Variant
    
    'Act:
    Set mySeq = SeqC.Deb(5)
    mySeq.InsertRange 3, Array("Hello", 42&, 3.142)
    
    myresult = mySeq.ToArray
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

'@TestMethod("SeqC")
Private Sub Test08c_InsertRange_Collection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)
    
    Dim myresult As Variant
   
    
    Dim myC As Collection
    Set myC = New Collection
    With myC
        .Add "Hello"
        .Add 42&
        .Add 3.142
    
    End With
    
    'Act:
    Set mySeq = SeqC(5)
    mySeq.InsertRange 3, myC
    
    myresult = mySeq.ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test08d_InsertRange_ArrayList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)
    
    Dim myresult As Variant
    
    
    Dim myAL As Object
    Set myAL = CreateObject("System.collections.arraylist")
    With myAL
        .Add "Hello"
        .Add 42&
        .Add 3.142
    
    End With
    
    'Act:
    Set mySeq = SeqC.Deb(5)
    mySeq.InsertRange 3, myAL
   
    myresult = mySeq.ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test08e_InsertRange_Dictionary()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello1", "There2", "World3", Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)
    
    Dim myresult As Variant
    
    Dim myD As KvpC
    Set myD = KvpC.Deb
    With myD
        .Add "Hello", 1
        .Add "There", 2
        .Add "World", 3
    
    End With
    
    'Act:
    Set mySeq = SeqC.Deb(5)
    mySeq.InsertRange 3, myD
    myresult = mySeq.ToArray
   
    myresult(3) = myresult(3)(0) & VBA.CStr(myresult(3)(1))
    myresult(4) = myresult(4)(0) & VBA.CStr(myresult(4)(1))
    myresult(5) = myresult(5)(0) & VBA.CStr(myresult(5)(1))
    
    
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


'@TestMethod("SeqC")
Private Sub Test09a_RemoveAt_SingleItem()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)
    
    Dim myresult As Variant
    Set mySeq = SeqC.Deb(Array(Empty, Empty, Empty, 42, Empty, Empty))
    
    'Act:
    mySeq.RemoveAt 4
  
    myresult = mySeq.ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test09b_Remove_ThreeItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)
    
    Dim myresult As Variant
    Set mySeq = SeqC.Deb(Array(Empty, 42, Empty, Empty, 42, Empty, Empty, 42))
    
    'Act:
    mySeq.RemoveAt 8, 2, 5
  
    myresult = mySeq.ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test10a_Remove_SingleItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, 42, "Hello", "Hello", Empty, Empty, 42, Empty, Empty)
    ReDim Preserve myExpected(1 To 11)
    
    Dim myresult As Variant
    Set mySeq = SeqC(Empty, 42, Empty, Empty, 42, "Hello", "Hello", "Hello", Empty, 3.142, Empty, 42, Empty, Empty)
    
    'Act:
    mySeq.Remove 42, 3.142, "Hello"
  
    myresult = mySeq.ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test11a_RemoveRange_SingleItem()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)
    
    Dim myresult As Variant
    Set mySeq = SeqC.Deb(Array(Empty, Empty, Empty, 42, Empty, Empty))
    
    'Act:
    mySeq.RemoveRange SeqC.Deb.AddItems(42)
  
    myresult = mySeq.ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test11b_RemoveRange_ThreeItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)
    
    Dim myresult As Variant
    Set mySeq = SeqC.Deb(Array(Empty, Empty, Empty, 42, 42, 42, Empty, Empty))
    
    'Act:
    mySeq.RemoveAtRange SeqC(4, 5, 6)
  
    myresult = mySeq.ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test12a_RemoveRange_SingleItem()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, 42, 42, Empty, Empty)
    ReDim Preserve myExpected(1 To 7)
    
    Dim myresult As Variant
    Set mySeq = SeqC(Empty, Empty, Empty, 42, 42, 42, Empty, Empty)
    
    'Act:
    mySeq.RemoveRange SeqC.Deb.AddItems(42)
  
    myresult = mySeq.ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test13a_RemoveAll_DefaultAll()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 0&
    
    Dim myresult As Variant
    Set mySeq = SeqC.Deb(Array(Empty, Empty, Empty, 42, 42, 42, Empty, Empty))
    
    'Act:
    mySeq.RemoveAll
    myresult = mySeq.Count
    
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


'@TestMethod("SeqC")
Private Sub Test13b_RemoveAll_Default_42AndHello()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim myExpected(1 To 5)
    
    Dim myresult As Variant
    Set mySeq = SeqC.Deb(Empty, "Hello", Empty, "Hello", "Hello", Empty, 42, 42, 42, Empty, Empty)
    
    'Act:
    mySeq.RemoveAll "Hello", 42
    myresult = mySeq.ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test13c_Reset()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 0&
    
    Dim myresult As Variant
    Set mySeq = SeqC.Deb(Array(Empty, Empty, Empty, 42, 42, 42, Empty, Empty))
    
    'Act:
    mySeq.Reset
    myresult = mySeq.Count
    
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

'@TestMethod("SeqC")
Private Sub Test13d_Clear()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 0&
    
    Dim myresult As Variant
    Set mySeq = SeqC.Deb(Array(Empty, Empty, Empty, 42, 42, 42, Empty, Empty))
    
    'Act:
    mySeq.Clear
    myresult = mySeq.Count
    
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

'@TestMethod("SeqC")
Private Sub Test14a_Fill()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(True, True, True)
    ReDim Preserve myExpected(1 To 3)
    
    Dim myresult As Variant
    ReDim myresult(1 To 3)
    Set mySeq = SeqC.Deb(Array(Empty, Empty, Empty))
    
    'Act:
    mySeq.Fill 42, 10
    myresult(1) = mySeq.Count = 13
    myresult(2) = mySeq.Item(4) = 42&
    myresult(3) = mySeq.Item(13) = 42&
    
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


'@TestMethod("SeqC")
Private Sub Test15a_Slice()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(3&, 4&, 5&)
    ReDim Preserve myExpected(1 To 3)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Slice(3, 3).ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test15b_SliceToEnd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 8)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Slice(3).ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test15c_SliceRunOnly()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&, 4&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Slice(ipRun:=4).ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test15d_Slice_Start3_End9_step2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(3&, 5&, 7&, 9&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Slice(3, 7, 2).ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test15e_Slice_Start3_End9_step2_ToCollection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(3&, 5&, 7&, 9&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant
    ReDim myresult(1 To 4)

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    Dim myC As Collection
    Set myC = mySeq.Slice(3, 7, 2).ToCollection
    myresult(1) = myC.Item(1)
    myresult(2) = myC.Item(2)
    myresult(3) = myC.Item(3)
    myresult(4) = myC.Item(4)
    
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

'@TestMethod("SeqC")
Private Sub Test15f_Slice_Start3_End9_step2_ToArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(3&, 5&, 7&, 9&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Slice(3, 7, 2).ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test16a_Head()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(1&)
    ReDim Preserve myExpected(1 To 1)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Head.ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test16b_Head_3Items()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&)
    ReDim Preserve myExpected(1 To 3)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Head(3).ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test16c_HeadZeroItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 0&
   
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Head(-2).Count
    
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

'@TestMethod("SeqC")
Private Sub Test16d_HeadFullSeqC()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Head(42).ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test17a_Tail()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 9)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Tail.ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test17b_Tail_3Items()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 7)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Tail(3).ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test17c_TailFullItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 0&
   
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Tail(42).Count
    
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

'@TestMethod("SeqC")
Private Sub Test17d_TailZeroSeqC()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    
    'Act:
    myresult = mySeq.Tail(-2).ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test18a_KnownIndexes_Available()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 9&, 10&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant
    ReDim myresult(1 To 4)

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult(1) = mySeq.FirstIndex
    myresult(2) = mySeq.FBOIndex
    myresult(3) = mySeq.LBOIndex
    myresult(4) = mySeq.LastIndex
    
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


'@TestMethod("SeqC")
Private Sub Test18b_KnownIndexes_Unavailable()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(-1&, -1&, -1&, -1&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant
    ReDim myresult(1 To 4)

    Set mySeq = SeqC.Deb
    
    'Act:
    myresult(1) = mySeq.FirstIndex
    myresult(2) = mySeq.FBOIndex
    myresult(3) = mySeq.LBOIndex
    myresult(4) = mySeq.LastIndex
    
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


'@TestMethod("SeqC")
Private Sub Test19a_KnownValues_Available()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 90&, 100&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant
    ReDim myresult(1 To 4)

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult(1) = mySeq.First
    myresult(2) = mySeq.FBO
    myresult(3) = mySeq.LBO
    myresult(4) = mySeq.Last
    
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

'@TestMethod("SeqC")
Private Sub Test20a_IndexOf_WholeSeq_Present()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.IndexOf(50&)
    
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

'@TestMethod("SeqC")
Private Sub Test20b_IndexOf_WholeSeq_NotPresent()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = -1&
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.IndexOf(55&)
    
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

'@TestMethod("SeqC")
Private Sub Test20c_IndexOf_SubSeq_Present()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myresult As Variant
    
    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.IndexOf(50&, 4, 4)
    
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

'@TestMethod("SeqC")
Private Sub Test20d_IndexOf_SubSeq_NotPresent()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = -1&
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.IndexOf(20&, 4, 4)
    
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


'@TestMethod("SeqC")
Private Sub Test21a_LastIndexOf_WholeSeq_Present()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.LastIndexOf(50&)
    
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

'@TestMethod("SeqC")
Private Sub Test21b_LastIndexOf_WholeSeq_NotPresent()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = -1&
    
    Dim myresult As Variant
    
    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.LastIndexOf(55&)
    
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

'@TestMethod("SeqC")
Private Sub Test21c_LastIndexOf_SubSeq_Present()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myresult As Variant
    
    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.LastIndexOf(50&, 4, 4)
    
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

'@TestMethod("SeqC")
Private Sub Test21d_LastIndexOf_SubSeq_NotPresent()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = -1&
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.LastIndexOf(20&, 4, 4)
    
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

'@TestMethod("SeqC")
Private Sub Test22a_Push()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&, 1000&)
    ReDim Preserve myExpected(1 To 11)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.Push(1000&).ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test22b_PushRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&, 11&, 12&, 13&, 14&, 15&)
    ReDim Preserve myExpected(1 To 15)
    
    Dim myresult As Variant
    
    Dim myArray As Variant
    myArray = Array(11&, 12&, 13&, 14&, 15&)
    Set mySeq = SeqC.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.PushRange(myArray).ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test23a_Pop()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 100&
    
    Dim myExpected2 As Variant
    myExpected2 = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&)
    ReDim Preserve myExpected2(1 To 9)
    
    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.Pop
    myResult2 = mySeq.ToArray
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    Assert.SequenceEquals myExpected2, myResult2
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqC")
Private Sub Test23b_PopRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(100&, 90&, 80&, 70&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myExpected2 As Variant
    myExpected2 = Array(10&, 20&, 30&, 50&, 50&, 60&)
    ReDim Preserve myExpected2(1 To 6)
    
    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.PopRange(4).ToArray
    myResult2 = mySeq.ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    Assert.SequenceEquals myExpected2, myResult2
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SeqC")
Private Sub Test23c_PopRange_ExceedsHost()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(100&, 90&, 80&, 70&, 60&, 50&, 40&, 30&, 20&, 10&)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myExpected2 As Variant
    myExpected2 = 0&
    
    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.PopRange(25).ToArray
    myResult2 = mySeq.Count
    
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    Assert.AreEqual myExpected2 = 0, myResult2 = 0
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqC")
Private Sub Test23d_PopRange_NegativeRun()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 0&
    
    Dim myExpected2 As Variant
    myExpected2 = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected2(1 To 10)
    
    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.PopRange(-2).Count
    myResult2 = mySeq.ToArray
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    Assert.SequenceEquals myExpected2, myResult2
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SeqC")
Private Sub Test24a_Enqueue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&, 1000&)
    ReDim Preserve myExpected(1 To 11)
    
    Dim myresult As Variant
    
    Set mySeq = SeqC.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.enQueue(1000&).ToArray
    
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

'@TestMethod("SeqC")
Private Sub Test24b_EnqueueRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&, 11&, 12&, 13&, 14&, 15&)
    ReDim Preserve myExpected(1 To 15)
    
    Dim myresult As Variant
    
    Dim myArray As Variant
    myArray = Array(11&, 12&, 13&, 14&, 15&)
    Set mySeq = SeqC.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.EnqueueRange(myArray).ToArray
    
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


'@TestMethod("SeqC")
Private Sub Test25a_Dequeue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 10&
    
    Dim myExpected2 As Variant
    myExpected2 = Array(20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected2(1 To 9)
    
    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.Dequeue
    myResult2 = mySeq.ToArray
    
    'Assert:
    Assert.AreEqual myExpected, myresult
    Assert.SequenceEquals myExpected2, myResult2
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqC")
Private Sub Test25b_DeqeueRange()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myExpected2 As Variant
    myExpected2 = Array(50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected2(1 To 6)
    
    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.DequeueRange(4).ToArray
    myResult2 = mySeq.ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    Assert.SequenceEquals myExpected2, myResult2
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SeqC")
Private Sub Test25c_DequeueRange_ExceedsHost()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myExpected2 As Variant
    myExpected2 = 0&
    
    
    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.DequeueRange(25).ToArray
    myResult2 = mySeq.Count
    
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    Assert.AreEqual myExpected2 = 0, myResult2 = 0
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SeqC")
Private Sub Test26a_Sort()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 10)

    Dim myresult As Variant

    Set mySeq = SeqC.Deb(30&, 70&, 40&, 50&, 60&, 80&, 20&, 90&, 10&, 100&)
    
    'Act:
    myresult = mySeq.Sort.ToArray
  
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

'@TestMethod("SeqC")
Private Sub Test26b_Sorted()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myExpected2 As Variant
    myExpected2 = Array(30&, 70&, 40&, 50&, 60&, 80&, 20&, 90&, 10&, 100&)
    ReDim Preserve myExpected2(1 To 10)
    
    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqC.Deb(30&, 70&, 40&, 50&, 60&, 80&, 20&, 90&, 10&, 100&)
    
    'Act:
    myresult = mySeq.Sorted.ToArray
    myResult2 = mySeq.ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    Assert.SequenceEquals myExpected2, myResult2
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SeqC")
Private Sub Test27a_Reverse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(100&, 90&, 80&, 70&, 60&, 50&, 40&, 30&, 20&, 10&)
    ReDim Preserve myExpected(1 To 10)
    
    Dim mySeq As SeqC
    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
     Dim myresult As Variant
    'Act:
    myresult = mySeq.Reverse.ToArray
  
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

'@TestMethod("SeqC")
Private Sub Test27b_Reversed()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(100&, 90&, 80&, 70&, 60&, 50&, 40&, 30&, 20&, 10&)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myExpected2 As Variant
    myExpected2 = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected2(1 To 10)
    
    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqC.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    
    'Act:
    myresult = mySeq.Reverse.ToArray
    myResult2 = mySeq.Reversed.ToArray
    
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    Assert.SequenceEquals myExpected2, myResult2
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("SeqC")
Private Sub Test28a_Unique()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    
    Set mySeq = SeqC.Deb(10&, 100&, 20&, 30&, 40&, 50&, 30&, 30&, 60&, 100&, 70&, 100&, 80&, 90&, 100&)
    
    'Act:
    ' The array needs to be sorted because unique copies the first item encountered
    myresult = mySeq.Unique.Sorted.ToArray
   
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

'@TestMethod("SeqC")
Private Sub Test28b_Unique_SIngleItem()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&)
    ReDim Preserve myExpected(1 To 1)
    
    Dim myresult As Variant
    
    Set mySeq = SeqC.Deb.AddItems(10&)
    
    'Act:
    ' The array needs to be sorted because unique copies the first item encountered
    myresult = mySeq.Unique.Sorted.ToArray
   
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

'@TestMethod("SeqC")
Private Sub Test28c_Unique_NoItems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = 0&
    
    Dim myresult As Variant
   
    Set mySeq = SeqC.Deb
    
    'Act:
    ' The array needs to be sorted because unique copies the first item encountered
    myresult = mySeq.Count
   
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

'@TestMethod("SeqC")
Private Sub Test29a_SetOfConnom()

    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 6)
    
    Dim myresult As Variant
    
    Set mySeq = SeqC.Deb(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)
    
    'Act:
    myresult = mySeq.SetOf(m_Common, SeqC(Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&))).ToArray
   
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


'@TestMethod("SeqC")
Private Sub Test29b_SetOfHostOnly()

    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(110&, 120&, 130&, 140&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)
    
    'Act:
    myresult = mySeq.SetOf(m_HostOnly, SeqC(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)).ToArray
   
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

'@TestMethod("SeqC")
Private Sub Test29c_SetOfParamOnly()

    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant

    Set mySeq = SeqC.Deb(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)
    
    'Act:
    myresult = mySeq.SetOf(m_ParamOnly, SeqC(Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&))).ToArray
   
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

'@TestMethod("SeqC")
Private Sub Test29d_SetOfNotCommon()

    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 110&, 120&, 130&, 140&)
    ReDim Preserve myExpected(1 To 8)
    
    Dim myresult As Variant
    
    Set mySeq = SeqC.Deb(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = mySeq.SetOf(m_NotCommon, SeqC(Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&))).Sorted.ToArray
   
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

'@TestMethod("SeqC")
Private Sub Test29e_SetOfUnique()

    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)
    ReDim Preserve myExpected(1 To 14)
    
    Dim myresult As Variant
    
    Set mySeq = SeqC.Deb(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = mySeq.SetOf(m_Unique, SeqC(Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&))).Sorted.ToArray
   
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



'@TestMethod("SeqC")
Private Sub Test30a_Swap()

    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqC
    Dim myExpected As Variant
    myExpected = Array(140&, 130&, 120&, 110&, 100&, 90&, 80&, 70&, 60&, 50&)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    
    Set mySeq = SeqC.Deb(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)
    
    'Act:
    mySeq.Swap 1, 10
    mySeq.Swap 2, 9
    mySeq.Swap 3, 8
    mySeq.Swap 4, 7
    mySeq.Swap 5, 6
    
    myresult = mySeq.ToArray
    
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



