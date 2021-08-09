Attribute VB_Name = "TestLyst"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
'Private Fakes As Rubberduck.FakesProvider


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
'    Set Fakes = New Rubberduck.FakesProvider
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
'    Set Fakes = Nothing
End Sub


''@TestInitialize
'Private Sub TestInitialize()
'    'This method runs before every test in the module..
'End Sub
'
'
''@TestCleanup
'Private Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub


'@TestMethod("Lyst")
Private Sub Test01_NewLystIsObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb
    
    Dim myResult As Boolean
    
    'Act:
    myResult = VBA.IsObject(myLyst)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Lyst")
Private Sub Test02_NewLystIsLystObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "Lyst"
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb
    
    Dim myResult As String
    
    'Act:
    myResult = VBA.TypeName(myLyst)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Count")
Private Sub Test03_NewLystCountIsZero()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 0
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb
    
    Dim myResult As Long
    
    'Act:
    myResult = myLyst.Count
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Add.Count")
Private Sub Test04_AddFiveItemsCountIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 5
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb
    
    myLyst.Add 10
    myLyst.Add 20
    myLyst.Add 30
    myLyst.Add 40
    myLyst.Add 50
    
    Dim myResult As Long
    
    'Act:
    myResult = myLyst.Count
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("AddRange.Count")
Private Sub Test05_AddRangeArrayOfFiveFilledIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 5
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb
    
    myLyst.addrange Array(10, 20, 30, 40, 50)
    
    Dim myResult As Long
    
    'Act:
    myResult = myLyst.Count
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Add.Count")
Private Sub Test06_AddByDebArrayOfFiveCountIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 5
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult As Long
    
    'Act:
    myResult = myLyst.Count
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("AddRange.Count")
Private Sub Test07_AddRangeStackOfFiveCountIsFive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 5
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb
    
    Dim myStack As Stack
    Set myStack = New Stack
    With myStack
        .Push 10
        .Push 20
        .Push 30
        .Push 40
        .Push 50
    End With
    
    myLyst.addrange myStack
    
    Dim myResult As Long
    
    'Act:
    myResult = myLyst.Count
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Item")
Private Sub Test07a_AssignItemTwoPrimitive()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 300
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb.addrange(Array(10, 20, 30, 40, 50))
    
    Dim myResult As Long
    
    'Act:
    myLyst.Item(2) = 300
    myResult = myLyst.Item(2)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Item")
Private Sub Test07b_AssignItemTwoObject()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As lyst
    Set myExpected = lyst.deb.addrange(Array(100, 200, 300, 400, 500))
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb.addrange(Array(10, 20, 30, 40, 50))
    
    Dim myResult As lyst
    
    'Act:
    myLyst.Item(2) = lyst.deb.addrange(Array(100, 200, 300, 400, 500))
    Set myResult = myLyst.Item(2)
   
    'Assert:
    Assert.SequenceEquals myExpected.ToArraylist.toarray, myResult.ToArraylist.toarray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Clear")
Private Sub Test08_Clear()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 0
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult As Long
    
    'Act:
    If myLyst.Count = 5 Then myLyst.Clear
    myResult = myLyst.Count
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Clone/ToArray")
Private Sub Test09_Clone()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 40, 50)
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult As lyst
    
    'Act:
    Set myResult = myLyst.Clone
   
    'Assert:
    Assert.SequenceEquals myExpected, myResult.toarray

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("HoldsValue")
Private Sub Test10_HoldsValueTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult As Boolean
    
    'Act:
    myResult = myLyst.HoldsValue(10)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("HoldsValue")
Private Sub Test11_HoldsValueFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult As Boolean
    
    'Act:
    myResult = myLyst.HoldsValue(100)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("LacksValue")
Private Sub Test12_LacksValueTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult As Boolean
    
    'Act:
    myResult = myLyst.LacksValue(100)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("LacksValue")
Private Sub Test13_LacksValueFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult As Boolean
    
    'Act:
    myResult = myLyst.LacksValue(10)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("CopyTo")
Private Sub Test14_CopyToAllArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 40, 50)
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult(0 To 4) As Variant
    
    'Act:
    myLyst.CopyTo myResult
   
    'Assert:
    Assert.SequenceEquals myExpected, myResult
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("CopyTo")
Private Sub Test15_CopyToItem2ToEnd()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(30, 40, 50)
   
    
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult(0 To 2) As Variant
    
    'Act:
    myLyst.CopyTo myResult, 2
   
    'Assert:
    Assert.SequenceEquals myExpected, myResult
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("CopyTo")
Private Sub Test16_CopyToItem1toItem3()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(20, 30, 40)
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult(0 To 2) As Variant
    
    'Act:
    myLyst.CopyTo myResult, 1, 3
   
    'Assert:
    Assert.SequenceEquals myExpected, myResult
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("GetRange")
Private Sub Test17_GetRangeItem1ToItem3()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(20, 30, 40)
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult As lyst
    
    'Act:
    Set myResult = myLyst.GetRange(1, 3)
   
    'Assert:
    Assert.SequenceEquals myExpected, myResult.toarray
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("IndexOf")
Private Sub Test18_IndexOfWholeList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 2
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult As Long
    
    'Act:
     myResult = myLyst.IndexOf(30)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("IndexOf")
Private Sub Test19_IndexOfFromItem1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 2
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult As Long
    
    'Act:
     myResult = myLyst.IndexOf(30, 1)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("InsertAt")
Private Sub Test20_InsertAtItem1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 70, 30, 40, 50)
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult As lyst
    
    'Act:
     Set myResult = myLyst.InsertAt(2, 70)
   
    'Assert:
    Assert.SequenceEquals myExpected, myResult.ToArraylist.toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("InsertRange")
Private Sub Test21_InsertRangeFivetemsFromItem1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 15, 16, 17, 18, 19, 30, 40, 50)
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 50))
    
    Dim myResult As lyst
    
    'Act:
     Set myResult = myLyst.InsertRange(2, Array(15, 16, 17, 18, 19))
   
    'Assert:
    Assert.SequenceEquals myExpected, myResult.ToArraylist.toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("LastIndexOf")
Private Sub Test22_LastIndexOfWholeLyst()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 6
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 40, 40, 40, 50))
    
    Dim myResult As Long
    
    'Act:
    myResult = myLyst.LastIndexOf(40)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("LastIndexOf")
Private Sub Test23_LastIndexOfStartItem4()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 6
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 40, 40, 40, 50))
    
    Dim myResult As Long
    
    'Act:
    myResult = myLyst.LastIndexOf(40, 4)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("LastIndexOf")
Private Sub Test24_LastIndexOfStartItem1EndItem4()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 4
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 40, 40, 40, 50))
    
    Dim myResult As Long
    
    'Act:
    myResult = myLyst.LastIndexOf(40, 1, 4)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("RemoveValue")
Private Sub Test25_RemoveValueOf40()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As lyst
    Set myExpected = lyst.deb.addrange(Array(10, 20, 30, 40, 40, 40, 50))
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 40, 40, 40, 50))
    
    Dim myResult As lyst
    
    'Act:
    Set myResult = myLyst.removevalue(40)
   
    'Assert:
    Assert.SequenceEquals myExpected.ToArraylist.toarray, myResult.ToArraylist.toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("RemoveAt")
Private Sub Test26_RemoveAtValueOf20FromItem4()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As lyst
    Set myExpected = lyst.deb.addrange(Array(10, 20, 30, 40, 40, 40, 50))
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 20, 40, 40, 50))
    
    Dim myResult As lyst
    
    'Act:
    Set myResult = myLyst.RemoveAt(4)
   
    'Assert:
    Assert.SequenceEquals myExpected.ToArraylist.toarray, myResult.ToArraylist.toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("RemoveRange")
Private Sub Test27_RemoveRangeItem3Count4()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As lyst
    Set myExpected = lyst.deb.addrange(Array(10, 20, 30, 50))
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 40, 40, 40, 50))
    
    Dim myResult As lyst
    
    'Act:
    Set myResult = myLyst.RemoveRange(3, 4)
   
    'Assert:
    Assert.SequenceEquals myExpected.ToArraylist.toarray, myResult.ToArraylist.toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Reverse")
Private Sub Test28_ReverseAll()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As lyst
    Set myExpected = lyst.deb.addrange(Array(50, 40, 40, 40, 40, 30, 20, 10))
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 40, 40, 40, 50))
    
    Dim myResult As lyst
    
    'Act:
    Set myResult = myLyst.Reverse
   
    'Assert:
    Assert.SequenceEquals myExpected.ToArraylist.toarray, myResult.ToArraylist.toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Reverse")
Private Sub Test29_ReverseItem1Count4()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As lyst
    Set myExpected = lyst.deb.addrange(Array(10, 40, 40, 30, 20, 40, 40, 50))
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 40, 40, 40, 50))
    
    Dim myResult As lyst
    
    'Act:
    Set myResult = myLyst.Reverse(1, 4)
   
    'Assert:
    Assert.SequenceEquals myExpected.ToArraylist.toarray, myResult.ToArraylist.toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Reverse")
Private Sub Test30_SetRangeItem1ToFouritems()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As lyst
    Set myExpected = lyst.deb.addrange(Array(10, 50, 50, 50, 50, 40, 40, 50))
   
    '@Ignore IntegerDataType
    Dim myLyst As lyst
    Set myLyst = lyst.deb(Array(10, 20, 30, 40, 40, 40, 40, 50))
    
    Dim myResult As lyst
    
    'Act:
    Set myResult = myLyst.SetRange(1, Array(50, 50, 50, 50))
   
    'Assert:
    Assert.SequenceEquals myExpected.ToArraylist.toarray, myResult.ToArraylist.toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub
