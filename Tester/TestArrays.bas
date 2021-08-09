Attribute VB_Name = "TestArrays"
'@IgnoreModule UnassignedVariableUsage
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


Private Function MakeRowColArray(ByVal ipRows As Long, ByVal ipCols As Long) As Variant
   
    '@Ignore VariableNotAssigned
    Dim myArray As Variant
    ReDim myArray(1 To ipRows, 1 To ipCols)
    Dim myValue As Long
    myValue = 1
    
    Dim myRow As Long
    For myRow = 1 To ipRows
    
        Dim myCol As Long
        For myCol = 1 To ipCols
        
            myArray(myRow, myCol) = myValue
            myValue = myValue + 1
            
        Next
        
    Next
        
    MakeRowColArray = myArray
    
End Function


Private Function MakeColRowArray(ByVal ipRows As Long, ByVal ipCols As Long) As Variant
   
    '@Ignore VariableNotAssigned
    Dim myArray As Variant
    ReDim myArray(1 To ipRows, 1 To ipCols)
    Dim myValue As Long
    myValue = 1
    
    Dim myCol As Long
    For myCol = 1 To ipCols
    
        Dim myRow As Long
        For myRow = 1 To ipRows
        
            myArray(myRow, myCol) = myValue
            myValue = myValue + 1
            
        Next
        
    Next
        
    MakeColRowArray = myArray
    
End Function


'@TestMethod("Ranks")
Private Sub Test01_RanksEmptyArrayIsZero()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 0
    
    Dim myArray() As Long
    Dim myResult As Long
    
    'Act:
    myResult = Types.Arrays.Ranks(myArray).Value
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Ranks")
Private Sub Test02_Ranks1DArrayIs1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 1
    
    Dim myArray(1 To 10) As Long
    Dim myResult As Long
    
    'Act:
    myResult = Types.Arrays.Ranks(myArray).Value
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Ranks")
Private Sub Test03_Ranks2DArrayIs2()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 2
    
    Dim myArray(1 To 10, 1 To 10) As Long
    Dim myResult As Long
    
    'Act:
    myResult = Types.Arrays.Ranks(myArray).Value
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Ranks")
Private Sub Test04_Ranks3DArrayIs3()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 3
    
    Dim myArray(1 To 10, 1 To 10, 1 To 10) As Long
    Dim myResult As Long
    
    'Act:
    myResult = Types.Arrays.Ranks(myArray).Value
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("HasRank")
Private Sub Test05_HaRankEmptyArrayFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.HasRank(myArray, 2)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("HasRank")
Private Sub Test06_HaRank1DArrayFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 10) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.HasRank(myArray, 2)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("HasRank")
Private Sub Test07_HaRank3DArrayTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 10, 1 To 10, 1 To 10) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.HasRank(myArray, 2)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("HasRank")
Private Sub Test08_HaRank3DArrayNegativeBoundsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(-10 To -1, -10 To -1, -10 To -1) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.HasRank(myArray, 2)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("HasRank")
Private Sub Test09_HaRank3DArraySingleItemNegativeBoundsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(-1 To -1, -1 To -1, -1 To -1) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.HasRank(myArray, 2)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("IsId")
Private Sub Test10_Is1DEmptyArrayFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.IsListArray(myArray)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("IsId")
Private Sub Test11_Is1D1DArrayTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 10) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.IsListArray(myArray)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Is2d")
Private Sub Test12_Is1D2DArrayFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 10, 1 To 10) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.IsListArray(myArray)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Is2d")
Private Sub Test13_Is2DEmptyArrayFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.IsTableArray(myArray)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Is2d")
Private Sub Test14_Is2D1DArrayFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 10) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.IsTableArray(myArray)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Is2d")
Private Sub Test15_Is2D2DArrayTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 10, 1 To 10) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.IsTableArray(myArray)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("IsMd")
Private Sub Test16_IsMDEmptyArrayFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray() As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.IsMDArray(myArray)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("IsMd")
Private Sub Test17_IsMD1DArrayFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 10) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.IsMDArray(myArray)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("IsMd")
Private Sub Test18_IsMD2DArrayFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myArray(1 To 10, 1 To 10) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.IsMDArray(myArray)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("IsMd")
Private Sub Test19_IsMD3DArrayTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myArray(1 To 10, 1 To 10, 1 To 10) As Long
    Dim myResult As Boolean
    
    'Act:
    myResult = Types.Arrays.IsMDArray(myArray)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("TryGetSize")
Private Sub Test20_TryGetSizeEmptyArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedStatus  As Boolean
    myExpectedStatus = False
    
    Dim myExpectedValue As Long
    myExpectedValue = -1
    
    Dim myArray() As Long
    Dim myIRL As resultlong
    'Set myIRL = ResultLong.deb
    
    Dim myResultStatus As Boolean
    Dim myResultValue As Long
    
    'Act:
    myResultStatus = Types.Arrays.TryGetSize(myArray, 1, myIRL).Status
    myResultValue = myIRL.Value
    'Assert:
    Assert.AreEqual myExpectedStatus, myResultStatus, "Status"
    Assert.AreEqual myExpectedValue, myResultValue, "Value"
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("TryGetSize")
Private Sub Test21_TryGetSizeRank2Is10()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedStatus  As Boolean
    myExpectedStatus = True
    
    Dim myExpectedValue As Long
    myExpectedValue = 10
    
    Dim myArray(1 To 10, 1 To 10, 1 To 10) As Long
    Dim myIRL As resultlong
    Set myIRL = resultlong.deb
    
    Dim myResultStatus As Boolean
    Dim myResultValue As Long
    
    'Act:
    myResultStatus = Types.Arrays.TryGetSize(myArray, 2, myIRL).Status
    myResultValue = myIRL.Value
    'Assert:
    Assert.AreEqual myExpectedStatus, myResultStatus, "Status"
    Assert.AreEqual myExpectedValue, myResultValue, "Value"
    
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("TryGetLBound")
Private Sub Test22_TryGetLBoundEmptyArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedStatus  As Boolean
    myExpectedStatus = False
    
    Dim myExpectedValue As Long
    myExpectedValue = -1
    
    Dim myArray() As Long
    'Dim myResult As Boolean
    Dim myIRL As resultlong
    Set myIRL = resultlong.deb
    
    Dim myResultStatus As Boolean
    Dim myResultValue As Long
    
    'Act:
    myResultStatus = Types.Arrays.TryGetLbound(myArray, 1, myIRL).Status
    myResultValue = myIRL.Value
    'Assert:
    Assert.AreEqual myExpectedStatus, myResultStatus, "Status"
    Assert.AreEqual myExpectedValue, myResultValue, "Value"
    
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("TryGetLBound")
Private Sub Test23_TryGetLBoundLboundIs5()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedStatus  As Boolean
    myExpectedStatus = True
    
    Dim myExpectedValue As Long
    myExpectedValue = 5
    
    Dim myArray(5 To 10, 5 To 10, 5 To 10) As Long
    'Dim myResult As Boolean
    Dim myIRL As resultlong
    Set myIRL = resultlong.deb
    
    Dim myResultStatus As Boolean
    Dim myResultValue As Long
    
    'Act:
    myResultStatus = Types.Arrays.TryGetLbound(myArray, 2, myIRL).Status
    myResultValue = myIRL.Value
    'Assert:
    Assert.AreEqual myExpectedStatus, myResultStatus, "Status"
    Assert.AreEqual myExpectedValue, myResultValue, "Value"
    
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("TryGetUBound")
Private Sub Test24_TryGetUBoundEmptyArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedStatus  As Boolean
    myExpectedStatus = False
    
    Dim myExpectedValue As Long
    myExpectedValue = -1
    
    Dim myArray() As Long
    'Dim myResult As Boolean
    Dim myIRL As resultlong
    Set myIRL = resultlong.deb
    
    Dim myResultStatus As Boolean
    Dim myResultValue As Long
    
    'Act:
    myResultStatus = Types.Arrays.TryGetUbound(myArray, 1, myIRL).Status
    myResultValue = myIRL.Value
    'Assert:
    Assert.AreEqual myExpectedStatus, myResultStatus, "Status"
    Assert.AreEqual myExpectedValue, myResultValue, "Value"
    
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("TryGetUBound")
Private Sub Test25_TryGetUBoundUboundIs10()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedStatus  As Boolean
    myExpectedStatus = True
    
    Dim myExpectedValue As Long
    myExpectedValue = 10
    
    Dim myArray(5 To 10, 5 To 10, 5 To 10) As Long
    'Dim myResult As Boolean
    Dim myIRL As resultlong
    Set myIRL = resultlong.deb
    
    Dim myResultStatus As Boolean
    Dim myResultValue As Long
    
    'Act:
    myResultStatus = Types.Arrays.TryGetUbound(myArray, 2, myIRL).Status
    myResultValue = myIRL.Value
    'Assert:
    Assert.AreEqual myExpectedStatus, myResultStatus, "Status"
    Assert.AreEqual myExpectedValue, myResultValue, "Value"
    
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub



'@TestMethod("TryRotate")
Private Sub Test26_Transposearray()
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpectedStatus As Boolean
    myExpectedStatus = True
    
    Dim myExpectedArray As Variant
    myExpectedArray = MakeColRowArray(5, 4)
    
    Dim mySource As Variant
    mySource = MakeRowColArray(4, 5)
    
   
    Dim myResult As resultvariant
    
    
    'Act:
    Set myResult = Types.Arrays.TryTranspose(mySource)
    
    'Assert:
    Assert.AreEqual myExpectedStatus, myResult.Status, "Status"
    Assert.SequenceEquals myExpectedArray, myResult.Value, "Value"
    
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("TryToLystOfLyst")
Public Sub Test27_ArrayToLystOfLystsByRow()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedStatus As Boolean
    myExpectedStatus = True
    
    Dim myExpectedLyst As lyst
    Set myExpectedLyst = lyst.deb
    
    With myExpectedLyst
    
        .Add lyst.deb(Array(1&, 2&, 3&, 4&))
        .Add lyst.deb(Array(5&, 6&, 7&, 8&))
        .Add lyst.deb(Array(9&, 10&, 11&, 12&))
        .Add lyst.deb(Array(13&, 14&, 15&, 16&))
        .Add lyst.deb(Array(17&, 18&, 19&, 20&))
        
    End With
    
    Dim mySource As Variant
    mySource = MakeRowColArray(5, 4)
    
    Dim myResult As resultlyst
    
    'Act:
    Set myResult = Types.Arrays.TryToLystOfLyst(mySource, myResult, ByRow)
    
    'Assert:
    Assert.AreEqual myExpectedStatus, myResult.Status
    Assert.SequenceEquals myExpectedLyst.Item(2).toarray, myResult.Value.Item(2).toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TryToLystOfLyst")
Public Sub Test28_ArrayToLystOfLystsByCol()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedStatus As Boolean
    myExpectedStatus = True
    
    Dim myExpectedLyst As lyst
    Set myExpectedLyst = lyst.deb
    
    With myExpectedLyst
        
        .Add lyst.deb(Array(1&, 5&, 9&, 13&, 17&))
        .Add lyst.deb(Array(2&, 6&, 10&, 14&, 18&))
        .Add lyst.deb(Array(3&, 7&, 11&, 15&, 19&))
        .Add lyst.deb(Array(4&, 8&, 12&, 16&, 20&))
        
    End With
    
    Dim mySource As Variant
    mySource = MakeRowColArray(5, 4)
    
    Dim myResult As resultlyst
    
    'Act:
    Set myResult = Types.Arrays.TryToLystOfLyst(mySource, myResult, KeyOrientation.ByColumn)
    
    'Assert:
    Assert.AreEqual myExpectedStatus, myResult.Status
    Assert.SequenceEquals myExpectedLyst.Item(2).toarray, myResult.Value.Item(2).toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TryToLystOfLyst")
Public Sub Test29_ArrayToLystOfLystsSplitFirstByRow()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedStatus As Boolean
    myExpectedStatus = True
    
    Dim myExpectedLyst As lyst
    Set myExpectedLyst = lyst.deb
    
    With myExpectedLyst
    
        Dim myLyst As lyst
        
        Set myLyst = lyst.deb
        myLyst.Add 1&
        myLyst.Add lyst.deb(Array(2&, 3&, 4&))
        .Add myLyst
        
        Set myLyst = lyst.deb
        myLyst.Add 5&
        myLyst.Add lyst.deb(Array(6&, 7&, 8&))
        .Add myLyst
        
        Set myLyst = lyst.deb
        myLyst.Add 9&
        myLyst.Add lyst.deb(Array(10&, 11&, 12&))
        .Add myLyst
       
        
        Set myLyst = lyst.deb
        myLyst.Add 13&
        myLyst.Add lyst.deb(Array(14&, 15&, 16&))
        .Add myLyst
        
        
        Set myLyst = lyst.deb
        myLyst.Add 17&
        myLyst.Add lyst.deb(Array(18&, 19&, 20&))
        .Add myLyst
        
    End With
    
    Dim mySource As Variant
    mySource = MakeRowColArray(5, 4)
    
    Dim myResult As resultlyst
    
    'Act:
    Set myResult = Types.Arrays.TryToLystOfLystSplitFirst(mySource, myResult, KeyOrientation.ByRow)
    
    'Assert:
    Assert.AreEqual myExpectedStatus, myResult.Status
    Assert.SequenceEquals myExpectedLyst.Item(2).Item(1).toarray, myResult.Value.Item(2).Item(1).toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("TryToLystOfLyst")
Public Sub Test30_ArrayToLystOfLystsSplitFirstByCol()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedStatus As Boolean
    myExpectedStatus = True
    
    Dim myExpectedLyst As lyst
    Set myExpectedLyst = lyst.deb
    
    With myExpectedLyst
    
        Dim myLyst As lyst
        
        Set myLyst = lyst.deb
        myLyst.Add 1&
        myLyst.Add lyst.deb(Array(5&, 9&, 13&, 17&))
        .Add myLyst
        
        Set myLyst = lyst.deb
        myLyst.Add 2&
        myLyst.Add lyst.deb(Array(6&, 10&, 14&, 18&))
        .Add myLyst
        
        Set myLyst = lyst.deb
        myLyst.Add 3&
        myLyst.Add lyst.deb(Array(7&, 11&, 15&, 19&))
        .Add myLyst
       
        
        Set myLyst = lyst.deb
        myLyst.Add 4&
        myLyst.Add lyst.deb(Array(8&, 12&, 16&, 20&))
        .Add myLyst
        
        
'        Set myLyst = lyst.deb
'        myLyst.Add 17&
'        myLyst.Add lyst.deb(Array(18&, 19&, 20&))
'        .Add myLyst
        
    End With
    
    Dim mySource As Variant
    mySource = MakeRowColArray(5, 4)
    
    Dim myResult As resultlyst
    
    'Act:
    
    Set myResult = Types.Arrays.TryToLystOfLystSplitFirst(mySource, myResult, KeyOrientation.ByColumn)
    
    'Assert:
    Assert.AreEqual myExpectedStatus, myResult.Status
    Assert.SequenceEquals myExpectedLyst.Item(2).Item(1).toarray, myResult.Value.Item(2).Item(1).toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TryToLystOfLyst")
Public Sub Test31_ArrayToLystOfLystsCopyFirstByRow()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedStatus As Boolean
    myExpectedStatus = True
    
    Dim myExpectedLyst As lyst
    Set myExpectedLyst = lyst.deb
    
    With myExpectedLyst
    
        Dim myLyst As lyst
        
        Set myLyst = lyst.deb
        myLyst.Add 1&
        myLyst.Add lyst.deb(Array(1&, 2&, 3&, 4&))
        .Add myLyst
        
        Set myLyst = lyst.deb
        myLyst.Add 5&
        myLyst.Add lyst.deb(Array(5&, 6&, 7&, 8&))
        .Add myLyst
        
        Set myLyst = lyst.deb
        myLyst.Add 9&
        myLyst.Add lyst.deb(Array(9&, 10&, 11&, 12&))
        .Add myLyst
       
        
        Set myLyst = lyst.deb
        myLyst.Add 13&
        myLyst.Add lyst.deb(Array(13&, 14&, 15&, 16&))
        .Add myLyst
        
        
        Set myLyst = lyst.deb
        myLyst.Add 17&
        myLyst.Add lyst.deb(Array(17&, 18&, 19&, 20&))
        .Add myLyst
        
    End With
    
    Dim mySource As Variant
    mySource = MakeRowColArray(5, 4)
    
    Dim myResult As resultlyst
    
    'Act:
    Set myResult = Types.Arrays.TryToLystOfLystCopyFirst(mySource, myResult, KeyOrientation.ByRow)
    
    'Assert:
    Assert.AreEqual myExpectedStatus, myResult.Status
    Assert.SequenceEquals myExpectedLyst.Item(2).Item(1).toarray, myResult.Value.Item(2).Item(1).toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("TryToLystOfLyst")
Public Sub Test32_ArrayToLystOfLystsSplitFirstByCol()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpectedStatus As Boolean
    myExpectedStatus = True
    
    Dim myExpectedLyst As lyst
    Set myExpectedLyst = lyst.deb
    
    With myExpectedLyst
    
        Dim myLyst As lyst
        
        Set myLyst = lyst.deb
        myLyst.Add 1&
        myLyst.Add lyst.deb(Array(1&, 5&, 9&, 13&, 17&))
        .Add myLyst
        
        Set myLyst = lyst.deb
        myLyst.Add 2&
        myLyst.Add lyst.deb(Array(2&, 6&, 10&, 14&, 18&))
        .Add myLyst
        
        Set myLyst = lyst.deb
        myLyst.Add 3&
        myLyst.Add lyst.deb(Array(3&, 7&, 11&, 15&, 19&))
        .Add myLyst
       
        
        Set myLyst = lyst.deb
        myLyst.Add 4&
        myLyst.Add lyst.deb(Array(4&, 8&, 12&, 16&, 20&))
        .Add myLyst
        
        
'        Set myLyst = lyst.deb
'        myLyst.Add 17&
'        myLyst.Add lyst.deb(Array(18&, 19&, 20&))
'        .Add myLyst
        
    End With
    
    Dim mySource As Variant
    mySource = MakeRowColArray(5, 4)
    
    Dim myResult As resultlyst
    
    'Act:
    Set myResult = Types.Arrays.TryToLystOfLystCopyFirst(mySource, myResult, KeyOrientation.ByColumn)
    
    'Assert:
    Assert.AreEqual myExpectedStatus, myResult.Status
    Assert.SequenceEquals myExpectedLyst.Item(2).Item(1).toarray, myResult.Value.Item(2).Item(1).toarray
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

