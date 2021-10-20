Attribute VB_Name = "TestArrays"
''@IgnoreModule
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

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

    '@TestModule
    '@Folder("Tests")
       Public Sub ArraysTests()

          ' Test33_pvParseParamArrayToLystIsOkay_ArrayOf5Primitive

           Test01_RanksEmptyArrayIsZero
           Test02_ListArrayHasOneDimension
           Test03_TableArrayHasTwoDImensions
           Test04_MDArrayHasMoreThanTwoDimensions
           Test05_HaRankEmptyArrayFalse
           Test06_HaRank1DArrayFalse
           Test07_HaRank3DArrayTrue
           Test08_HaRank3DArrayNegativeBoundsTrue
           Test09_HaRank3DArraySingleItemNegativeBoundsTrue
           Test10_IsListArrayOnEmptyArraIsFalse
           Test11_IsListArrayForDefinedArrayIsTrue
           Test12_IsListArrayForTableArrayIsFalse
           Test13_IsTableArrayOnEmptyArrayIsFalse
           Test14_IsTableArrayOnListArrayIsFalse
           Test15_IsTableArrayOnTableArrayIsTrue
           Test16_IsMDArrayOnEmptyArrayIsFalse
           Test17_IsMDArayOnListArrayIsFalse
           Test18_IsMDArrayOnTableArrayIsFalse
           Test19_IsMDArrayOnMDArrayIsTrue
           Test20_CountOnEmptyArrayIsMinusOne
           Test21_TableArrayRankTwoHasCountOfTen
           Test22_TryExtentOnEmptyArrayHasStatusOfFalse
           Test23_TryExtentOnFilledMDArrayRankTwoHasFiveTenSix
           'Test24_TryGetUBoundEmptyArray
          ' Test25_TryGetUBoundUboundIs10
           Test26_TransposeArray
           Test27_ArrayToLystOfLystsByRow
           Test28_ArrayToLystOfLystsByCol
           Test29_ArrayToLystOfLystsByRowFirstItemActionIsSplitFirstRow
           Test30_ArrayToLystOfLystsByColFirstItemActionIsSplitFirstItem
           Test31_ArrayToLystOfLystsByRowFirstItemActionIsCopyFirstItem
           Test32_ArrayToLystOfLystsByColFirstActionItemIsSPlitFirstItem

           Debug.Print CurrentComponentName & vbTab & vbTab & vbTab & "testing completed"

       End Sub
        
        '@IgnoreModule UnassignedVariableUsage

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
    
    Private Function GetParamArray(ParamArray ipArgs() As Variant) As Variant
        GetParamArray = ipArgs
    End Function

    ' '@TestMethod("Ranks")
    ' Private Sub Test33_pvParseParamArrayToLystIsOkay_ArrayOf5Primitive()
    '     On Error GoTo TestFail
        
    '     'Arrange:
    '     Dim myExpected  As Variant
    '     myExpected = Array(1, 2, 3, 4, 5)
       
        
    '     Dim myExpectedCode As ParseResultCode
    '     myExpectedCode = ParseResultCode.IsSingleLyst
        
    '     Dim myResult As ParseResult
    '    ' Dim myResultState As Boolean
        
    '     'Act:
    '     Set myResult = Arrays.ParseVariantUsingSingleItemSpecialCase(GetParamArray(1, 2, 3, 4, 5))
    
    '     'Assert:
    '     Assert.AreEqual myExpectedCode, myResult.Code  ', CurrentProcedureName
    '     Assert.SequenceEquals myExpected, myResult.Items.ToArray  ', CurrentProcedureName
    ' TestExit:
    '     Exit Sub
        
    ' TestFail:
    '     Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
    '     Resume TestExit
        
    ' End Sub

   '@TestMethod("Ranks")
    Private Sub Test01_RanksEmptyArrayIsZero()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Long
        myExpected = 0
        
        Dim myArray() As Long
        Dim myResult As Long
        
        'Act:
        myResult = Arrays.Ranks(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test01_RanksEmptyArrayIsZero raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("Ranks")
    Private Sub Test02_ListArrayHasOneDimension()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Long
        myExpected = 1
        
        Dim myArray(1 To 10) As Long
        Dim myResult As Long
        
        'Act:
        myResult = Arrays.Ranks(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test02_Ranks1DArrayIs1 raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("Ranks")
    Private Sub Test03_TableArrayHasTwoDImensions()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Long
        myExpected = 2
        
        Dim myArray(1 To 10, 1 To 10) As Long
        Dim myResult As Long
        
        'Act:
        myResult = Arrays.Ranks(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("Ranks")
    Private Sub Test04_MDArrayHasMoreThanTwoDimensions()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Long
        myExpected = 3
        
        Dim myArray(1 To 10, 1 To 10, 1 To 10) As Long
        Dim myResult As Long
        
        'Act:
        myResult = Arrays.Ranks(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
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
        myResult = Arrays.HasRank(myArray, 2)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
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
        myResult = Arrays.HasRank(myArray, 2)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
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
        myResult = Arrays.HasRank(myArray, 2)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
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
        myResult = Arrays.HasRank(myArray, 2)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
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
        myResult = Arrays.HasRank(myArray, 2)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("IsId")
    Private Sub Test10_IsListArrayOnEmptyArraIsFalse()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = False
        
        Dim myArray() As Long
        Dim myResult As Boolean
        
        'Act:
        myResult = Arrays.IsListArray(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("IsId")
    Private Sub Test11_IsListArrayForDefinedArrayIsTrue()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = True
        
        Dim myArray(1 To 10) As Long
        Dim myResult As Boolean
        
        'Act:
        myResult = Arrays.IsListArray(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("Is2d")
    Private Sub Test12_IsListArrayForTableArrayIsFalse()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = False
        
        Dim myArray(1 To 10, 1 To 10) As Long
        Dim myResult As Boolean
        
        'Act:
        myResult = Arrays.IsListArray(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("Is2d")
    Private Sub Test13_IsTableArrayOnEmptyArrayIsFalse()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = False
        
        Dim myArray() As Long
        Dim myResult As Boolean
        
        'Act:
        myResult = Arrays.IsTableArray(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("Is2d")
    Private Sub Test14_IsTableArrayOnListArrayIsFalse()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = False
        
        Dim myArray(1 To 10) As Long
        Dim myResult As Boolean
        
        'Act:
        myResult = Arrays.IsTableArray(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("Is2d")
    Private Sub Test15_IsTableArrayOnTableArrayIsTrue()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = True
        
        Dim myArray(1 To 10, 1 To 10) As Long
        Dim myResult As Boolean
        
        'Act:
        myResult = Arrays.IsTableArray(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("IsMd")
    Private Sub Test16_IsMDArrayOnEmptyArrayIsFalse()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = False
        
        Dim myArray() As Long
        Dim myResult As Boolean
        
        'Act:
        myResult = Arrays.IsMDArray(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("IsMd")
    Private Sub Test17_IsMDArayOnListArrayIsFalse()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = False
        
        Dim myArray(1 To 10) As Long
        Dim myResult As Boolean
        
        'Act:
        myResult = Arrays.IsMDArray(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("IsMd")
    Private Sub Test18_IsMDArrayOnTableArrayIsFalse()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = False
        
        Dim myArray(1 To 10, 1 To 10) As Long
        Dim myResult As Boolean
        
        'Act:
        myResult = Arrays.IsMDArray(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("IsMd")
    Private Sub Test19_IsMDArrayOnMDArrayIsTrue()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = True
        
        Dim myArray(1 To 10, 1 To 10, 1 To 10) As Long
        Dim myResult As Boolean
        
        'Act:
        myResult = Arrays.IsMDArray(myArray)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("TryGetSize")
    Private Sub Test20_CountOnEmptyArrayIsMinusOne()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected As Long
        myExpected = -1
        
        Dim myArray() As Long
        Dim myResult As Long
        
        'Act:
        myResult = Arrays.Count(myArray)
        
        'Assert:
       ' Assert.AreEqual myExpectedStatus, myResultStatus, "Status"
        Assert.AreEqual myExpected, myResult, "Value"
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("TryGetSize")
    Private Sub Test21_TableArrayRankTwoHasCountOfTen()
        On Error GoTo TestFail
        
        'Arrange:
        
        
        Dim myExpectedCount As Long
        myExpectedCount = 10
        
        Dim myArray(1 To 10, 1 To 10, 1 To 10) As Long
        Dim myResultCount As Long
        
       
        Dim myResultValue As Long
        
        'Act:
        myResultCount = Arrays.Count(myArray, 2)
       
        'Assert:
        Assert.AreEqual myExpectedCount, myResultCount  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("TryGetLBound")
    Private Sub Test22_TryExtentOnEmptyArrayHasStatusOfFalse()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus  As Boolean
        myExpectedStatus = False
        
        Dim myExpectedValue As Long
        myExpectedValue = -1
        
        Dim myArray() As Long
        'Dim myResult As Boolean
        Dim myResult As Result
       
        
        Dim myResultState As Boolean
        
        
        'Act:
        myResultState = Arrays.TryExtent(myArray, myResult, 1)
        
        'Assert:
        Assert.AreEqual myExpectedStatus, myResultState, "Status"
        'Assert.AreEqual myExpectedValue, myResult.item(TryLboundResult), "Value"
        
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("TryGetLBound")
    Private Sub Test23_TryExtentOnFilledMDArrayRankTwoHasFiveTenSix()
        On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus  As Boolean
        myExpectedStatus = True
        
        Dim myExpectedFirstIndex As Long
        myExpectedFirstIndex = 5
        
        Dim myExpectedLastIndex As Long
        myExpectedLastIndex = 10
        
        Dim myExpectedCount As Long
        myExpectedCount = 6
        
        Dim myArray(5 To 10, 5 To 10, 5 To 10) As Long
        'Dim myResult As Boolean
        Dim myResult As Result
        
        Dim myResultState As Boolean
        
        
        'Act:
        myResultState = Arrays.TryExtent(myArray, myResult, 2)
        
        'Assert:
        Assert.AreEqual myExpectedStatus, myResultState, "Status"
        Assert.AreEqual myExpectedFirstIndex, myResult.Item(IsFirstIndex), "FirstIndex"
        Assert.areequal myExpectedLastIndex, myResult.Item(IsLastIndex), "LastIndex"
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub
   

   '@TestMethod("TryRotate")
    Private Sub Test26_TransposeArray()
        On Error GoTo TestFail
        
        'Arrange:
        
        Dim myExpectedStatus As Boolean
        myExpectedStatus = True
        
        Dim myexpectedarray As Variant
        myexpectedarray = MakeColRowArray(5, 4)
        
        Dim mySource As Variant
        mySource = MakeRowColArray(4, 5)
        
    
        Dim myResult As Variant
        
        
        'Act:
        myResult = Arrays.Transpose(mySource)
        
        'Assert:
        
        Assert.SequenceEquals myexpectedarray, myResult, "Value"
        
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("TryToLystOfLyst")
    Public Sub Test27_ArrayToLystOfLystsByRow()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus As Boolean
        myExpectedStatus = True
        
        Dim myExpectedLyst As Lyst
        Set myExpectedLyst = Lyst.Deb
        
        ' This is an example of the idiosyncrasy introduced by ParseVariantUsingSingleItemSpecialCase
        ' Which is that if we wish to add a single iterable to a Lyst as a single item
        ' the single iterable must be encapsulated in an array.
        
        With myExpectedLyst
        
            .Add Lyst.Deb.Add(1&, 2&, 3&, 4&)
            .Add Lyst.Deb.Add(5&, 6&, 7&, 8&)
            .Add Lyst.Deb.Add(9&, 10&, 11&, 12&)
            .Add Lyst.Deb.Add(13&, 14&, 15&, 16&)
            .Add Lyst.Deb.Add(17&, 18&, 19&, 20&)
            
        End With
        
        Dim mySource As Variant
        mySource = MakeRowColArray(5, 4)
        
        Dim myResult As Lyst
        
        'Act:
        Set myResult = Arrays.ToLystOfRankLysts(mySource, RankIsRowFirstItemActionIsNoAction)
        
        'Assert:
        
        Dim myIndex As Long
        For myIndex = 0 To 4
           
            Assert.Permissive.SequenceEquals myExpectedLyst.Item(myIndex).ToArray, myResult.Item(myIndex).ToArray  ', CurrentProcedureName
        Next
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("TryToLystOfLyst")
    Public Sub Test28_ArrayToLystOfLystsByCol()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus As Boolean
        myExpectedStatus = True
        
        Dim myExpectedLyst As Lyst
        Set myExpectedLyst = Lyst.Deb
        
        With myExpectedLyst
            
            .Add Lyst.Deb.Add(1&, 5&, 9&, 13&, 17&)
            .Add Lyst.Deb.Add(2&, 6&, 10&, 14&, 18&)
            .Add Lyst.Deb.Add(3&, 7&, 11&, 15&, 19&)
            .Add Lyst.Deb.Add(4&, 8&, 12&, 16&, 20&)
            
        End With
        
        Dim mySource As Variant
        mySource = MakeRowColArray(5, 4)
        
        Dim myResult As Lyst
        
        'Act:
        Set myResult = Arrays.ToLystOfRankLysts(mySource, TableToLystAction.RankIsColumnFirstItemIsNoAction)
        
        'Assert:
       
        Dim myIndex As Long
        For myIndex = 0 To 3
            Assert.SequenceEquals myExpectedLyst.Item(myIndex).ToArray, myResult.Item(myIndex).ToArray  ', CurrentProcedureName
        Next
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("TryToLystOfLyst")
    Public Sub Test29_ArrayToLystOfLystsByRowFirstItemActionIsSplitFirstRow()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus As Boolean
        myExpectedStatus = True
        
        Dim myExpectedLyst As Lyst
        Set myExpectedLyst = Lyst.Deb
        
        With myExpectedLyst
        
            Dim myFirstValues As Lyst
            Set myFirstValues = Lyst.Deb.AddRange(Array(1&, 5&, 9&, 13&, 17&))
            
            .Add myFirstValues
            
            Dim myRankValues As Lyst
            Set myRankValues = Lyst.Deb
        
            With myRankValues
            
                .Add Lyst.Deb.Add(2&, 3&, 4&)
                .Add Lyst.Deb.Add(6&, 7&, 8&)
                .Add Lyst.Deb.Add(10&, 11&, 12&)
                .Add Lyst.Deb.Add(14&, 15&, 16&)
                .Add Lyst.Deb.Add(18&, 19&, 20&)
            
            End With
        
            .Add myRankValues
            
        End With
        
        Dim mySource As Variant
        mySource = MakeRowColArray(5, 4)
        
        Dim myResult As Lyst
        
        'Act:
        Set myResult = Arrays.ToLystOfRankLysts(mySource, RankIsRowFirstItemActionIsSplit)
        
        'Assert:
        Assert.Permissive.SequenceEquals myFirstValues.ToArray, myResult.First.ToArray  ', CurrentProcedureName
       
        Dim myIndex As Long
        For myIndex = 0 To 4
            ' Dim myE As
            ' myE = myRankValues.Item(myIndex).toarray
            Assert.Permissive.SequenceEquals myExpectedLyst.Item(1).Item(myIndex).ToArray, myResult.Item(1).Item(myIndex).ToArray  ', CurrentProcedureName
        Next
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("TryToLystOfLyst")
    Public Sub Test30_ArrayToLystOfLystsByColFirstItemActionIsSplitFirstItem()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus As Boolean
        myExpectedStatus = True
        
        Dim myExpectedLyst As Lyst
        Set myExpectedLyst = Lyst.Deb
        
        With myExpectedLyst
        
            Dim myFirstValues As Lyst
            Set myFirstValues = Lyst.Deb.Add(1&, 2&, 3&, 4&)
            
            .Add myFirstValues
            
            Dim myRankValues As Lyst
            Set myRankValues = Lyst.Deb
        
            With myRankValues
            
                .Add Lyst.Deb.Add(5&, 9&, 13&, 17&)
                .Add Lyst.Deb.Add(6&, 10&, 14&, 18&)
                .Add Lyst.Deb.Add(7&, 11&, 15&, 19&)
                .Add Lyst.Deb.Add(8&, 12&, 16&, 20&)
            
            End With
        
            .Add myRankValues
    
        
        End With
        
        Dim mySource As Variant
        mySource = MakeRowColArray(5, 4)
        
        Dim myResult As Lyst
        
        'Act:
        
        Set myResult = Arrays.ToLystOfRankLysts(mySource, RankIsColumnFirstItemActionIsSplit)
        
        'Assert:
        Dim myexpectedarray As Variant
        myexpectedarray = myFirstValues.ToArray
        
        Dim myResultarray As Variant
        myResultarray = myResult.First.ToArray
        
        Assert.Permissive.SequenceEquals myFirstValues.ToArray, myResult.First.ToArray  ', CurrentProcedureName
        Dim myIndex As Long
        For myIndex = 0 To 3
            Assert.SequenceEquals myRankValues.Item(myIndex).ToArray, myResult.Item(1).Item(myIndex).ToArray  ', CurrentProcedureName
        Next
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("TryToLystOfLyst")
    Public Sub Test31_ArrayToLystOfLystsByRowFirstItemActionIsCopyFirstItem()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus As Boolean
        myExpectedStatus = True
        
        ' Dim myExpectedLyst As Lyst
        ' Set myExpectedLyst = Lyst.Deb
        
        'With myExpectedLyst
        
            Dim myFirstValues As Variant
            myFirstValues = Array(1&, 5&, 9&, 13&, 17&)
            
           ' .Add myFirstValues
            
            Dim myRankValues As Lyst
            Set myRankValues = Lyst.Deb
        
            With myRankValues
            
                .Add Lyst.Deb.Add(1&, 2&, 3&, 4&)
                .Add Lyst.Deb.Add(5&, 6&, 7&, 8&)
                .Add Lyst.Deb.Add(9&, 10&, 11&, 12&)
                .Add Lyst.Deb.Add(13&, 14&, 15&, 16&)
                .Add Lyst.Deb.Add(17&, 18&, 19&, 20&)
            
            End With
        
            '.Add myRankValues
            
       ' End With
        
        Dim mySource As Variant
        mySource = MakeRowColArray(5, 4)
        
        Dim myResult As Lyst
        
        'Act:
        Set myResult = Arrays.ToLystOfRankLysts(mySource, RankIsRowFirstItemActionIsCopy)
        
        'Assert:
        Assert.Permissive.SequenceEquals myFirstValues, myResult.First.ToArray  ', CurrentProcedureName
        Dim myIndex As Long
        For myIndex = 0 To 4
            Assert.SequenceEquals myRankValues.Item(myIndex).ToArray, myResult.Item(1).Item(myIndex).ToArray  ', CurrentProcedureName
        Next
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("TryToLystOfLyst")
    Public Sub Test32_ArrayToLystOfLystsByColFirstActionItemIsSPlitFirstItem()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus As Boolean
        myExpectedStatus = True
        
        ' Dim myExpectedLyst As Lyst
        ' Set myExpectedLyst = Lyst.Deb
        
      '  With myExpectedLyst
        
            Dim myFirstValues As Variant
            myFirstValues = Array(1&, 2&, 3&, 4&)
            
        '    .Add myFirstValues
            
            Dim myRankValues As Lyst
            Set myRankValues = Lyst.Deb
        
            With myRankValues
            
                .Add Lyst.Deb.Add(5&, 9&, 13&, 17&)
                .Add Lyst.Deb.Add(6&, 10&, 14&, 18&)
                .Add Lyst.Deb.Add(7&, 11&, 15&, 19&)
                .Add Lyst.Deb.Add(8&, 12&, 16&, 20&)
            
            
            End With
        
         '   .Add myRankValues
    
        
      '  End With
        
        Dim mySource As Variant
        mySource = MakeRowColArray(5, 4)
        
        Dim myResult As Lyst
        
        'Act:
        Set myResult = Arrays.ToLystOfRankLysts(mySource, RankIsColumnFirstItemActionIsSplit)
        
        'Assert:
        Assert.Permissive.SequenceEquals myFirstValues, myResult.First.ToArray  ', CurrentProcedureName
        Dim myIndex As Long
        For myIndex = 0 To 3
            Assert.SequenceEquals myRankValues.Item(myIndex).ToArray, myResult.Item(1).Item(myIndex).ToArray  ', CurrentProcedureName
        Next
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test  raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
