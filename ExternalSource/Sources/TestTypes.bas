Attribute VB_Name = "TestTypes"
    '@IgnoreModule
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

   ' Public myPlace As String
    

   Public Sub TypesTests()
        
        
        myPlace = CurrentSourceFile & ":" & CurrentComponentName & "."
        
       Test01_HasItemsEmptyArrayOfIntegerIsFalse
       Test02_HasItemsArrayOfIntegerIsTrue
       Test03_HoldsItemVariantHoldingEmptyIsFalse
       Test04_HoldsItemVariantIsNullArrayIsFalse
       Test05_HoldsItemVariantIsArrayOfIntegerIsTrue
       Test06_HoldsItemArrayListIsNothingIsFalse
       Test07_HoldsItemArrayListIsPopulatedTrue
       Test08_HoldsItemArrayListIsNothingIsFalse
       Test09_HoldsItemArrayListIsPopulatedTrue

    '    Test10_CountEmptyArray
    '    Test11_CountArray
    '    Test12_CountEmptyArrayList
    '    Test13_CountArrayList
    '    Test14_CountEmptyCollection
    '    Test15_CountCollection
    '    Test16_CountEmptyQueue
    '    Test17_CountQueue
    '    Test18_CountEmptyStack
    '    Test19_CountStack
    '    Test20_CountEmptyArray
    '    Test21_CountArray
    '    Test22_CountEmptyArrayList
    '    Test23_CountArrayList
    '    Test24_CountEmptyCollection
    '    Test27_CountCollection
    '    Test28_CountEmptyQueue
    '    Test29_CountQueue
    '    Test30_CountEmptyStack

       Test31_TryExtentWithEmptyArray
       Test32_TryExtentWithFilledArray
       Test33_TryExtentWithEmptyArrayList
       Test34_TryExtentWithFilledArrayList
       Test35_TryExtentWithEmptyCollection
       Test36_TryExtentWithFilledCollection
       Test37_TryExtentWithEmptyQueue
       Test38_TryExtentWithFilledQueue
       Test39_TryExtentWithEmptyStack
       Test40_TryExtentWithFilledStack
       

    '    Test51_LastIndexTryGetFromEmptyArray
    '    Test52_LastIndexTryGetFromArray
    '    Test53_LastIndexTryGetFromEmptyArrayList
    '    Test54_LastIndexTryGetFromArrayList
    '    Test55_LastIndexTryGetFromEmptyCollection
    '    Test56_LastIndexTryGetFromCollection
    '    Test57_LastIndexTryGetFromEmptyQueue
    '    Test58_LastIndexTryGetFromQueue
    '    Test59_LastIndexTryGetFromEmptyStack
    '    Test60_LastIndexTryGetFromStack
    '    Test61_LastIndexGetFromEmptyArray
    '    Test62_LastIndexGetFromArray
    '    Test63_GetFromEmptyArrayList
    '    Test64_LastIndexGetFromArrayList
    '    Test65_LastIndexGetFromEmptyCollection
    '    Test66_LastIndexGetFromCollection
    '    Test67_LastIndexGetFromEmptyQueue
    '    Test68_LastIndexGetFromQueue
    '    Test69_LastIndexGetFromEmptyStack
    '    Test70_LastIndexGetFromStack

       Test71_ToArrayFromArray
       Test72_ToArrayFromArrayList
       Test73_ToArrayFromCollection
       Test74_ToArrayFromQueue
       Test75_ToArrayFromStack

       Test76_ToArrayListFromArray
       Test77_ToArrayListFromArrayList
       Test78_ToArrayListFromCollection
       Test79_ToArrayListFromQueue
       Test80_ToArrayListFromStack

       Test81_ToCollectionFromArray
       Test82_ToCollectionFromArrayList
       Test83_ToCollectionFromCollection
       Test84_ToCollectionFromQueue
       Test85_ToCollectionFromStack

       Test86_ToLystFromArray
       Test87_ToLystFromArrayList
       Test88_ToLystFromLyst
       Test89_ToLystFromQueue
       Test90_ToLystFromStack

       Test91_ToQueueFromArray
       Test92_ToQueueFromArrayList
       Test93_ToQueueFromCollection
       Test94_ToQueueFromQueue
       Test95_ToQueueFromStack

       Test96_ToStackFromArray
       Test97_ToStackFromArrayList
       Test98_ToStackFromCollection
       Test99_ToStackFromQueue
       Test100_ToStackFromStack

       Test101_MinMaxFromArray
       Test102_MinMaxFromArrayList
       Test103_MinMaxFromCollection
       Test104_MinMaxFromQueue
       Test105_MinMaxFromStack

       Test106_SumFromArray
       Test107_SumFromArrayList
       Test108_SumFromCollection
       Test109_SumFromQueue
       Test110_SumFromStack

       Debug.Print CurrentComponentName & vbTab & vbTab & vbTab & "testing completed"

   End Sub
    

#Region "HoldsItem"

    '@TestMethod("HoldsItem") ]
    Private Sub Test01_HasItemsEmptyArrayOfIntegerIsFalse()
        '''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = False
        
        Dim myResult  As Boolean
        Dim myIterable() As Integer
        'Act:
        myResult = Types.Iterable.HasItems(myIterable)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Array")
    Private Sub Test02_HasItemsArrayOfIntegerIsTrue()
       ' ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus  As Boolean
        myExpectedStatus = True
        
        Dim myResultStatus  As Boolean
        Dim myIterable(1 To 5) As Integer
        'Act:
        myResultStatus = Types.Iterable.HasItems(myIterable)
    
        'Assert:
        Assert.AreEqual myExpectedStatus, myResultStatus  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Variant")
    Private Sub Test03_HoldsItemVariantHoldingEmptyIsFalse()
        On Error Resume Next
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = False
        
        Dim myResult  As Boolean
        Dim myIterable As Variant
        myIterable = Empty
        'Act:
        myResult = Types.Iterable.HasItems(myIterable)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        On Error GoTo 0
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Variant")
    Private Sub Test04_HoldsItemVariantIsNullArrayIsFalse()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = False
        
        Dim myResult  As Boolean
        Dim myIterable As Variant
        'Act:
        myIterable = Array()
        myResult = Types.Iterable.HasItems(myIterable)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Variant")
    Private Sub Test05_HoldsItemVariantIsArrayOfIntegerIsTrue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = True
        
        Dim myResult  As Boolean
        Dim myIterable As Variant
        'Act:
        myIterable = Array(1, 2, 3, 4, 5)
        myResult = Types.Iterable.HasItems(myIterable)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("ArrayList")
    Private Sub Test06_HoldsItemArrayListIsNothingIsFalse()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = False
        
        Dim myResult  As Boolean
        Dim myIterable As ArrayList
        Set myIterable = New ArrayList
        'Act:
        
        myResult = Types.Iterable.HasItems(myIterable)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("ArrayList")
    Private Sub Test07_HoldsItemArrayListIsPopulatedTrue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected As Boolean
        myExpected = True
        
        Dim myResult  As Boolean
        Dim myIterable As ArrayList
        Set myIterable = New ArrayList
        With myIterable
        
            .Add 10
            .Add 20
            .Add 30
            .Add 40
            .Add 50
            
        End With
        'Act:
        
        myResult = Types.Iterable.HasItems(myIterable)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Collection")
    Private Sub Test08_HoldsItemArrayListIsNothingIsFalse()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = False
        
        Dim myResult  As Boolean
        Dim myColl As Collection
        Set myColl = New Collection
        'Act:
        
        myResult = Types.Iterable.HasItems(myColl)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Collection")
    Private Sub Test09_HoldsItemArrayListIsPopulatedTrue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Boolean
        myExpected = True
        
        Dim myResult  As Boolean
        Dim myColl As Collection
        Set myColl = New Collection
        myColl.Add 10
        myColl.Add 20
        myColl.Add 30
        myColl.Add 40
        myColl.Add 50
        'Act:
        
        myResult = Types.Iterable.HasItems(myColl)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
    
#End Region


#Region "TryExtent"

   '@TestMethod("Array")
    Private Sub Test31_TryExtentWithEmptyArray()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus  As Boolean
        myExpectedStatus = False
        
        '@Ignore IntegerDataType
        Dim myIterable() As Integer
        
        Dim myResultStatus As Boolean
        Dim myResult As Result
        
        'Act:
        '@Ignore ImplicitDefaultMemberAccess
        myResultStatus = Types.Iterable.TryExtent(myIterable, myResult)
    
        'Assert:
        Assert.AreEqual myExpectedStatus, myResultStatus ' , myPlace & CurrentProcedureName & ":Status"
        

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("Array")
    Private Sub Test32_TryExtentWithFilledArray()
        ''On Error GoTo TestFail
       ' If Assert Is Nothing Then Set Assert = New AssertClass
        'Arrange:
        Dim myExpectedStatus  As Boolean
        myExpectedStatus = True
        
        Dim myExpectedFirstIndex As Long
        myExpectedFirstIndex = 10&
        
        Dim myExpectedLastIndex As Long
        myExpectedLastIndex = 15&
        
        Dim myExpectedRun As Long
        myExpectedRun = 6&
        
        
        '@Ignore IntegerDataType
        Dim myIterable(10 To 15) As Integer
        
        
        Dim myResult As Result
        Dim myResultStatus As Boolean
        
        'Act:
        myResultStatus = Types.Iterable.TryExtent(myIterable, myResult)
    
        'Assert:
        Assert.AreEqual myExpectedStatus, myResultStatus, ":Status" ' , myPlace & CurrentProcedureName & ":Status"
        Assert.AreEqual myExpectedFirstIndex, myResult.Item(IsFirstIndex), ":FirstIndex" ' , myPlace & CurrentProcedureName & ":FirstIndex"
        Assert.AreEqual myExpectedLastIndex, myResult.Item(IsLastIndex), ":LastIndex" ' , myPlace & CurrentProcedureName & ":LastIndex "
        Assert.AreEqual myExpectedRun, myResult.Item(IsRun), ":Value" ' , myPlace & CurrentProcedureName & ":Value"
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("ArrayList")
    Private Sub Test33_TryExtentWithEmptyArrayList()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus  As Boolean
        myExpectedStatus = False
        
        Dim myExpected As Long
        myExpected = -1
        
        Dim myIterable As ArrayList
        Set myIterable = New ArrayList
        
        Dim myResult As Result
        Dim myResultStatus As Boolean
        'Act:
        myResultStatus = Types.Iterable.TryExtent(myIterable, myResult)
    
        'Assert:
        Assert.AreEqual myExpectedStatus, myResultStatus ' , myPlace & CurrentProcedureName & ":Status"
       

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("ArrayList")
    Private Sub Test34_TryExtentWithFilledArrayList()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Long
        myExpected = 0
        
        Dim myExpectedStatus As Boolean
        myExpectedStatus = True
        
        Dim myIterable As ArrayList
        Set myIterable = New ArrayList
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        Dim myResult As Result
        Dim myResultStatus As Boolean
        
        'Act:
        
        myResultStatus = Types.Iterable.TryExtent(myIterable, myResult)
    
        'Assert:
        Assert.AreEqual myExpectedStatus, myResultStatus ' , myPlace & CurrentProcedureName & ":Status"
        'Assert.AreEqual myExpected, myResult.Item(0) ' , myPlace & CurrentProcedureName & ":Value"
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Collection")
    Private Sub Test35_TryExtentWithEmptyCollection()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus  As Boolean
        myExpectedStatus = False
        
        Dim myExpected  As Long
        myExpected = -1
        
        Dim myIterable As Collection
        Set myIterable = New Collection
        
        Dim myResultStatus As Boolean
        Dim myResult As Result
        'Act:
        
        myResultStatus = Types.Iterable.TryExtent(myIterable, myResult)
    
        'Assert:
        Assert.AreEqual myExpectedStatus, myResultStatus ' , myPlace & CurrentProcedureName & ":Status"
        'Assert.AreEqual myExpected, myResult.Item(0) ' , myPlace & CurrentProcedureName & ":Value"
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("Collection")
    Private Sub Test36_TryExtentWithFilledCollection()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus  As Boolean
        myExpectedStatus = True
        
        Dim myExpectedFirstIndex As Long
        myExpectedFirstIndex = 1
        
        Dim myExpectedLastIndex As Long
        myExpectedLastIndex = 4
        
        Dim myExpectedRun As Long
        myExpectedRun = 4
        
        Dim myIterable As Collection
        Set myIterable = New Collection
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        Dim myResult As Result
        Dim myResultStatus As Boolean
        
        'Act:
        '@Ignore ImplicitDefaultMemberAccess
        myResultStatus = Types.Iterable.TryExtent(myIterable, myResult)
    
        'Assert:
        Assert.AreEqual myExpectedStatus, myResultStatus ' , myPlace & CurrentProcedureName & ":Status"
        Assert.AreEqual myExpectedFirstIndex, myResult.Item(IsFirstIndex), ":FirstIndex" ' , myPlace & CurrentProcedureName & ":FirstIndex"
        Assert.AreEqual myExpectedLastIndex, myResult.Item(IsLastIndex), ":LastIndex" ' , myPlace & CurrentProcedureName & ":LastIndex "
        Assert.AreEqual myExpectedRun, myResult.Item(IsRun), ":Value" ' , myPlace & CurrentProcedureName & ":run"
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Queue")
    Private Sub Test37_TryExtentWithEmptyQueue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus  As Boolean
        myExpectedStatus = False
        
        Dim myExpected As Long
        myExpected = -1
        
        Dim myIterable As Collection
        Set myIterable = New Collection
        
        Dim myResultStatus As Boolean
        Dim myResult As Result
        'Act:
        myResultStatus = Types.Iterable.TryExtent(myIterable, myResult)
    
        'Assert:
        Assert.AreEqual myExpectedStatus, myResultStatus ' , myPlace & CurrentProcedureName & ":Status"
        'Assert.AreEqual myExpected, myResult.Item(TryFirstIndexResult) ' , myPlace & CurrentProcedureName & ":Value"

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Queue")
    Private Sub Test38_TryExtentWithFilledQueue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus  As Boolean
        myExpectedStatus = False
        
        Dim myExpectedCount As Long
        myExpectedCount = 4
        
        Dim myIterable As Queue
        Set myIterable = New Queue
        myIterable.EnQueue 10
        myIterable.EnQueue 20
        myIterable.EnQueue 30
        myIterable.EnQueue 40
        
        Dim myResult As Result
        Dim myResultStatus As Boolean
        
        'Act:
        myResultStatus = Types.Iterable.TryExtent(myIterable, myResult)
    
        'Assert:
        Assert.AreEqual myExpectedStatus, myResult.Status  ' ' , myPlace & CurrentProcedureName & ":Status"
        Assert.AreEqual myExpectedCount, myResult.Item(Extent.IsCount)  ' ' , myPlace & CurrentProcedureName & ":Value"
        Assert.areequal Empty, myResult.Item(Extent.IsFirstIndex)
        Assert.areequal Empty, myResult.Item(Extent.IsLastIndex)
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Stack")
    Private Sub Test39_TryExtentWithEmptyStack()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus  As Boolean
        myExpectedStatus = False
        
        Dim myExpected As Long
        myExpected = -1
        
        Dim myIterable As Stack
        Set myIterable = New Stack
        
        Dim myResultStatus As Boolean
        Dim myResult As Result
        'Act:
        '@Ignore ImplicitDefaultMemberAccess
        myResultStatus = Types.Iterable.TryExtent(myIterable, myResult)
    
        'Assert:
        Assert.AreEqual myExpectedStatus, myResultStatus  ' ' , myPlace & CurrentProcedureName & ":Status"
        'Assert.AreEqual -1&, myResult.Item(Types.Iterable.GetConst(TryFirstIndexResult))  ' ' , myPlace & CurrentProcedureName & ":Value"
TestExit:

        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Stack")
    Private Sub Test40_TryExtentWithFilledStack()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpectedStatus  As Boolean
        myExpectedStatus = False
        
        Dim myExpectedCount As Long
        myExpectedCount = 4
        
        Dim myIterable As Stack
        Set myIterable = New Stack
        myIterable.Push 10
        myIterable.Push 20
        myIterable.Push 30
        myIterable.Push 40
        
        Dim myResult As Result
        Dim myResultStatus As Boolean
        
        'Act:
        '@Ignore ImplicitDefaultMemberAccess
        myResultStatus = Types.Iterable.TryExtent(myIterable, myResult)
    
        'Assert:
        Assert.AreEqual myExpectedStatus, myResultStatus  ' ' , myPlace & CurrentProcedureName & ":Status"
        Assert.AreEqual myExpectedCount, myResult.Item(Extent.IsCount)  ' ' , myPlace & CurrentProcedureName & ":Value"
        Assert.areequal Empty, myResult.Item(Extent.IsFirstIndex)
        Assert.areequal Empty, myResult.Item(Extent.IsLastIndex)
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
    
#End Region


#Region "ToArray"
   '@TestMethod("Array")
    Private Sub Test71_ToArrayFromArray()
        '''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Variant
        myExpected = Array(10, 20, 30, 40)
        
       
        
        Dim myResult As Variant
        
        '@Ignore IntegerDataType
        Dim myIterable(0 To 3) As Integer
        myIterable(0) = 10
        myIterable(1) = 20
        myIterable(2) = 30
        myIterable(3) = 40
        
        'Act:
        myResult = Types.Iterable.ToArray(myIterable)
    
        'Assert:
        
        Assert.Permissive.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & ":Value"

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Array")
    Private Sub Test72_ToArrayFromArrayList()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Variant
        myExpected = Array(10, 20, 30, 40)
        
        
        
        Dim myResult As Variant
       
        Dim myIterable As ArrayList
        Set myIterable = New ArrayList
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        'Act:
        myResult = Types.Iterable.ToArray(myIterable)
    
        'Assert:
        
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & ":Value"

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("Array")
    Private Sub Test73_ToArrayFromCollection()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Variant
        myExpected = Array(10, 20, 30, 40)
        
        Dim myResult As Variant
        
        Dim myIterable As Collection
        Set myIterable = New Collection
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        'Act:
        myResult = Types.Iterable.ToArray(myIterable)
    
        'Assert:
        
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & ":Value"

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("Array")
    Private Sub Test74_ToArrayFromQueue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Variant
        myExpected = Array(10, 20, 30, 40)
        
        
        
        Dim myResult As Variant
        
        Dim myIterable As Queue
        Set myIterable = New Queue
        myIterable.EnQueue 10
        myIterable.EnQueue 20
        myIterable.EnQueue 30
        myIterable.EnQueue 40
        
        'Act:
        myResult = Types.Iterable.ToArray(myIterable)
    
        'Assert:
        
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & ":Value"

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Array")
    Private Sub Test75_ToArrayFromStack()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Variant
        myExpected = Array(10, 20, 30, 40)
        
        Dim myResult As Variant
        
        
        Dim myIterable As Stack
        ' To allow sequence equals we need to push onto the stack in reverse order
        ' because the last item pushed is item 1 which is significant
        ' when using the for each loop to transfer values
        Set myIterable = New Stack
        myIterable.Push 40 ' Item 4
        myIterable.Push 30 ' Item 3
        myIterable.Push 20 ' Item 2
        myIterable.Push 10 ' Item 1
        
        'Act:
        myResult = Types.Iterable.ToArray(myIterable)
    
        'Assert:
       
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & ":Value"

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
#End Region
    
#Region "ToArrayList"

   '@TestMethod("ArrayList")
    Private Sub Test76_ToArrayListFromArray()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As ArrayList
        Set myExpected = New ArrayList
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        
        
        Dim myResult As ArrayList
        
        '@Ignore IntegerDataType
        Dim myIterable(10 To 13) As Integer
        myIterable(10) = 10
        myIterable(11) = 20
        myIterable(12) = 30
        myIterable(13) = 40
        
        'Act:
         Set myResult = Types.Iterable.ToArrayList(myIterable)
    
        'Assert:
        
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray ' , myPlace & CurrentProcedureName & ":Value"


TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("ArrayList")
    Private Sub Test77_ToArrayListFromArrayList()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As ArrayList
        Set myExpected = New ArrayList
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        
        Dim myResult As ArrayList
       
        
        Dim myIterable As ArrayList
        Set myIterable = New ArrayList
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        'Act:
        Set myResult = Types.Iterable.ToArrayList(myIterable)
    
        'Assert:
         Assert.SequenceEquals myExpected.ToArray, myResult.ToArray ' , myPlace & CurrentProcedureName & ":Value"

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("ArrayList")
    Private Sub Test78_ToArrayListFromCollection()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As ArrayList
        Set myExpected = New ArrayList
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
       
        
        Dim myResult As ArrayList
       
        
        Dim myIterable As Collection
        Set myIterable = New Collection
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        'Act:
         Set myResult = Types.Iterable.ToArrayList(myIterable)
    
        'Assert:
        
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray ' , myPlace & CurrentProcedureName & ":Value"

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("ArrayList")
    Private Sub Test79_ToArrayListFromQueue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As ArrayList
        Set myExpected = New ArrayList
        
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        
        
        Dim myResult As ArrayList
        
        
        Dim myIterable As Queue
        Set myIterable = New Queue
        myIterable.EnQueue 10
        myIterable.EnQueue 20
        myIterable.EnQueue 30
        myIterable.EnQueue 40
        
        'Act:
         Set myResult = Types.Iterable.ToArrayList(myIterable)
    
        'Assert:
        
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray ' , myPlace & CurrentProcedureName & ":Value"

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("ArrayList")
    Private Sub Test80_ToArrayListFromStack()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As ArrayList
        Set myExpected = New ArrayList
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        Dim myResult As ArrayList
        
        Dim myIterable As Stack
        ' To allow sequence equals we need to push onto the stack in reverse order
        ' because the last item pushed is item 1 which is significant
        ' when using the for each loop to transfer values
        Set myIterable = New Stack
        myIterable.Push 40
        myIterable.Push 30
        myIterable.Push 20
        myIterable.Push 10
        
        'Act:
        Set myResult = Types.Iterable.ToArrayList(myIterable)
    
        'Assert:
        
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray ' , myPlace & CurrentProcedureName & ":Value"

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
    
    
#End Region

#Region "ToCollection"

   '@TestMethod("Collection")
    Private Sub Test81_ToCollectionFromArray()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Collection
        Set myExpected = New Collection
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        
        '@Ignore IntegerDataType
        Dim myIterable(10 To 13) As Integer
        myIterable(10) = 10
        myIterable(11) = 20
        myIterable(12) = 30
        myIterable(13) = 40
        
        Dim myResult As Collection
        
        'Act:
        Set myResult = Types.Iterable.ToCollection(myIterable)
    
        'Assert:
        
         Assert.AreEqual myExpected.Item(1), myResult.Item(1)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(2), myResult.Item(2)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(3), myResult.Item(3)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(4), myResult.Item(4)  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Collection")
    Private Sub Test82_ToCollectionFromArrayList()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Collection
        Set myExpected = New Collection
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        
        
        Dim myResult As Collection
        
        Dim myIterable As ArrayList
        Set myIterable = New ArrayList
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        'Act:
        Set myResult = Types.Iterable.ToCollection(myIterable)
    
        'Assert:
        
         Assert.AreEqual myExpected.Item(1), myResult.Item(1)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(2), myResult.Item(2)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(3), myResult.Item(3)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(4), myResult.Item(4)  ', CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("Collection")
    Private Sub Test83_ToCollectionFromCollection()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Collection
        Set myExpected = New Collection
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        Dim myResult As Collection
       
        Dim myIterable As Collection
        Set myIterable = New Collection
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        'Act:
         Set myResult = Types.Iterable.ToCollection(myIterable)
    
        'Assert:
        
         Assert.AreEqual myExpected.Item(1), myResult.Item(1)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(2), myResult.Item(2)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(3), myResult.Item(3)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(4), myResult.Item(4)  ', CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("Collection")
    Private Sub Test84_ToCollectionFromQueue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Collection
        Set myExpected = New Collection
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        
        
        Dim myResult As Collection
        Dim myResultStatus  As Boolean
        
        Dim myIterable As Queue
        Set myIterable = New Queue
        myIterable.EnQueue 10
        myIterable.EnQueue 20
        myIterable.EnQueue 30
        myIterable.EnQueue 40
        
        'Act:
        Set myResult = Types.Iterable.ToCollection(myIterable)
    
        'Assert:
        
         Assert.AreEqual myExpected.Item(1), myResult.Item(1)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(2), myResult.Item(2)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(3), myResult.Item(3)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(4), myResult.Item(4)  ', CurrentProcedureName
         
TestExit:
        Exit Sub
        
TestFail:
    
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("Collection")
    Private Sub Test85_ToCollectionFromStack()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Collection
        Set myExpected = New Collection
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        
        
        Dim myResult As Collection
        
        
        Dim myIterable As Stack
        ' To allow sequence equals we need to push onto the stack in reverse order
        ' because the last item pushed is item 1 which is significant
        ' when using the for each loop to transfer values
         Set myIterable = New Stack
         myIterable.Push 40
         myIterable.Push 30
         myIterable.Push 20
         myIterable.Push 10
        
        'Act:
         Set myResult = Types.Iterable.ToCollection(10, 20, 30, 40)
    
        'Assert:
        
        'cOMPARE ITEM BY TEM AS COLLECTIONS DON'T HAVE A TO ARRAY METHOD
         Assert.AreEqual myExpected.Item(1), myResult.Item(1)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(2), myResult.Item(2)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(3), myResult.Item(3)  ', CurrentProcedureName
         Assert.AreEqual myExpected.Item(4), myResult.Item(4)  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
#End Region

#Region "ToLyst"

   '@TestMethod("Lyst")
    Private Sub Test86_ToLystFromArray()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Lyst
        Set myExpected = Lyst.Deb
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
       Dim myArray(0 To 3) As Long
       myArray(0) = 10
       myArray(1) = 20
       myArray(2) = 30
       myArray(3) = 40
       
        Dim myResult As Lyst
        
        
        'Act:
        Set myResult = Types.Iterable.ToLyst(myArray)
    
        'Assert:
        Assert.Permissive.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName
        ' Assert.AreEqual myExpected.Item(0), myResult.Item(0)  ', CurrentProcedureName
        '  Assert.AreEqual myExpected.Item(1), myResult.Item(1)  ', CurrentProcedureName
        '  Assert.AreEqual myExpected.Item(2), myResult.Item(2)  ', CurrentProcedureName
        '  Assert.AreEqual myExpected.Item(3), myResult.Item(3)  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("Lyst")
    Private Sub Test87_ToLystFromArrayList()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Lyst
        Set myExpected = Lyst.Deb
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        
        
        Dim myResult As Lyst
        
    
        
        'Act:
         Set myResult = Types.Iterable.ToLyst(10, 20, 30, 40)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("Lyst")
    Private Sub Test88_ToLystFromLyst()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Lyst
        Set myExpected = Lyst.Deb
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        Dim myResult As Lyst
        Dim myIterable As Lyst
        Set myIterable = Lyst.Deb
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        'Act:
        Set myResult = Types.Iterable.ToLyst(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArrayList.ToArray, myResult.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub


   '@TestMethod("Lyst")
    Private Sub Test89_ToLystFromQueue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Lyst
        Set myExpected = Lyst.Deb
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        Dim myResult As Lyst
        Dim myIterable As Queue
        Set myIterable = New Queue
        myIterable.EnQueue 10
        myIterable.EnQueue 20
        myIterable.EnQueue 30
        myIterable.EnQueue 40
        
        'Act:
        Set myResult = Types.Iterable.ToLyst(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

   '@TestMethod("Lyst")
    Private Sub Test90_ToLystFromStack()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Lyst
        Set myExpected = Lyst.Deb
        myExpected.Add 10
        myExpected.Add 20
        myExpected.Add 30
        myExpected.Add 40
        
        Dim myResult As Lyst
        Dim myIterable As Stack
        Set myIterable = New Stack
        ' To allow sequence equals we need to push onto the stack in reverse order
        ' because the last item pushed is item 1 which is significant
        ' when using the for each loop to transfer values
        ' Set myIterable = New Stack
        myIterable.Push 40
        myIterable.Push 30
        myIterable.Push 20
        myIterable.Push 10
        
        'Act:
        Set myResult = Types.Iterable.ToLyst(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArrayList.ToArray, myResult.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
        
    End Sub

#End Region

#Region "ToQueue"

   '@TestMethod("Queue")
    Private Sub Test91_ToQueueFromArray()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Queue
        Set myExpected = New Queue
        myExpected.EnQueue 10
        myExpected.EnQueue 20
        myExpected.EnQueue 30
        myExpected.EnQueue 40
        
        
        Dim myIterable(10 To 13) As Integer
        myIterable(10) = 10
        myIterable(11) = 20
        myIterable(12) = 30
        myIterable(13) = 40
        
        Dim myResult As Queue
        
        'Act:
        Set myResult = Types.Iterable.ToQueue(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Queue")
    Private Sub Test92_ToQueueFromArrayList()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Queue
        Set myExpected = New Queue
        myExpected.EnQueue 10
        myExpected.EnQueue 20
        myExpected.EnQueue 30
        myExpected.EnQueue 40
        
        Dim myIterable As ArrayList
        Set myIterable = New ArrayList
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        Dim myResult As Queue
        
        'Act:
        Set myResult = Types.Iterable.ToQueue(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("Queue")
    Private Sub Test93_ToQueueFromCollection()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Queue
        Set myExpected = New Queue
        myExpected.EnQueue 10
        myExpected.EnQueue 20
        myExpected.EnQueue 30
        myExpected.EnQueue 40
        
        Dim myIterable As Collection
        Set myIterable = New Collection
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        Dim myResult As Queue
        
        'Act:
        Set myResult = Types.Iterable.ToQueue(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("Queue")
    Private Sub Test94_ToQueueFromQueue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Queue
        Set myExpected = New Queue
        myExpected.EnQueue 10
        myExpected.EnQueue 20
        myExpected.EnQueue 30
        myExpected.EnQueue 40
        
        Dim myResult As Queue
        Dim myIterable As Queue
        Set myIterable = New Queue
        myIterable.EnQueue 10
        myIterable.EnQueue 20
        myIterable.EnQueue 30
        myIterable.EnQueue 40
        
        'Act:
        Set myResult = Types.Iterable.ToQueue(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Queue")
    Private Sub Test95_ToQueueFromStack()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Queue
        Set myExpected = New Queue
        myExpected.EnQueue 10
        myExpected.EnQueue 20
        myExpected.EnQueue 30
        myExpected.EnQueue 40
        
        Dim myResult As Queue
        Dim myIterable As Stack
        ' To allow sequence equals we need to push onto the stack in reverse order
        ' because the last item pushed is item 1 which is significant
        ' when using the for each loop to transfer values
        Set myIterable = New Stack
        myIterable.Push 40
        myIterable.Push 30
        myIterable.Push 20
        myIterable.Push 10
        
        'Act:
        Set myResult = Types.Iterable.ToQueue(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

#End Region

#Region "ToStack"

   '@TestMethod("Stack")
    Private Sub Test96_ToStackFromArray()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Stack
        Set myExpected = New Stack
        myExpected.Push 10
        myExpected.Push 20
        myExpected.Push 30
        myExpected.Push 40
        
        
        Dim myIterable(10 To 13) As Integer
        myIterable(10) = 10
        myIterable(11) = 20
        myIterable(12) = 30
        myIterable(13) = 40
        
        Dim myResult As Stack
        
        'Act:
        Set myResult = Types.Iterable.ToStack(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Stack")
    Private Sub Test97_ToStackFromArrayList()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Stack
        Set myExpected = New Stack
        myExpected.Push 10
        myExpected.Push 20
        myExpected.Push 30
        myExpected.Push 40
        
        Dim myIterable As ArrayList
        Set myIterable = New ArrayList
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        Dim myResult As Stack
        
        'Act:
        Set myResult = Types.Iterable.ToStack(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("Stack")
    Private Sub Test98_ToStackFromCollection()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Stack
        Set myExpected = New Stack
        myExpected.Push 10
        myExpected.Push 20
        myExpected.Push 30
        myExpected.Push 40
        
        Dim myIterable As Collection
        Set myIterable = New Collection
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        Dim myResult As Stack
        
        'Act:
        Set myResult = Types.Iterable.ToStack(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("Stack")
    Private Sub Test99_ToStackFromQueue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Stack
        Set myExpected = New Stack
        myExpected.Push 10
        myExpected.Push 20
        myExpected.Push 30
        myExpected.Push 40
        
        
        Dim myIterable As Queue
        Set myIterable = New Queue
        myIterable.EnQueue 10
        myIterable.EnQueue 20
        myIterable.EnQueue 30
        myIterable.EnQueue 40
        
        Dim myResult As Stack
        
        'Act:
        Set myResult = Types.Iterable.ToStack(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("Stack")
    Private Sub Test100_ToStackFromStack()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Stack
        Set myExpected = New Stack
        myExpected.Push 40
        myExpected.Push 30
        myExpected.Push 20
        myExpected.Push 10
        
        ' To allow sequence equals we need to push onto the stack in reverse order
        ' because the last item pushed is item 1 which is significant
        ' when using the for each loop to transfer values
        Dim myIterable As Stack
        Set myIterable = New Stack
        myIterable.Push 40
        myIterable.Push 30
        myIterable.Push 20
        myIterable.Push 10
        
        
        Dim myResult As Stack
        'Act:
        Set myResult = Types.Iterable.ToStack(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected.ToArray, myResult.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
    
#End Region

#Region "MinMax"

   '@TestMethod("MinMax")
    Private Sub Test101_MinMaxFromArray()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Variant
        myExpected = Array(10, 40)
        
        '@Ignore IntegerDataType
        Dim myIterable(10 To 13) As Integer
        myIterable(10) = 20
        myIterable(11) = 10
        myIterable(12) = 40
        myIterable(13) = 30
        
        Dim myResult As Variant
        
        'Act:
        myResult = Types.Iterable.MinMax(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("MinMax")
    Private Sub Test102_MinMaxFromArrayList()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Variant
        myExpected = Array(10, 40)
        
        Dim myIterable As ArrayList
        Set myIterable = New ArrayList
        myIterable.Add 20
        myIterable.Add 10
        myIterable.Add 40
        myIterable.Add 30
        
        Dim myResult As Variant
        
        'Act:
        myResult = Types.Iterable.MinMax(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("MinMax")
    Private Sub Test103_MinMaxFromCollection()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Variant
        myExpected = Array(10, 40)
            
        
        Dim myIterable As Collection
        Set myIterable = New Collection
        myIterable.Add 20
        myIterable.Add 10
        myIterable.Add 40
        myIterable.Add 30
        
        Dim myResult As Variant
            
        'Act:
        myResult = Types.Iterable.MinMax(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("MinMax")
    Private Sub Test104_MinMaxFromQueue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Variant
        myExpected = Array(10, 40)
            
        Dim myIterable As Queue
        Set myIterable = New Queue
        myIterable.EnQueue 20
        myIterable.EnQueue 10
        myIterable.EnQueue 40
        myIterable.EnQueue 30
        
        Dim myResult As Variant
        
        'Act:
         myResult = Types.Iterable.MinMax(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("MinMax")
    Private Sub Test105_MinMaxFromStack()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Variant
        myExpected = Array(10, 40)
            
        Dim myIterable As Stack
        ' To allow sequence equals we need to push onto the stack in reverse order
        ' because the last item pushed is item 1 which is significant
        ' when using the for each loop to transfer values
        Set myIterable = New Stack
        myIterable.Push 30
        myIterable.Push 40
        myIterable.Push 10
        myIterable.Push 20
        
        Dim myResult As Variant
        
        'Act:
        myResult = Types.Iterable.MinMax(myIterable)
    
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
#End Region

#Region "Sum"
   '@TestMethod("ArrayList")
    Private Sub Test106_SumFromArray()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Long
        myExpected = 100
        
        Dim myIterable(10 To 13) As Integer
        myIterable(10) = 10
        myIterable(11) = 20
        myIterable(12) = 30
        myIterable(13) = 40
        
        Dim myResult As Long
        
        'Act:
        myResult = Types.Iterable.Sum(myIterable)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("ArrayList")
    Private Sub Test107_SumFromArrayList()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Long
        myExpected = 100
        
        Dim myIterable As ArrayList
        Set myIterable = New ArrayList
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        Dim myResult As Long
        
        'Act:
        myResult = Types.Iterable.Sum(myIterable)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("ArrayList")
    Private Sub Test108_SumFromCollection()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Long
        myExpected = 100
            
        
        Dim myIterable As Collection
        Set myIterable = New Collection
        myIterable.Add 10
        myIterable.Add 20
        myIterable.Add 30
        myIterable.Add 40
        
        Dim myResult As Long
            
        'Act:
        myResult = Types.Iterable.Sum(myIterable)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub


   '@TestMethod("ArrayList")
    Private Sub Test109_SumFromQueue()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Long
        myExpected = 100
            
        Dim myIterable As Queue
        Set myIterable = New Queue
        myIterable.EnQueue 10
        myIterable.EnQueue 20
        myIterable.EnQueue 30
        myIterable.EnQueue 40
        
        Dim myResult As Long
        
        'Act:
        myResult = Types.Iterable.Sum(myIterable)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

   '@TestMethod("ArrayList")
    Private Sub Test110_SumFromStack()
        ''On Error GoTo TestFail
        
        'Arrange:
        Dim myExpected  As Long
        myExpected = 100
            
        Dim myIterable As Stack
        ' To allow sequence equals we need to push onto the stack in reverse order
        ' because the last item pushed is item 1 which is significant
        ' when using the for each loop to transfer values
        Set myIterable = New Stack
        myIterable.Push 40
        myIterable.Push 30
        myIterable.Push 20
        myIterable.Push 10
        
        Dim myResult As Long
        
        'Act:
        myResult = Types.Iterable.Sum(myIterable)
    
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
#End Region
