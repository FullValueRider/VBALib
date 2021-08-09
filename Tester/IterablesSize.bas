Attribute VB_Name = "IterablesSize"
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


'@TestMethod("Array")
Private Sub Test01_TryGetFromEmptyArray()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    '@Ignore IntegerDataType
    Dim myIterable() As Integer
    Dim myResult As resultlong
    
    'Act:
    Set myResult = Types.Iterable.TryGetSize(myIterable)

    'Assert:
    Assert.AreEqual myExpected, myResult.Status

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Array")
Private Sub Test02_TryGetFromArray()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myExpectedSize As Long
    myExpectedSize = 6
    '@Ignore IntegerDataType
    Dim myIterable(10 To 15) As Integer


    Dim myResultSize As resultlong
    Dim myResult As resultlong

    'Act:
    Set myResult = Types.Iterable.TryGetSize(myIterable, 1, myResultSize)

    'Assert:
    Assert.AreEqual myExpected, myResult.Status, "Status"
    Assert.AreEqual myExpectedSize, myResultSize.Value, "Value"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("ArrayList")
Private Sub Test03_TryGetFromEmptyArrayList()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myIterable As ArrayList
    Set myIterable = New ArrayList

    Dim myResult As resultlong
    Dim myResultSize As resultlong
    'Act:
    Set myResult = Types.Iterable.TryGetSize(myIterable, 1, myResultSize)

    'Assert:
    Assert.AreEqual myExpected, myResult.Status

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ArrayList")
Private Sub Test04_TryGetFromArrayList()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myExpectedSize As Long
    myExpectedSize = 4

    Dim myIterable As ArrayList
    Set myIterable = New ArrayList
    myIterable.Add 10
    myIterable.Add 20
    myIterable.Add 30
    myIterable.Add 40

    Dim myResultSize As resultlong
    Dim myResult As resultlong

    'Act:
    Set myResult = Types.Iterable.TryGetSize(myIterable, 1, myResultSize)

    'Assert:
    Assert.AreEqual myExpected, myResult.Status, "Status"
    Assert.AreEqual myExpectedSize, myResultSize.Value, "Value"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Collection")
Private Sub Test05_TryGetFromEmptyCollection()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myIterable As Collection
    Set myIterable = New Collection

    Dim myResult As resultlong
    Dim myResultSize As resultlong
    'Act:
    Set myResult = Types.Iterable.TryGetSize(myIterable, 1, myResultSize)

    'Assert:
    Assert.AreEqual myExpected, myResult.Status, "Status"
    Assert.AreEqual 0&, myResultSize.Value, "Value"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Collection")
Private Sub Test06_TryGetFromCollection()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myExpectedSize As Long
    myExpectedSize = 4

    Dim myIterable As Collection
    Set myIterable = New Collection
    myIterable.Add 10
    myIterable.Add 20
    myIterable.Add 30
    myIterable.Add 40

    Dim myResultSize As resultlong
    Dim myResult As resultlong

    'Act:
    Set myResult = Types.Iterable.TryGetSize(myIterable, 1, myResultSize)

    'Assert:
    Assert.AreEqual myExpected, myResult.Status, "Status"
    Assert.AreEqual myExpectedSize, myResultSize.Value, "Value"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Queue")
Private Sub Test07_TryGetFromEmptyQueue()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    Dim myIterable As Collection
    Set myIterable = New Collection

    Dim myResult As resultlong
    Dim myResultSize As resultlong
    'Act:
    Set myResult = Types.Iterable.TryGetSize(myIterable, 1, myResultSize)

    'Assert:
    Assert.AreEqual myExpected, myResult.Status, "Status"
    Assert.AreEqual 0&, myResultSize.Value, "Value"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Queue")
Private Sub Test08_TryGetFromQueue()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myExpectedSize As Long
    myExpectedSize = 4

    Dim myIterable As Queue
    Set myIterable = New Queue
    myIterable.enqueue 10
    myIterable.enqueue 20
    myIterable.enqueue 30
    myIterable.enqueue 40

    Dim myResultSize As resultlong
    Dim myResult As resultlong

    'Act:
    Set myResult = Types.Iterable.TryGetSize(myIterable, 1, myResultSize)
    'Assert:
    Assert.AreEqual myExpected, myResult.Status, "Status"
    Assert.AreEqual myExpectedSize, myResultSize.Value, "Value"
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stack")
Private Sub Test09_TryGetFromEmptyStack()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myIterable As Stack
    Set myIterable = New Stack

    Dim myResult As resultlong
    Dim myResultSize As resultlong
    'Act:
    Set myResult = Types.Iterable.TryGetSize(myIterable, 1, myResultSize)

    'Assert:
    Assert.AreEqual myExpected, myResult.Status, "Status"
    Assert.AreEqual 0&, myResultSize.Value, "Value"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stack")
Private Sub Test10_TryGetFromStack()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myExpectedSize As Long
    myExpectedSize = 4

    Dim myIterable As Stack
    Set myIterable = New Stack
    myIterable.Push 10
    myIterable.Push 20
    myIterable.Push 30
    myIterable.Push 40

    Dim myResultSize As resultlong
    Dim myResult As resultlong

    'Act:
    Set myResult = Types.Iterable.TryGetSize(myIterable, 1, myResultSize)

    'Assert:
    Assert.AreEqual myExpected, myResult.Status, "Status"
    Assert.AreEqual myExpectedSize, myResultSize.Value, "Value"

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Array")
Private Sub Test11_GetFromEmptyArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = -1
    
    '@Ignore IntegerDataType
    Dim myIterable() As Integer
    Dim myResult As Long
    
    'Act:
    myResult = Types.Iterable.GetSize(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Array")
Private Sub Test12_GetFromArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 6
    
    '@Ignore IntegerDataType
    Dim myIterable(10 To 15) As Integer
    Dim myResult As Long
    
    'Act:
    myResult = Types.Iterable.GetSize(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("ArrayList")
Private Sub Test13_GetFromEmptyArrayList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 0
    
    Dim myIterable As ArrayList
    Set myIterable = New ArrayList
    
    Dim myResult As Long
    
    'Act:
    myResult = Types.Iterable.GetSize(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ArrayList")
Private Sub Test14_GetFromArrayList()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 4
    
    
    Dim myIterable As ArrayList
    Set myIterable = New ArrayList
    myIterable.Add 10
    myIterable.Add 20
    myIterable.Add 30
    myIterable.Add 40
    
    Dim myResult As Long
    
    'Act:
    myResult = Types.Iterable.GetSize(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Collection")
Private Sub Test15_GetFromEmptyCollection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 0
    
    Dim myIterable As Collection
    Set myIterable = New Collection
    
    Dim myResult As Long

    'Act:
    myResult = Types.Iterable.GetSize(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Collection")
Private Sub Test16_GetFromCollection()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 4
    
    Dim myIterable As Collection
    Set myIterable = New Collection
    myIterable.Add 10
    myIterable.Add 20
    myIterable.Add 30
    myIterable.Add 40
    
    Dim myResult As Long
    
    'Act:
    myResult = Types.Iterable.GetSize(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Queue")
Private Sub Test17_GetFromEmptyQueue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 0
    
    Dim myIterable As Collection
    Set myIterable = New Collection
    
    Dim myResult As Long
    
    'Act:
    myResult = Types.Iterable.GetSize(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Queue")
Private Sub Test18_GetFromQueue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 4
    
    Dim myIterable As Queue
    Set myIterable = New Queue
    myIterable.enqueue 10
    myIterable.enqueue 20
    myIterable.enqueue 30
    myIterable.enqueue 40
    
    Dim myResult As Long
    
    'Act:
    myResult = Types.Iterable.GetSize(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stack")
Private Sub Test19_GetFromEmptyStack()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 0
    
    Dim myIterable As Stack
    Set myIterable = New Stack
    
    Dim myResult As Long

    'Act:
    myResult = Types.Iterable.GetSize(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stack")
Private Sub Test20_GetFromStack()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Long
    myExpected = 4

    Dim myIterable As Stack
    Set myIterable = New Stack
    myIterable.Push 10
    myIterable.Push 20
    myIterable.Push 30
    myIterable.Push 40
    
    '@Ignore VariableNotAssigned
    'Dim myResultSize As resultlong
    Dim myResult As Long
    
    'Act:
    myResult = Types.Iterable.GetSize(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

