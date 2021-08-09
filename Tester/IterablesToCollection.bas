Attribute VB_Name = "IterablesToCollection"
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


'@TestMethod("Collection")
Private Sub Test01_FromArray()
    On Error GoTo TestFail
    
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
    Assert.AreEqual myExpected.Item(1), myResult.Item(1)
    Assert.AreEqual myExpected.Item(2), myResult.Item(2)
    Assert.AreEqual myExpected.Item(3), myResult.Item(3)
    Assert.AreEqual myExpected.Item(4), myResult.Item(4)
    
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Collection")
Private Sub Test02_FromArrayList()
    On Error GoTo TestFail
    
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
    Assert.AreEqual myExpected.Item(1), myResult.Item(1)
    Assert.AreEqual myExpected.Item(2), myResult.Item(2)
    Assert.AreEqual myExpected.Item(3), myResult.Item(3)
    Assert.AreEqual myExpected.Item(4), myResult.Item(4)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Collection")
Private Sub Test03_FromCollection()
    On Error GoTo TestFail
    
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
    Assert.AreEqual myExpected.Item(1), myResult.Item(1)
    Assert.AreEqual myExpected.Item(2), myResult.Item(2)
    Assert.AreEqual myExpected.Item(3), myResult.Item(3)
    Assert.AreEqual myExpected.Item(4), myResult.Item(4)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Collection")
Private Sub Test04_FromQueue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Collection
    Set myExpected = New Collection
    myExpected.Add 10
    myExpected.Add 20
    myExpected.Add 30
    myExpected.Add 40
    
    Dim myResult As Collection
    Dim myIterable As Queue
    Set myIterable = New Queue
    myIterable.enqueue 10
    myIterable.enqueue 20
    myIterable.enqueue 30
    myIterable.enqueue 40
    
    'Act:
    Set myResult = Types.Iterable.ToCollection(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected.Item(1), myResult.Item(1)
    Assert.AreEqual myExpected.Item(2), myResult.Item(2)
    Assert.AreEqual myExpected.Item(3), myResult.Item(3)
    Assert.AreEqual myExpected.Item(4), myResult.Item(4)
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Collection")
Private Sub Test05_FromStack()
    On Error GoTo TestFail
    
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
    Set myResult = Types.Iterable.ToCollection(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected.Item(1), myResult.Item(1)
    Assert.AreEqual myExpected.Item(2), myResult.Item(2)
    Assert.AreEqual myExpected.Item(3), myResult.Item(3)
    Assert.AreEqual myExpected.Item(4), myResult.Item(4)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

