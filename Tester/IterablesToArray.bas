Attribute VB_Name = "IterablesToArray"
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
Private Sub Test01_FromArray()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 40)
    
    Dim myResult As Variant
    '@Ignore IntegerDataType
    Dim myIterable(10 To 13) As Integer
    myIterable(10) = 10
    myIterable(11) = 20
    myIterable(12) = 30
    myIterable(13) = 40
    
    'Act:
    myResult = Types.Iterable.toarray(myIterable)
   
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Array")
Private Sub Test02_FromArrayList()
    On Error GoTo TestFail
    
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
    myResult = Types.Iterable.toarray(myIterable)
   
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Array")
Private Sub Test03_FromCollection()
    On Error GoTo TestFail
    
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
    myResult = Types.Iterable.toarray(myIterable)
   
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Array")
Private Sub Test04_FromQueue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Variant
    myExpected = Array(10, 20, 30, 40)
    
    Dim myResult As Variant
    Dim myIterable As Queue
    Set myIterable = New Queue
    myIterable.enqueue 10
    myIterable.enqueue 20
    myIterable.enqueue 30
    myIterable.enqueue 40
    
    'Act:
    myResult = Types.Iterable.toarray(myIterable)
   
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Array")
Private Sub Test05_FromStack()
    On Error GoTo TestFail
    
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
    myResult = Types.Iterable.toarray(myIterable)
   
    'Assert:
    Assert.SequenceEquals myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

