Attribute VB_Name = "IterablesHoldsItems"
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
Private Sub Test01_EmptyArrayOfIntegerIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    Dim myIterable() As Integer
    'Act:
    myResult = Types.Iterable.HasItems(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Array")
Private Sub Test02_ArrayOfIntegerOneToFiveIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    Dim myIterable(1 To 5) As Integer
    'Act:
    myResult = Types.Iterable.HasItems(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Variant")
Private Sub Test03_VariantHoldingEmptyIsFalse()
    On Error Resume Next
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    Dim myIterable As Variant
    myIterable = Empty
    'Act:
    myResult = Types.Iterable.HasItems(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    On Error GoTo 0
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Variant")
Private Sub Test04_VariantIsNullArrayIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    Dim myIterable As Variant
    'Act:
    myIterable = Array()
    myResult = Types.Iterable.HasItems(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Variant")
Private Sub Test05_VariantIsArrayOfIntegerIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    Dim myIterable As Variant
    'Act:
    myIterable = Array(1, 2, 3, 4, 5)
    myResult = Types.Iterable.HasItems(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ArrayList")
Private Sub Test06_ArrayListIsNothingIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    Dim myIterable As ArrayList
    Set myIterable = New ArrayList
    'Act:
    
    myResult = Types.Iterable.HasItems(myIterable)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("ArrayList")
Private Sub Test07_ArrayListIsPopulatedTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
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
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Collection")
Private Sub Test08_ArrayListIsNothingIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    Dim myColl As Collection
    Set myColl = New Collection
    'Act:
    
    myResult = Types.Iterable.HasItems(myColl)
   
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Collection")
Private Sub Test09_ArrayListIsPopulatedTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
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
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
