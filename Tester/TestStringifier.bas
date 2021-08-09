Attribute VB_Name = "TestStringifier"
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

'@TestMethod("Primitive")
Private Sub Test01_Primitive1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "1"
    
    
    Dim myResult As String
    
    'Act:
    myResult = stringifier.pvStringifyPrimitive(1)
   
    'Assert:
    Assert.AreEqual myExpected, myResult
    Debug.Print myResult
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Primitive")
Private Sub Test02_ToStringPrimitive1_2_3()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "1,2,3"
    
    
    Dim myResult As String
    
    'Act:
    myResult = stringifier.ToString(1, 2, 3)
   Debug.Print myResult
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Iterable")
Private Sub Test03_IterableArray_1_2_3()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "[1,2,3]"
    
    
    Dim myResult As String
    stringifier.SetArrayMarkup
    'Act:
    myResult = stringifier.pvStringifyIterable(Array(1, 2, 3))
    Debug.Print myResult
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Iterable")
Private Sub Test04_NoBracketsIterableArray_1_2_3()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "1,2,3"
    
    
    Dim myResult As String
    stringifier.SetArrayMarkup Char.NoString, Char.NoString, Char.comma
    'Act:
    myResult = stringifier.pvStringifyIterable(Array(1, 2, 3))
    Debug.Print myResult
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Iterable")
Private Sub Test05_Collection_1_2_3()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "{1,2,3}"
    
    
    Dim myResult As String
    Dim myColl As Collection
    Set myColl = New Collection
    myColl.Add 1
    myColl.Add 2
    myColl.Add 3
    'Act:
    myResult = stringifier.pvStringifyIterable(myColl)
    Debug.Print myResult
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Iterable")
Private Sub Test06_Dictionary()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "{""Hello"" 1,""2"" There,""World"" 3}"
    
    
    Dim myResult As String
    Dim mySD As Scripting.Dictionary
    Set mySD = New Scripting.Dictionary
    mySD.Add "Hello", 1
    mySD.Add 2, "There"
    mySD.Add "World", 3
    'Act:
    myResult = stringifier.pvStringifyIterable(mySD)
    Debug.Print myExpected
    Debug.Print myResult
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Iterable")
Private Sub Test07_DictionaryToString()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "{""Hello"" 1,""2"" There,""World"" 3}"
    
    
    Dim myResult As String
    Dim mySD As Scripting.Dictionary
    Set mySD = New Scripting.Dictionary
    mySD.Add "Hello", 1
    mySD.Add 2, "There"
    mySD.Add "World", 3
    'Act:
    myResult = stringifier.ToString(mySD)
    Debug.Print myExpected
    Debug.Print myResult
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Admin")
Private Sub Test08_Empty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "Empty"
    
    
    Dim myResult As String
    
    myResult = stringifier.pvStringifyAdmin(Empty)
    Debug.Print myExpected
    Debug.Print myResult
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Admin")
Private Sub Test09_ArrayOfEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "[Empty,Empty,Empty]"
    
    stringifier.SetArrayMarkup
    Dim myResult As String
    
    myResult = stringifier.pvStringifyAdmin(Array(Empty, Empty, Empty))
    Debug.Print myExpected
    Debug.Print myResult
    'Assert:
    Assert.AreEqual myExpected, myResult
    
TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub
