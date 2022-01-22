Attribute VB_Name = "TestStringifier"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")
'@IgnoreModule

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
    Private Fakes As Object
#Else
    Private Assert As Rubberduck.AssertClass
    Private Fakes As Rubberduck.FakesProvider
#End If

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.AssertClass")
        Set Fakes = CreateObject("Rubberduck.FakesProvider")
    #Else
        Set Assert = New Rubberduck.AssertClass
        Set Fakes = New Rubberduck.FakesProvider
    #End If
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


Public Sub StringifierTests()

    myInterim = Timer
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Debug.Print "Testing ", ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
    
    Test01_Primitive1
    Test02_ToStringPrimitive1_2_3
    Test03_IterableArray_1_2_3
    Test04_NoBracketsIterableArray_1_2_3
    Test05_Collection_1_2_3
    Test06_Dictionary
    Test07_DictionaryToString
    Test08_Empty
    Test09_ArrayOfEmpty
    
    Debug.Print "completed", Timer - myInterim

End Sub


'@TestMethod("Primitive")
Public Sub Test01_Primitive1()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "1"
    
    
    Dim myResult As String
    
    'Act:
    myResult = Stringifier.StringifyPrimitive(1)

    'Assert:
    Assert.AreEqual myExpected, myResult  ',  ErrEx.LiveCallstack.ProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print vbCrLf & ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Primitive")
Public Sub Test02_ToStringPrimitive1_2_3()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "1,2,3"
    
    
    Dim myResult As String
    
    'Act:
    myResult = Stringifier.ToString(1, 2, 3)

    'Assert:
    Assert.AreEqual myExpected, myResult  ',  ErrEx.LiveCallstack.ProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print vbCrLf & ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Iterable")
Public Sub Test03_IterableArray_1_2_3()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "[1,2,3]"
    
    
    Dim myResult As String
    Stringifier.ResetArrayMarkup
    'Act:
    myResult = Stringifier.StringifyIterable(Array(1, 2, 3))
    
    'Assert:
    Assert.AreEqual myExpected, myResult  ',  ErrEx.LiveCallstack.ProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print vbCrLf & ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Iterable")
Public Sub Test04_NoBracketsIterableArray_1_2_3()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "1,2,3"
    
    
    Dim myResult As String
    Stringifier.ResetArrayMarkup Char.twNoString, Char.twNoString, Char.twComma
    'Act:
    myResult = Stringifier.StringifyIterable(Array(1, 2, 3))
    
    'Assert:
    Assert.AreEqual myExpected, myResult  ',  ErrEx.LiveCallstack.ProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print vbCrLf & ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Iterable")
Public Sub Test05_Collection_1_2_3()
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
    myResult = Stringifier.StringifyIterable(myColl)
    
    'Assert:
    Assert.AreEqual myExpected, myResult  ',  ErrEx.LiveCallstack.ProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print vbCrLf & ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub


'@TestMethod("Iterable")
Public Sub Test06_Dictionary()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "{""1"": Hello,""2"": There,""3"": World}"
    
    Dim myResult As String
    Dim mySD As Scripting.Dictionary
    Set mySD = New Scripting.Dictionary
    mySD.Add 1, "Hello"
    mySD.Add 2, "There"
    mySD.Add 3, "World"
    
    'Act:
    myResult = Stringifier.StringifyIterable(mySD)
    
    'Assert:
    Assert.AreEqual myExpected, myResult  ',  ErrEx.LiveCallstack.ProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print vbCrLf & ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Iterable")
Public Sub Test07_DictionaryToString()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "{""1"": Hello,""2"": There,""3"": World}"
    
    Dim myResult As String
    Dim mySD As Scripting.Dictionary
    Set mySD = New Scripting.Dictionary
    mySD.Add 1, "Hello"
    mySD.Add 2, "There"
    mySD.Add 3, "World"
    
    'Act:
    myResult = Stringifier.ToString(mySD)
    
    'Assert:
    Assert.AreEqual myExpected, myResult  ',  ErrEx.LiveCallstack.ProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print vbCrLf & ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Admin")
Public Sub Test08_Empty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "Empty"
    
    Dim myResult As String
    myResult = Stringifier.StringifyAdmin(Empty)
    
    'Assert:
    Assert.AreEqual myExpected, myResult  ',  ErrEx.LiveCallstack.ProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print vbCrLf & ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("Admin")
Public Sub Test09_ArrayOfEmpty()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As String
    myExpected = "[Empty,Empty,Empty]"
    
    Stringifier.ResetArrayMarkup
    Dim myResult As String
    myResult = Stringifier.ToString(Array(Empty, Empty, Empty))
    
    'Assert:
    Assert.AreEqual myExpected, myResult  ',  ErrEx.LiveCallstack.ProcedureName
    
TestExit:
    Exit Sub
    
TestFail:
    Debug.Print vbCrLf & ErrEx.LiveCallstack.ProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

