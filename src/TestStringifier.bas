Attribute VB_Name = "TestStringifier"
'@TestModule
'@Folder("Tests")
'@IgnoreModule


Option Explicit
Option Private Module

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


'@TestMethod("Stringifier")
Private Sub Test01_StringifyItem_String()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello World!"
    
    Dim myresult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = Stringifier.StringifyItem("Hello World!")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test02_StringifyItem_Long()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "42"
    
    Dim myresult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = Stringifier.StringifyItem(42)
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

''@TestMethod("Stringifier.049")
'Private Sub Test49c_StringifyItem_Long()
'
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected As String
'    myExpected = "42"
'
'    Dim myresult As String
'
'    'Act:  Again we need to sort The result SeqC to get the matching array
'    myresult = Stringifier.StringifyItem(42)
'    'Assert:
'    Assert.AreEqual myExpected, myresult
'
'TestExit:
'    '@Ignore UnhandledOnErrorResumeNext
'    On Error Resume Next
'
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub


'@TestMethod("Stringifier")
Private Sub Test03_StringifyItem_Array()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "[1,2,3,4,5,6]"
    
    Dim myresult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = Stringifier.StringifyItem(Array(1, 2, 3, 4, 5, 6))
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test04_StringifyItem_SeqC()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "{1,2,3,4,5,6}"
    
    Dim myresult As String
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Stringifier.ResetMarkup
    myresult = Stringifier.StringifyItem(SeqC(1, 2, 3, 4, 5, 6))
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test05_StringifyItem_Collection()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "{1,2,3,4,5,6}"
    
    Dim myresult As String
    Dim myC As Collection
    Set myC = New Collection
    myC.Add 1
    myC.Add 2
    myC.Add 3
    myC.Add 4
    myC.Add 5
    myC.Add 6
    'Act:  Again we need to sort The result SeqC to get the matching array
    Stringifier.ResetMarkup
    myresult = Stringifier.StringifyItem(myC)
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test06_StringifyItem_Dictionary()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "{ 'One': 1, 'Two': 2, 'Three': 3, 'Four': 4, 'Five': 5, 'Six': 6}"
    
    Dim myresult As String
    Dim myC As KvpC
    Set myC = KvpC.Deb
    myC.Add "One", 1
    myC.Add "Two", 2
    myC.Add "Three", 3
    myC.Add "Four", 4
    myC.Add "Five", 5
    myC.Add "Six", 6
    'Act:  Again we need to sort The result SeqC to get the matching array
    Stringifier.ResetMarkup
    myresult = Stringifier.StringifyItem(myC)
    'Assert:
    'Debug.Print myExpected
    'Debug.Print myResult
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Stringifier")
Private Sub Test07_StringifyItem_CustomDictionaryMarkup()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "{|One|_1?|Two|_2?|Three|_3?|Four|_4?|Five|_5?|Six|_6}"
    
    Dim myresult As String
    Dim myC As KvpC
    Set myC = KvpC.Deb
    myC.Add "One", 1
    myC.Add "Two", 2
    myC.Add "Three", 3
    myC.Add "Four", 4
    myC.Add "Five", 5
    myC.Add "Six", 6
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myToString As Stringifier
    Set myToString = Stringifier.Deb.SetObjectMarkup(ipSeparator:="?").SetDictionaryItemMarkup("|", "_", "|")
    myresult = myToString.StringifyItem(myC)
    'Assert:
   'Debug.Print myExpected
    'Debug.Print myResult
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
