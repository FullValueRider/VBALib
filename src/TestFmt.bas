Attribute VB_Name = "TestFmt"
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




'@TestMethod("Fmt")
Private Sub Test50a_Fmt_Text_Nothing()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    '@Ignore EmptyStringLiteral
    myExpected = ""
    
    Dim myresult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = Fmt.Text(vbNullString)
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

'@TestMethod("Fmt")
Private Sub Test50b_Fmt_Text_NoParams()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello {0} World{0}"
    
    Dim myresult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = Fmt.Text("Hello {0} World{0}")
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

'@TestMethod("Fmt")
Private Sub Test50c_Fmt_Text_NoSubstitutions()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello World"
    
    Dim myresult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = Fmt.Text("Hello World", 1, "One", 3.142)
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

'@TestMethod("Fmt")
Private Sub Test50d_Fmt_Text_Formatting_Threevbcrlf()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello " & vbCrLf & vbCrLf & vbCrLf & " World"
    
    Dim myresult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = Fmt.Text("Hello {nl3} World", 1, "One", 3.142)
    'Assert:
   
    Assert.AreEqual VBA.Len(myExpected), VBA.Len(myresult)
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Fmt")
Private Sub Test50e_Fmt_Text_Formatting_Threeplainquotes()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello ''' World"
    
    Dim myresult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = Fmt.Text("Hello {sq3} World", 1, "One", 3.142)
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

'@TestMethod("Fmt")
Private Sub Test50f_Fmt_Text_Formatting_Zeroplainquotes()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello  World"
    
    Dim myresult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = Fmt.Text("Hello {sq0} World", 1, "One", 3.142)
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

'@TestMethod("Fmt")
Private Sub Test50g_Fmt_Text_Formatting_ThreeVariables()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello 1 One World [1,2,3] {3}"
    
    Dim myresult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = Fmt.Text("Hello {0} {1} World {2} {3}", 1, "One", Array(1, 2, 3))
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

