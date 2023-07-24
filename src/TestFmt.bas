Attribute VB_Name = "TestFmt"
'@TestModule
'@Folder("Tests")
'@IgnoreModule

Option Explicit
Option Private Module

'Private Assert As Object
'Private Fakes As Object

#If twinbasic Then
    'Do nothing
#Else

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
'    Set Assert = CreateObject("Rubberduck.AssertClass")
'    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
'    Set Assert = Nothing
'    Set Fakes = Nothing
End Sub


'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub


'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

#End If

Public Sub FmtTests()

    
#If twinbasic Then
    Debug.Print CurrentProcedureName; vbTab, vbTab,
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab, vbTab,
#End If

    Test01a_Fmt_Text_Nothing
    Test01b_Fmt_Text_NoParams
    Test01c_Fmt_Text_NoSubstitutions
    Test01d_Fmt_Text_Formatting_Threevbcrlf
    Test01e_Fmt_Text_Formatting_Threeplainquotes
    Test01f_Fmt_Text_Formatting_Zeroplainquotes
    Test01g_Fmt_Text_Formatting_ThreeVariables
    
    Debug.Print "Testing completed"

End Sub
    
'@TestMethod("Fmt")
Private Sub Test01a_Fmt_Text_Nothing()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    '@Ignore EmptyStringLiteral
    myExpected = ""
    
    Dim myResult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Fmt.Text(vbNullString)
    'Assert:
   
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    'On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Fmt")
Private Sub Test01b_Fmt_Text_NoParams()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello {0} World{0}"
    
    Dim myResult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Fmt.Text("Hello {0} World{0}")
    'Assert:
   
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    'On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Fmt")
Private Sub Test01c_Fmt_Text_NoSubstitutions()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello World"
    
    Dim myResult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Fmt.Text("Hello World", 1, "One", 3.142)
    'Assert:
   
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    'On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Fmt")
Private Sub Test01d_Fmt_Text_Formatting_Threevbcrlf()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello " & vbCrLf & vbCrLf & vbCrLf & " World"
    
    Dim myResult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Fmt.Text("Hello {nl3} World", 1, "One", 3.142)
    'Assert:
   
    AssertStrictAreEqual VBA.Len(myExpected), VBA.Len(myResult), myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    'On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Fmt")
Private Sub Test01e_Fmt_Text_Formatting_Threeplainquotes()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello ''' World"
    
    Dim myResult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Fmt.Text("Hello {sq3} World", 1, "One", 3.142)
    'Assert:
   
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    'On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Fmt")
Private Sub Test01f_Fmt_Text_Formatting_Zeroplainquotes()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello  World"
    
    Dim myResult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Fmt.Text("Hello {sq0} World", 1, "One", 3.142)
    'Assert:
   
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    'On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Fmt")
Private Sub Test01g_Fmt_Text_Formatting_ThreeVariables()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    

    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello 1 One World [1,2,3] {3}"
    
    Dim myResult As String
    Dim myC As KvpC
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = Fmt.Text("Hello {0} {1} World {2} {3}", 1, "One", Array(1, 2, 3))
    'Assert:
   
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    'On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

