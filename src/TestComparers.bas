Attribute VB_Name = "TestComparers"
'@TestModule
'@Folder("Tests")
'@IgnoreModule

Option Explicit
Option Private Module

Private Assert As Object
'Private Fakes As Object


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    'Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub


'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub


'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub




'@TestMethod("Comparer")
Private Sub Test01a_CmpEq_Long_True()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpEQ(42)
    
    myresult = myCmp.ExecCmp(42)
    
    'Act:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test01b_CmpEq_Long_False()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpEQ(42)
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = myCmp.ExecCmp(43)
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

'@TestMethod("Comparer")
Private Sub Test02a_CmpNEq_Long_True()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpNEQ(42)
    
    myresult = myCmp.ExecCmp(42)
    
    'Act:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test02b_CmpNEq_Long_False()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpNEQ(42)
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = myCmp.ExecCmp(43)
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


'@TestMethod("Comparer")
Private Sub Test03a_CmpMT_Long_True()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpMT(42)
    
    myresult = myCmp.ExecCmp(43)
    
    'Act:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test03b_CmpMT_Long_False()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpMT(42)
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = myCmp.ExecCmp(42)
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

'@TestMethod("Comparer")
Private Sub Test04a_CmpLT_Long_True()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpLT(42)
    
    myresult = myCmp.ExecCmp(41)
    
    'Act:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test04b_CmpLT_Long_False()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpLT(42)
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = myCmp.ExecCmp(42)
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

'@TestMethod("Comparer")
Private Sub Test05a_CmpMTEQ_Long_MTTrue()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpMTEQ(42)
    
    myresult = myCmp.ExecCmp(43)
    
    'Act:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test05b_CmpMTEQ_Long_EQTrue()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpMTEQ(42)
    
    myresult = myCmp.ExecCmp(42)
    
    'Act:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Comparer")
Private Sub Test04c_CmpMTEQ_Long_False()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpMTEQ(42)
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = myCmp.ExecCmp(41)
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

'@TestMethod("Comparer")
Private Sub Test06a_CmpLTEQ_Long_LTTrue()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpLTEQ(42)
    
    myresult = myCmp.ExecCmp(41)
    
    'Act:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test06b_CmpLTEQ_Long_EQTrue()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpLTEQ(42)
    
    myresult = myCmp.ExecCmp(42)
    
    'Act:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Comparer")
Private Sub Test06c_CmpLTEQ_Long_False()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    
    Dim myresult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpLTEQ(42)
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult = myCmp.ExecCmp(43)
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
