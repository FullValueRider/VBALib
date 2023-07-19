Attribute VB_Name = "TestComparers"
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
    'Set Assert = CreateObject("Rubberduck.AssertClass")
    'Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    'Set Assert = Nothing
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

#End If

Public Sub ComparerTests()

#If twinbasic Then
    Debug.Print CurrentProcedureName;
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName;
#End If

    Test01a_CmpEq_Long_True
    Test01b_CmpEq_Long_False
    
    Test02a_CmpNEq_Long_True
    Test02b_CmpNEq_Long_False
    
    Test03a_CmpMT_Long_True
    Test03b_CmpMT_Long_False
    
    Test04a_CmpLT_Long_True
    Test04b_CmpLT_Long_False
    
    Test05a_CmpMTEQ_Long_MTTrue
    Test05b_CmpMTEQ_Long_EQTrue
    Test05c_CmpMTEQ_Long_False
    
    Test06a_CmpLTEQ_Long_LTTrue
    Test06b_CmpLTEQ_Long_EQTrue
    Test06c_CmpLTEQ_Long_False
    
    
    Debug.Print vbTab, vbTab, "Testing completed"

End Sub
    

'@TestMethod("Comparer")
Private Sub Test01a_CmpEq_Long_True()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpEQ(42)
    
    myResult = myCmp.ExecCmp(42)
    
    'Act:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test01b_CmpEq_Long_False()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpEQ(42)
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = myCmp.ExecCmp(43)
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Comparer")
Private Sub Test02a_CmpNEq_Long_True()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpNEQ(42)
    
    myResult = myCmp.ExecCmp(42)
    
    'Act:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test02b_CmpNEq_Long_False()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpNEQ(42)
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = myCmp.ExecCmp(43)
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test03a_CmpMT_Long_True()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpMT(42)
    
    myResult = myCmp.ExecCmp(43)
    
    'Act:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test03b_CmpMT_Long_False()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpMT(42)
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = myCmp.ExecCmp(42)
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Comparer")
Private Sub Test04a_CmpLT_Long_True()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpLT(42)
    
    myResult = myCmp.ExecCmp(41)
    
    'Act:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test04b_CmpLT_Long_False()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpLT(42)
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = myCmp.ExecCmp(42)
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Comparer")
Private Sub Test05a_CmpMTEQ_Long_MTTrue()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpMTEQ(42)
    
    myResult = myCmp.ExecCmp(43)
    
    'Act:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test05b_CmpMTEQ_Long_EQTrue()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpMTEQ(42)
    
    myResult = myCmp.ExecCmp(42)
    
    'Act:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Comparer")
Private Sub Test05c_CmpMTEQ_Long_False()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpMTEQ(42)
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = myCmp.ExecCmp(41)
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Comparer")
Private Sub Test06a_CmpLTEQ_Long_LTTrue()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpLTEQ(42)
    
    myResult = myCmp.ExecCmp(41)
    
    'Act:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Comparer")
Private Sub Test06b_CmpLTEQ_Long_EQTrue()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpLTEQ(42)
    
    myResult = myCmp.ExecCmp(42)
    
    'Act:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Comparer")
Private Sub Test06c_CmpLTEQ_Long_False()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
       myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
      On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    
    Dim myCmp As IComparer
    Set myCmp = cmpLTEQ(42)
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult = myCmp.ExecCmp(43)
    'Assert:
    AssertStrictAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
