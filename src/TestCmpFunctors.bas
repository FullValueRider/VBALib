Attribute VB_Name = "TestCmpFunctors"
'@TestModule
'@Folder("Tests")
'@IgnoreModule

Option Explicit
Option Private Module

'Private Assert As Object
'Private Fakes As Object

#If TWINBASIC Then
    'Do nothing
#Else


    '@ModuleInitialize
Private Sub ModuleInitialize()
    GlobalAssert
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
'Functor: Whats a better name for a class with only one function?
Public Sub CmpFunctorTests()

    #If TWINBASIC Then
        Debug.Print CurrentProcedureName;
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName;
    #End If

    Test01a_CmpEq_Numbers
    Test01b_CmpEq_String
    Test01c_CmpEq_Boolean
    Test01d_CmpEq_Array
    Test01e_CmpEq_Seq
    Test01f_CmpEq_Kvp
    
    Test02a_CmpNEq_Numbers
    Test02b_CmpNEq_String
    Test02c_CmpNEq_Boolean
    Test02d_CmpNEq_Array
    Test02e_CmpNEq_Seq
    Test02f_CmpNEq_Kvp

    Test03a_CmpMT_Numbers
    Test03b_CmpMT_String
    Test03c_CmpMT_Boolean
    Test03d_CmpMT_Array
    Test03e_CmpMT_Seq
    Test03f_CmpMT_Kvp

    Test04a_CmpMTEq_Numbers
    Test04b_CmpMTEq_String
    Test04c_CmpMTEq_Boolean
    Test04d_CmpMTEq_Array
    Test04e_CmpMTEq_Seq
    Test04f_CmpMTEq_Kvp


    Test05a_CmpLT_Numbers
    Test05b_CmpLT_String
    Test05c_CmpLT_Boolean
    Test05d_CmpLT_Array
    Test05e_CmpLT_Seq
    Test05f_CmpLT_Kvp

    Test06a_CmpLTEQ_Numbers
    Test06b_CmpLTEQ_String
    Test06c_CmpLTEQ_Boolean
    Test06d_CmpLTEQ_Array
    Test06e_CmpLTEQ_Seq
    Test06f_CmpLTEQ_Kvp

    Debug.Print vbTab, vbTab, "Testing completed"

End Sub


'@TestMethod("CmpFunctor")
Private Sub Test01a_CmpEq_Numbers()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, True, True, True, True, True, False, False, False, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpEQ(42&)
    
    myresult(1) = myCmp.ExecCmp(42)
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(42)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(True)
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test01b_CmpEq_String()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(False, False, False, False, False, False, True, False, False, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpEQ("42")
    
    myresult(1) = myCmp.ExecCmp("43")
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(42)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(True)
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test01c_CmpEq_Boolean()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, False, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpEQ(True)
    
    myresult(1) = myCmp.ExecCmp(True)
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(False)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(True)
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test01d_CmpEq_Array()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, True, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpEQ(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    myresult(1) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(10, 20, 30, 40, 50, 60, 70, 80, 90))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    ' comparers do not compare differentiate container ttype only content
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(10) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test01e_CmpEq_Seq()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, True, False, False, False, False, False, False, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpEQ(SeqA.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    myresult(1) = myCmp.ExecCmp(SeqA.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqC.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 11))
    myresult(10) = myCmp.ExecCmp(SeqL.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test01f_CmpEq_Kvp()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, False, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpEQ(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    
    myresult(1) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(KvpHA.Deb.AddPairs(SeqA("Hundred", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    myresult(10) = myCmp.ExecCmp(SeqL(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test02a_CmpNEq_Numbers()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(False, False, False, False, False, False, True, True, True, True, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpNEQ(42&)
    
    myresult(1) = myCmp.ExecCmp(42)
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(42)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(True)
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test02b_CmpNEq_String()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(False, True, True, True, True, True, False, True, True, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpNEQ("42")
    
    myresult(1) = myCmp.ExecCmp("42")
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(42)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(True)
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test02c_CmpNEq_Boolean()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(False, True, True, True, True, True, True, True, True, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpNEQ(True)
    
    myresult(1) = myCmp.ExecCmp(True)
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(False)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(True)
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test02d_CmpNEq_Array()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(False, True, True, True, True, True, True, True, False, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpNEQ(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    myresult(1) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(10, 20, 30, 40, 50, 60, 70, 80, 90))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(10) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test02e_CmpNEq_Seq()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(False, True, False, True, True, True, True, True, True, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpNEQ(SeqA.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    myresult(1) = myCmp.ExecCmp(SeqA.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqC.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 11))
    myresult(10) = myCmp.ExecCmp(SeqL.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test02f_CmpNEq_Kvp()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(False, True, True, True, True, True, True, True, True, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpNEQ(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    
    myresult(1) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("Hundred", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    myresult(10) = myCmp.ExecCmp(SeqL.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub Test03a_CmpMT_Numbers()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, False, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpMT(42&)
    
    myresult(1) = myCmp.ExecCmp(43)                         ' True
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))              ' False
    myresult(3) = myCmp.ExecCmp(True)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(43)
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test03b_CmpMT_String()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, False, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpMT("42")
    
    myresult(1) = myCmp.ExecCmp("43")
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(42)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp(True)
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp("43")
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test03c_CmpMT_Boolean()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(False, False, False, False, False, False, False, False, False, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpMT(True)
    
    myresult(1) = myCmp.ExecCmp(True)
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(False)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(True)
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test03d_CmpMT_Array()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, False, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpMT(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    myresult(1) = myCmp.ExecCmp(Array(2, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 1, 1, 4, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(10) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 9, 9, 11))
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test03e_CmpMT_Seq()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, True, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpMT(SeqA.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    myresult(1) = myCmp.ExecCmp(SeqA.Deb(2, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 2, 2, 2, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    ' 4 > 1 so the next test returns true
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqC.Deb(1, 2, 3, 4, 5, 6, 6, 8, 9, 11))
    myresult(10) = myCmp.ExecCmp(SeqL.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 9))
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test03f_CmpMT_Kvp()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, False, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpMT(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    
    myresult(1) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 6, 6, 7, 8, 9, 11)))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    ' Hundred is longer than one so next test should be true
    myresult(9) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("Hundred", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    ' next item is false because Seven is longer than sixe
    myresult(10) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Sixe", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test04a_CmpMTEq_Numbers()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, True, True, True, True, False, False, False, False, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpMTEQ(42&)
    
    myresult(1) = myCmp.ExecCmp(42)
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(43)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(43))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(41))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(True)
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test04b_CmpMTEq_String()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, True, False, False, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpMTEQ("42")
    
    myresult(1) = myCmp.ExecCmp("43")
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp("41")
    myresult(4) = myCmp.ExecCmp("121")
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(True)
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test04c_CmpMTEq_Boolean()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, False, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpMTEQ(True)
    
    myresult(1) = myCmp.ExecCmp(True)
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(False)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(True)
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test04d_CmpMTEq_Array()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, True, True, False, False, False, False, True, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpMTEQ(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    myresult(1) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100))
    myresult(4) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    ' comparers do not compare differentiate container ttype only content
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(10) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test04e_CmpMTEq_Seq()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, True, False, False, False, False, False, True, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpMTEQ(SeqA.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    myresult(1) = myCmp.ExecCmp(SeqA.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqC.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 11))
    myresult(10) = myCmp.ExecCmp(SeqL.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test04f_CmpMTEq_Kvp()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, True, False, False, False, False, False, False, True, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpMTEQ(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    
    myresult(1) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    myresult(2) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(2, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    myresult(3) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(KvpHA.Deb.AddPairs(SeqA("Hundred", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)))
    myresult(10) = myCmp.ExecCmp(KvpHA.Deb.AddPairs(SeqA("Hundred", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9)))
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub Test05a_CmpLT_Numbers()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, True, False, False, False, False, False, False, False, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpLT(42&)
    
    myresult(1) = myCmp.ExecCmp(41)
    myresult(2) = myCmp.ExecCmp(VBA.CByte(41))
    myresult(3) = myCmp.ExecCmp(True)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(43)
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test05b_CmpLT_String()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, True, False, False, False, False, False, False, False, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpLT("42")
    
    myresult(1) = myCmp.ExecCmp("41")
    myresult(2) = myCmp.ExecCmp("121")
    myresult(3) = myCmp.ExecCmp(42)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp(True)
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp("43")
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test05c_CmpLT_Boolean()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(False, False, False, False, False, False, False, False, False, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpLT(True)
    
    myresult(1) = myCmp.ExecCmp(True)
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(False)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(True)
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test05d_CmpLT_Array()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, True, False, False, False, False, True, False, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpLT(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    myresult(1) = myCmp.ExecCmp(Array(-1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 1, 1, 4, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(10) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 9, 9, 11))
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test05e_CmpLT_Seq()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, True, False, False, False, False, True, False, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpLT(SeqA.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    myresult(1) = myCmp.ExecCmp(SeqA.Deb(-1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 2, 2, 2, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqC.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 11))
    myresult(10) = myCmp.ExecCmp(SeqL.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 9))
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test05f_CmpLT_Kvp()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, True, False)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpLT(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    
    myresult(1) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(-1, 2, 3, 4, 6, 6, 7, 8, 9, 10)))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("Hundred", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    myresult(10) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Sixe", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub Test06a_CmpLTEQ_Numbers()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, True, True, True, True, True, False, False, False, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpLTEQ(42&)
    
    myresult(1) = myCmp.ExecCmp(41)
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(-10)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(0.1)
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test06b_CmpLTEQ_String()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, True, False, False, False, False, False, False, False, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpLTEQ("42")
    
    myresult(1) = myCmp.ExecCmp("41")
    myresult(2) = myCmp.ExecCmp("4")
    myresult(3) = myCmp.ExecCmp("422")
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp(True)
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp("42")
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test06c_CmpLTEQ_Boolean()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, False, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpLTEQ(True)
    
    myresult(1) = myCmp.ExecCmp(True)
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(False)
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(42))
    myresult(10) = myCmp.ExecCmp(True)
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test06d_CmpLTEQ_Array()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, True, False, False, False, False, True, True, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpLTEQ(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    myresult(1) = myCmp.ExecCmp(Array(0, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(2, 4, 6, 4, 5, 6, 7, 8, 9))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqA.Deb.AddItems(1, 2, 3, 4, 5, 6, 7, 8, 9, 9))
    myresult(10) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 9, 9, 11))
    
    
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("CmpFunctor")
Private Sub Test06e_CmpLTEQ_Seq()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, True, False, False, False, False, True, True, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpLTEQ(SeqA.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    
    myresult(1) = myCmp.ExecCmp(SeqA.Deb(0, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 2, 2, 2, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(SeqC.Deb(1, 2, 3, 4, 5, 6, 6, 8, 9, 10))
    myresult(10) = myCmp.ExecCmp(SeqL.Deb(1, 2, 3, 4, 5, 6, 7, 8, 9, 9))
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("CmpFunctor")
Private Sub Test06f_CmpLTEQ_Kvp()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, False, False, False, False, False, False, False, True, True)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myresult As Variant
    ReDim myresult(1 To 10)
    
    Dim myCmp As IComparer
    Set myCmp = cmpLTEQ(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    
    myresult(1) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 6, 6, 7, 8, 9, -1)))
    myresult(2) = myCmp.ExecCmp(VBA.CByte(42))
    myresult(3) = myCmp.ExecCmp(Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10))
    myresult(4) = myCmp.ExecCmp(VBA.CDec(42))
    myresult(5) = myCmp.ExecCmp(VBA.CLngLng(42))
    myresult(6) = myCmp.ExecCmp(VBA.CDate(42))
    myresult(7) = myCmp.ExecCmp("42")
    myresult(8) = myCmp.ExecCmp(Array(42))
    myresult(9) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("Hundred", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    myresult(10) = myCmp.ExecCmp(KvpA.Deb.AddPairs(SeqA("One", "Two", "Three", "Four", "Five", "Six", "Sixe", "Eight", "Nine", "Ten"), SeqA(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)))
    
    'Act:
    AssertStrictSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
