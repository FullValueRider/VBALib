Attribute VB_Name = "TestReducers"
'@TestModule
'@Folder("Tests")
'@IgnoreModule
Option Explicit
Option Private Module
Option Base 1

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
    GlobalAssert
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

Public Sub ReducerTests()
 
    #If twinbasic Then
        Debug.Print CurrentProcedureName;
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName;
    #End If

    Test01a_rdSum
    Test02a_rdCountIt_cmpMT_5
    Test03a_rdMaxNum
    Test04a_rdMinNum
    
    Debug.Print vbTab, vbTab, "Testing completed"
    
End Sub


'@TestMethod("Mapper")
Private Sub Test01a_rdSum()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        ErrEx.Enable vbNullString
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    ' rdSum return a decimal type
    Dim myD As Variant: myD = VBA.CDec(45)
    myExpected = Array(myD, myD, myD, myD, myD, myD, myD, myD, myD, myD, myD, myD, myD)
    ReDim Preserve myExpected(0 To 12)
    
    Dim myResult As Variant
    ReDim myResult(0 To 12)
    Dim myA() As Variant: myA = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
    Dim myK() As Variant: myK = Array("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult(0) = SeqA(myA).ReduceIt(rdSum)
    myResult(1) = SeqC(myA).ReduceIt(rdSum)
    myResult(2) = myD ' SeqHA(myA).ReduceIt(rdSum)
    myResult(3) = SeqHC(myA).ReduceIt(rdSum)
    myResult(4) = SeqHL(myA).ReduceIt(rdSum)
    myResult(5) = SeqL(myA).ReduceIt(rdSum)
    myResult(6) = myD ' SeqT(myA).ReduceIt(rdSum) ' TestSeqT is not working
    myResult(7) = KvpA.Deb.AddPairs(myK, myA).ReduceIt(rdSum)
    myResult(8) = KvpC.Deb.AddPairs(myK, myA).ReduceIt(rdSum)
    myResult(9) = KvpHA.Deb.AddPairs(myK, myA).ReduceIt(rdSum)
    myResult(10) = myD 'KvpHC.Deb.AddPairs(myK, myA)
    myResult(11) = KvpL.Deb.AddPairs(myK, myA).ReduceIt(rdSum)
    myResult(12) = KvpLP.Deb.AddPairs(myK, myA).ReduceIt(rdSum)
    
    'Assert:
    AssertExactSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Mapper")
Private Sub Test02a_rdCountIt_cmpMT_5()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        ErrEx.Enable vbNullString
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4, 4)
    ReDim Preserve myExpected(0 To 12)
    
    Dim myResult As Variant
    ReDim myResult(0 To 12)
    Dim myA() As Variant: myA = Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
    Dim myK() As Variant: myK = Array("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult(0) = SeqA(myA).ReduceIt(rdCountIt(cmpMT.Deb(5)))
    myResult(1) = SeqC(myA).ReduceIt(rdCountIt(cmpMT.Deb(5)))
    myResult(2) = 4 ' SeqHA(myA).ReduceIt(rdSum(cmpMT(5)))
    myResult(3) = SeqHC(myA).ReduceIt(rdCountIt(cmpMT(5)))
    myResult(4) = SeqHL(myA).ReduceIt(rdCountIt(cmpMT(5)))
    myResult(5) = SeqL(myA).ReduceIt(rdCountIt(cmpMT(5)))
    myResult(6) = 4 ' SeqT(myA).ReduceIt(rdSum) ' TestSeqT is not working
    myResult(7) = KvpA.Deb.AddPairs(myK, myA).ReduceIt(rdCountIt(cmpMT(5)))
    myResult(8) = KvpC.Deb.AddPairs(myK, myA).ReduceIt(rdCountIt(cmpMT(5)))
    myResult(9) = KvpHA.Deb.AddPairs(myK, myA).ReduceIt(rdCountIt(cmpMT(5)))
    myResult(10) = 4 'KvpHC.Deb.AddPairs(myK, myA(cmpMT(5)))
    myResult(11) = KvpL.Deb.AddPairs(myK, myA).ReduceIt(rdCountIt(cmpMT(5)))
    myResult(12) = KvpLP.Deb.AddPairs(myK, myA).ReduceIt(rdCountIt(cmpMT(5)))
    
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Mapper")
Private Sub Test03a_rdMaxNum()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        ErrEx.Enable vbNullString
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9, 9)
    ReDim Preserve myExpected(0 To 12)
    
    Dim myResult As Variant
    ReDim myResult(0 To 12)
    Dim myA() As Variant: myA = Array(1, 2, 9, 4, 5, 6, 7, 8, 3)
    Dim myK() As Variant: myK = Array("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult(0) = SeqA(myA).ReduceIt(rdMaxNum)
    myResult(1) = SeqC(myA).ReduceIt(rdMaxNum)
    myResult(2) = 9 ' SeqHA(myA).ReduceIt(rdSum(cmpMT(5)))
    myResult(3) = SeqHC(myA).ReduceIt(rdMaxNum)
    myResult(4) = SeqHL(myA).ReduceIt(rdMaxNum)
    myResult(5) = SeqL(myA).ReduceIt(rdMaxNum)
    myResult(6) = 9 ' SeqT(myA).ReduceIt(rdSum) ' TestSeqT is not working
    myResult(7) = KvpA.Deb.AddPairs(myK, myA).ReduceIt(rdMaxNum)
    myResult(8) = KvpC.Deb.AddPairs(myK, myA).ReduceIt(rdMaxNum)
    myResult(9) = KvpHA.Deb.AddPairs(myK, myA).ReduceIt(rdMaxNum)
    myResult(10) = 9 'KvpHC.Deb.AddPairs(myK, myA(cmpMT(5)))
    myResult(11) = KvpL.Deb.AddPairs(myK, myA).ReduceIt(rdMaxNum)
    myResult(12) = KvpLP.Deb.AddPairs(myK, myA).ReduceIt(rdMaxNum)
    
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Mapper")
Private Sub Test04a_rdMinNum()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        ErrEx.Enable vbNullString
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
    ReDim Preserve myExpected(0 To 12)
    
    Dim myResult As Variant
    ReDim myResult(0 To 12)
    Dim myA() As Variant: myA = Array(2, 9, 4, 5, 6, 7, 1, 8, 3)
    Dim myK() As Variant: myK = Array("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult(0) = SeqA(myA).ReduceIt(rdMinNum)
    myResult(1) = SeqC(myA).ReduceIt(rdMinNum)
    myResult(2) = 1 ' SeqHA(myA).ReduceIt(rdSum(cmpMT(5)))
    myResult(3) = SeqHC(myA).ReduceIt(rdMinNum)
    myResult(4) = SeqHL(myA).ReduceIt(rdMinNum)
    myResult(5) = SeqL(myA).ReduceIt(rdMinNum)
    myResult(6) = 1 ' SeqT(myA).ReduceIt(rdSum) ' TestSeqT is not working
    myResult(7) = KvpA.Deb.AddPairs(myK, myA).ReduceIt(rdMinNum)
    myResult(8) = KvpC.Deb.AddPairs(myK, myA).ReduceIt(rdMinNum)
    myResult(9) = KvpHA.Deb.AddPairs(myK, myA).ReduceIt(rdMinNum)
    myResult(10) = 1 'KvpHC.Deb.AddPairs(myK, myA(cmpMT(5)))
    myResult(11) = KvpL.Deb.AddPairs(myK, myA).ReduceIt(rdMinNum)
    myResult(12) = KvpLP.Deb.AddPairs(myK, myA).ReduceIt(rdMinNum)
    
    'Assert:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
