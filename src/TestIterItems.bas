Attribute VB_Name = "TestIterItems"
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
    GlobalAssert
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

Public Sub IterItemsTests()
    
    #If twinbasic Then
        Debug.Print CurrentProcedureName;
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName;
    #End If

    Test01a_IsObjectAndName
    
    Test02a_GetItem0Seq
    Test02b_GetItem0SeqAfterThreeMovenext
    Test02c_GetItem0SeqAfterThreeMoveNextTwoMovePrev
    
    Test03a_GetItemSeqAtOffset3
    Test03b_GetItemSeqAtOffsetMinus3
    Test03c_GetItemSeqIndexGreaterThanSize
    Test03d_GetItemSeqIndexDeforeIndex1
    
    Test04a_GetKeySeq
    Test04b_GetIndexSeq
    
    Test05a_GetItemArray
    Test05b_GetKeyArray
    Test05c_GetIndexArray
    
    Test06a_GetItemCollection
    Test06b_GetKeyCollection
    Test06c_GetIndexCollection
    
    Test07a_GetItemArrayList
    Test07b_GetKeyArrayList
    Test07c_GetIndexArrayList
    
    Test08a_GetItemDictionary
    Test08b_GetKeyDictionary
    Test08c_GetIndexDIctionary
    
    Test09a_GetIndexDictionary
    
    Debug.Print vbTab, vbTab, "Testing completed"

End Sub


'@TestMethod("IterItems")
Private Sub Test01a_IsObjectAndName()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, "IterItems", "IterItems")
    
    Dim myResult As Variant
    ReDim myResult(0 To 2)
    
    Dim myI As IterItems
    Set myI = IterItems(SeqC(1, 2, 3, 4, 5))
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult(0) = VBA.IsObject(myI)
    myResult(1) = VBA.TypeName(myI)
    myResult(2) = myI.TypeName
   
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


'@TestMethod("IterItems")
Private Sub Test02a_GetItem0Seq()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 1
    
    Dim myResult As Variant
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myS As SeqC
    Set myS = SeqC(1, 2, 3, 4, 5)
    Dim myI As IterItems
    Set myI = IterItems(myS)
       
    myResult = myI.CurItem(0)
   
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test02b_GetItem0SeqAfterThreeMovenext()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, True, True, 40)
    
    Dim myResult As Variant
    ReDim myResult(0 To 3)
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50))
    myResult(0) = myI.MoveNext
    myResult(1) = myI.MoveNext
    myResult(2) = myI.MoveNext
    myResult(3) = myI.CurItem(0)
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


'@TestMethod("IterItems")
Private Sub Test02c_GetItem0SeqAfterThreeMoveNextTwoMovePrev()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(True, True, True, True, True, 20)
    
    Dim myResult As Variant
    ReDim myResult(0 To 5)
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50))
    myResult(0) = myI.MoveNext
    myResult(1) = myI.MoveNext
    myResult(2) = myI.MoveNext
    myResult(3) = myI.MovePrev
    myResult(4) = myI.MovePrev
    myResult(5) = myI.CurItem(0)
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


'@TestMethod("IterItems")
Private Sub Test03a_GetItemSeqAtOffset3()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 80
    
    Dim myResult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50, 60, 70, 80, 90))
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myResult = myI.CurItem(3)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test03b_GetItemSeqAtOffsetMinus3()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 20
    
    Dim myResult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50, 60, 70, 80, 90))
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myResult = myI.CurItem(-3)
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test03c_GetItemSeqIndexGreaterThanSize()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = True
    
    Dim myResult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50, 60, 70, 80, 90))
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myResult = myI.CurItem(5)
    
    'Assert:
    AssertExactAreEqual myExpected, VBA.IsNull(myResult), myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test03d_GetItemSeqIndexDeforeIndex1()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = True
    
    Dim myResult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(SeqC(10, 20, 30, 40, 50, 60, 70, 80, 90))
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myResult = myI.CurItem(-5)
    
    'Assert:
    AssertExactAreEqual myExpected, VBA.IsNull(myResult), myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test04a_GetKeySeq()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myResult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myC As SeqC
    Set myC = SeqC(10, 20, 30, 40, 50, 60, 70, 80, 90)
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myResult = myI.CurKey(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test04b_GetIndexSeq()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 4&
    
    Dim myResult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myS As SeqC
    Set myS = SeqC(10, 20, 30, 40, 50, 60, 70, 80, 90)
    Dim myI As IterItems
    Set myI = IterItems(myS)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myResult = myI.CurOffset(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test05a_GetItemArray()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 50

    Dim myArray As Variant
    myArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    ReDim Preserve myArray(-4 To 4)
    
    Dim myResult As Variant
    
    'Act:
    Dim myI As IterItems
    Set myI = IterItems(myArray)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myResult = myI.CurItem(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test05b_GetKeyArray()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 0&
    
    Dim myArray As Variant
    myArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    ReDim Preserve myArray(-4 To 4)
    
    Dim myResult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(myArray)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myResult = myI.CurKey(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test05c_GetIndexArray()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 4&
    
    Dim myArray As Variant
    myArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    ReDim Preserve myArray(-4 To 4)
    
    Dim myResult As Variant
    
    
    'Act:  Again we need to sort The result SeqC to get the matching array
    Dim myI As IterItems
    Set myI = IterItems(myArray)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myResult = myI.CurOffset(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test06a_GetItemCollection()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 50

    Dim myC As Collection
    Set myC = New Collection
    With myC
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    Dim myResult As Variant
    
    'Act:
    myResult = myI.CurItem(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test06b_GetKeyCollection()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 5&
    
    Dim myC As Collection
    Set myC = New Collection
    With myC
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myResult As Variant
    
    
    'Act:
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myResult = myI.CurKey(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test06c_GetIndexCollection()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 4&
    
    Dim myC As Collection
    Set myC = New Collection
    With myC
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    Dim myResult As Variant
    
    
    'Act:
    myResult = myI.CurOffset(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test07a_GetItemArrayList()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 50

    Dim myC As ArrayList
    Set myC = New ArrayList
    With myC
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    Dim myResult As Variant
    
    'Act:
    myResult = myI.CurItem(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test07b_GetKeyArrayList()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 4&
    
    Dim myC As ArrayList
    Set myC = New ArrayList
    With myC
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myResult As Variant
    
    
    'Act:
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myResult = myI.CurKey(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test07c_GetIndexArrayList()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 4&
    
    Dim myC As ArrayList
    Set myC = New ArrayList
    With myC
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
        .Add 70
        .Add 80
        .Add 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myC)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    Dim myResult As Variant
    
    
    'Act:
    myResult = myI.CurOffset(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test08a_GetItemDictionary()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 50

    Dim myK As KvpC
    Set myK = KvpC.Deb
    With myK
        .Add "Ten", 10
        .Add "Twenty", 20
        .Add "Thirty", 30
        .Add "Forty", 40
        .Add "Fifty", 50
        .Add "Sixty", 60
        .Add "Seventy", 70
        .Add "Eighty", 80
        .Add "Ninety", 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myK)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    Dim myResult As Variant
    
    'Act:
    myResult = myI.CurItem(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test08b_GetKeyDictionary()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = "Fifty"
    
    Dim myK As KvpC
    Set myK = KvpC.Deb
    With myK
        .Add "Ten", 10
        .Add "Twenty", 20
        .Add "Thirty", 30
        .Add "Forty", 40
        .Add "Fifty", 50
        .Add "Sixty", 60
        .Add "Seventy", 70
        .Add "Eighty", 80
        .Add "Ninety", 90
    End With
    
    Dim myResult As Variant
    
    
    'Act:
    Dim myI As IterItems
    Set myI = IterItems(myK)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    myResult = myI.CurKey(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IterItems")
Private Sub Test08c_GetIndexDIctionary()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = 4&
    
    Dim myK As KvpC
    Set myK = KvpC.Deb
    With myK
        .Add "Ten", 10
        .Add "Twenty", 20
        .Add "Thirty", 30
        .Add "Forty", 40
        .Add "Fifty", 50
        .Add "Sixty", 60
        .Add "Seventy", 70
        .Add "Eighty", 80
        .Add "Ninety", 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myK)
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    myI.MoveNext
    
    Dim myResult As Variant
    
    
    'Act:
    myResult = myI.CurOffset(0)
    
    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'

'@TestMethod("IterItems")
Private Sub Test09a_GetIndexDictionary()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50, 60, 70, 80, 90)
    
    
    Dim myK As KvpC
    Set myK = KvpC.Deb
    With myK
        .Add "Ten", 10
        .Add "Twenty", 20
        .Add "Thirty", 30
        .Add "Forty", 40
        .Add "Fifty", 50
        .Add "Sixty", 60
        .Add "Seventy", 70
        .Add "Eighty", 80
        .Add "Ninety", 90
    End With
    
    Dim myI As IterItems
    Set myI = IterItems(myK)
    
    
    
    Dim myResult As Variant
    ReDim myResult(0 To 8)
    
    'Act:
    Do
            
        myResult(myI.CurOffset(0)) = VBA.CVar(myI.CurItem(0))
        
    Loop While myI.MoveNext
    
    
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


