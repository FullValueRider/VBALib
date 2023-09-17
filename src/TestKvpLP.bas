Attribute VB_Name = "TestKvpLP"
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


#End If


Public Sub KvpLPTests()
 
    #If TWINBASIC Then
        Debug.Print CurrentProcedureName,
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName,
    #End If

    Test01_ObjAndName
    Test02_Add_ThreeItems
    Test03_Add_Pairs
    Test04a_GetItem
    Test04b_LetItem
    Test04c_SetItem
    Test05a_Remove
    Test05b_Remove
    Test06_RemoveAfter
    Test07_RemoveBefore
    Test08a_RemoveAll
    Test08b_Clear
    Test08c_Reset
    Test09_Clone
    Test10_Hold_Lacks_FilledSeq
    Test11_MappedIt
    Test12_MapIt
    Test13_FilterIt
    Test14_ReduceIt
    Test15a_KeyByIndex
    Test15b_KeyOf
    Test16a_GetFirst
    Test16b_LetFirst
    Test16c_GetLast
    Test16d_LetLast
    Test16e_GetFirstKey
    Test16f_LastKey
        
    
    Debug.Print vbTab, vbTab, "Testing completed"

End Sub


'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test01_ObjAndName()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb
    Dim myExpected As Variant
    myExpected = Array(True, "KvpLP", "KvpLP")
    
    Dim myresult(0 To 2) As Variant
    
    'Act:
    myresult(0) = VBA.IsObject(myK)
    myresult(1) = VBA.TypeName(myK)
    myresult(2) = myK.TypeName
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test02_Add_ThreeItems()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb
    
    Dim myItemsExpected As Variant
    myItemsExpected = Array(3, "Hello", True)
    ReDim Preserve myItemsExpected(1 To 3)
    
    Dim myKeysExpected As Variant
    myKeysExpected = Array(1, 2, 3)
    ReDim Preserve myKeysExpected(1 To 3)
    
    
    Dim myItemsResult As Variant
    Dim myKeysResult As Variant
    
    'Act:
    myK.Add 1, 3
    myK.Add 2, "Hello"
    myK.Add 3, True
    
    myItemsResult = myK.Items
    myKeysResult = myK.Keys
    'Assert:
    AssertExactSequenceEquals myItemsExpected, myItemsResult, myProcedureName
    AssertExactSequenceEquals myKeysExpected, myKeysResult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test03_Add_Pairs()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb
    
    Dim myItemsExpected As Variant
    myItemsExpected = Array(3, "Hello", True)
    ReDim Preserve myItemsExpected(1 To 3)
    
    Dim myKeysExpected As Variant
    myKeysExpected = Array(1, 2, 3)
    ReDim Preserve myKeysExpected(1 To 3)
    
    
    Dim myItemsResult As Variant
    Dim myKeysResult As Variant
    
    'Act:
    myK.AddPairs SeqL(1, 2, 3), SeqL(3, "Hello", True)
   
    myItemsResult = myK.Items
    myKeysResult = myK.Keys
    'Assert:
    AssertExactSequenceEquals myItemsExpected, myItemsResult, myProcedureName
    AssertExactSequenceEquals myKeysExpected, myKeysResult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test04a_GetItem()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&), SeqL(3, "Hello", True))
   
    Dim myExpected As String
    myExpected = "Hello"
    
    Dim myresult As String
    
    'Act:
    myresult = myK.Item(2&)
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test04b_LetItem()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&), SeqL(3, "Hello", True))
   
    Dim myExpected As String
    myExpected = "World"
    
    Dim myresult As String
    
    'Act:
    myK.Item(2) = "World"
    myresult = myK.Item(2&)
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test04c_SetItem()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&), SeqL(3, "Hello", True))
   
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&)
    ReDim Preserve myExpected(1 To 3)
    
    Dim myresult As Variant
    
    'Act:
    Set myK.Item(2) = SeqL(1&, 2&, 3&)
    myresult = myK.Item(2&).ToArray
    
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test05a_Remove()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, True, 1&, 2&, 3&, 4&)
    ReDim Preserve myExpected(1 To 6)
    
    Dim myresult As Variant
    
    'Act:
    myK.Remove 2&
    myresult = myK.Items
    
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test05b_Remove()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, True, 2&, 4&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant
    
    'Act:
    myK.Remove 2&, 4&, 6&
    myresult = myK.Items
    
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test06_RemoveAfter()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, "Hello", 3&, 4&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.RemoveAfter(2&, 3).Items
    
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test07_RemoveBefore()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, "Hello", 3&, 4&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.RemoveBefore(6&, 3).Items
    
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test08a_RemoveAll()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Long
    myExpected = -1
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.RemoveAll.Count
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test08b_Clear()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Long
    myExpected = -1
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.Clear.Count
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test08c_Reset()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Long
    myExpected = -1
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.Clear.Count
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test09_Clone()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb
    
    Dim myItemsExpected As Variant
    myItemsExpected = SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&).ToArray
    
    Dim myKeysExpected As Variant
    myKeysExpected = SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&).ToArray
    
    Dim myItemsResult As Variant
    Dim myKeysResult As Variant
    
    'Act:
    Dim myT As KvpLP
    
    Set myT = myK.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&)).Clone
   
    myItemsResult = myT.Items
    myKeysResult = myT.Keys
    
    'Assert:
    AssertExactSequenceEquals myItemsExpected, myItemsResult, myProcedureName
    AssertExactSequenceEquals myKeysExpected, myKeysResult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test10_Hold_Lacks_FilledSeq()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(True, False, True, True, False, False, False, False, True, True, True, True, False, False, False, False, True, True)
    ReDim Preserve myExpected(1 To 18)
    
    Dim myresult As Variant
    ReDim myresult(1 To 18)
    'Act:
    myresult(1) = myK.HoldsItems                '
    myresult(2) = myK.LacksItems
    
    myresult(3) = myK.HoldsItem("Hello")
    myresult(4) = myK.HoldsItem(4&)
    myresult(5) = myK.HoldsItem(42&)
    myresult(6) = myK.HoldsItem("World")
    
    myresult(7) = myK.LacksItem("Hello")
    myresult(8) = myK.LacksItem(4&)
    myresult(9) = myK.LacksItem(42&)
    myresult(10) = myK.LacksItem("World")
    
    myresult(11) = myK.HoldsKey(2&)
    myresult(12) = myK.HoldsKey(6&)
    myresult(13) = myK.HoldsKey(42&)
    myresult(14) = myK.HoldsKey("Hello")
    
    myresult(15) = myK.LacksKey(2&)
    myresult(16) = myK.LacksKey(6&)
    myresult(17) = myK.LacksKey(42&)
    myresult(18) = myK.LacksKey("Hello")
    
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test11_MappedIt()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(4&, "Hellp", True, 2&, 3&, 4&, 5&)
    ReDim Preserve myExpected(1 To 7)
    
    Dim myresult As Variant
    
    'Act:
    Set myresult = myK.MappedIt(mpInc.Deb)
    myresult = myresult.Items
    
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
   
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test12_MapIt()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myOrigExpected As Variant
    myOrigExpected = Array(3&, "Hello", True, 1&, 2&, 3&, 4&)
    ReDim Preserve myOrigExpected(1 To 7)
    
    Dim myMapExpected As Variant
    myMapExpected = Array(4&, "Hellp", True, 2&, 3&, 4&, 5&)
    ReDim Preserve myMapExpected(1 To 7)
    
    Dim myOrigResult As Variant
    Dim myMapresult As KvpLP
    
    'Act:
    myOrigResult = myK.Items
    Set myMapresult = myK.MapIt(mpInc.Deb)
    
    'Assert:
    AssertExactSequenceEquals myOrigExpected, myOrigResult, myProcedureName
    AssertExactSequenceEquals myMapExpected, myMapresult.Items, myProcedureName
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test13_FilterIt()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, 3&, 4&)
    ReDim Preserve myExpected(1 To 3)
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.FilterIt(cmpMT(2)).Items
    
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test14_ReduceIt()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As LongLong
    myExpected = VBA.CLngLng(1 + 2 + 3 + 4 + 3)
    
    Dim myresult As LongLong
    
    'Act:
    myresult = myK.ReduceIt(rdSum)
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test15a_KeyByIndex()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 30&
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.KeyByIndex(3)
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test15b_KeyOf()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 20&
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.KeyOf("Hello")
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test16a_GetFirst()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 3&
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.First
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test16b_LetFirst()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 42&
    
    Dim myresult As Variant
    
    'Act:
    myK.First = 42&
    myresult = myK.First
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test16c_GetLast()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 4&
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.Last
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test16d_LetLast()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 42&
    
    Dim myresult As Variant
    
    'Act:
    myK.Last = 42&
    myresult = myK.Last
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test16e_GetFirstKey()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 10&
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.FirstKey
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("VBALib.KvpLP")
Private Sub Test16f_LastKey()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpLP
    Set myK = KvpLP.Deb.AddPairs(SeqL(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqL(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 70&
    
    Dim myresult As Variant
    
    'Act:
    myresult = myK.LastKey
    
    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

