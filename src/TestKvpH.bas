Attribute VB_Name = "TestKvpH"
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

Public Sub KvpHTests()
 
    #If twinbasic Then
        Debug.Print CurrentProcedureName;
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName;
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
    
    Debug.Print vbTab, vbTab, vbTab, "Testing completed"

End Sub


'@TestMethod("VBALib.KvpH")
Private Sub Test01_ObjAndName()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb
    Dim myExpected As Variant
    myExpected = Array(True, "KvpH", "KvpH")
    
    Dim myResult(0 To 2) As Variant
    
    'Act:
    myResult(0) = VBA.IsObject(myK)
    myResult(1) = VBA.TypeName(myK)
    myResult(2) = myK.TypeName
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


'@TestMethod("VBALib.KvpH")
Private Sub Test02_Add_ThreeItems()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test03_Add_Pairs()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb
    
    Dim myItemsExpected As Variant
    myItemsExpected = Array(3, "Hello", True)
    ReDim Preserve myItemsExpected(1 To 3)
    
    Dim myKeysExpected As Variant
    myKeysExpected = Array(1, 2, 3)
    ReDim Preserve myKeysExpected(1 To 3)
    
    
    Dim myItemsResult As Variant
    Dim myKeysResult As Variant
    
    'Act:
    myK.AddPairs SeqA(1, 2, 3), SeqA(3, "Hello", True)
   
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


'@TestMethod("VBALib.KvpH")
Private Sub Test04a_GetItem()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&), SeqA(3, "Hello", True))
   
    Dim myExpected As String
    myExpected = "Hello"
    
    Dim myResult As String
    
    'Act:
    myResult = myK.Item(2&)
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test04b_LetItem()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&), SeqA(3, "Hello", True))
   
    Dim myExpected As String
    myExpected = "World"
    
    Dim myResult As String
    
    'Act:
    myK.Item(2) = "World"
    myResult = myK.Item(2&)
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test04c_SetItem()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&), SeqA(3, "Hello", True))
   
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&)
    ReDim Preserve myExpected(1 To 3)
    
    Dim myResult As Variant
    
    'Act:
    Set myK.Item(2) = SeqA(1&, 2&, 3&)
    myResult = myK.Item(2&).ToArray
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test05a_Remove()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, True, 1&, 2&, 3&, 4&)
    ReDim Preserve myExpected(1 To 6)
    
    Dim myResult As Variant
    
    'Act:
    myK.Remove 2&
    myResult = myK.Items
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test05b_Remove()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, True, 2&, 4&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myResult As Variant
    
    'Act:
    myK.Remove 2&, 4&, 6&
    myResult = myK.Items
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test06_RemoveAfter()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, "Hello", 3&, 4&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myResult As Variant
    
    'Act:
    myResult = myK.RemoveAfter(2&, 3).Items
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test07_RemoveBefore()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, "Hello", 3&, 4&)
    ReDim Preserve myExpected(1 To 4)
    
    Dim myResult As Variant
    
    'Act:
    myResult = myK.RemoveBefore(6&, 3).Items
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test08a_RemoveAll()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Long
    myExpected = -1
    
    Dim myResult As Variant
    
    'Act:
    myResult = myK.RemoveAll.Count
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test08b_Clear()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Long
    myExpected = -1
    
    Dim myResult As Variant
    
    'Act:
    myResult = myK.Clear.Count
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test08c_Reset()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Long
    myExpected = -1
    
    Dim myResult As Variant
    
    'Act:
    myResult = myK.Clear.Count
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test09_Clone()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb
    
    Dim myItemsExpected As Variant
    myItemsExpected = SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&).ToArray
    
    Dim myKeysExpected As Variant
    myKeysExpected = SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&).ToArray
    
    Dim myItemsResult As Variant
    Dim myKeysResult As Variant
    
    'Act:
    Dim myT As KvpH
    
    Set myT = myK.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&)).Clone
   
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


'@TestMethod("VBALib.KvpH")
Private Sub Test10_Hold_Lacks_FilledSeq()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(True, False, True, True, False, False, False, False, True, True, True, True, False, False, False, False, True, True)
    ReDim Preserve myExpected(1 To 18)
    
    Dim myResult As Variant
    ReDim myResult(1 To 18)
    'Act:
    myResult(1) = myK.HoldsItems
    myResult(2) = myK.LacksItems
    
    myResult(3) = myK.HoldsItem("Hello")
    myResult(4) = myK.HoldsItem(4&)
    myResult(5) = myK.HoldsItem(42&)
    myResult(6) = myK.HoldsItem("World")
    
    myResult(7) = myK.LacksItem("Hello")
    myResult(8) = myK.LacksItem(4&)
    myResult(9) = myK.LacksItem(42&)
    myResult(10) = myK.LacksItem("World")
    
    myResult(11) = myK.HoldsKey(2&)
    myResult(12) = myK.HoldsKey(6&)
    myResult(13) = myK.HoldsKey(42&)
    myResult(14) = myK.HoldsKey("Hello")
    
    myResult(15) = myK.LacksKey(2&)
    myResult(16) = myK.LacksKey(6&)
    myResult(17) = myK.LacksKey(42&)
    myResult(18) = myK.LacksKey("Hello")
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test11_MappedIt()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(4&, "Hellp", True, 2&, 3&, 4&, 5&)
    ReDim Preserve myExpected(1 To 7)
    
    Dim myResult As Variant
    
    'Act:
    myK.MappedIt mpInc.Deb
    
    myResult = myK.Items
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


'@TestMethod("VBALib.KvpH")
Private Sub Test12_MapIt()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myOrigExpected As Variant
    myOrigExpected = Array(3&, "Hello", True, 1&, 2&, 3&, 4&)
    ReDim Preserve myOrigExpected(1 To 7)
    
    Dim myMapExpected As Variant
    myMapExpected = Array(4&, "Hellp", True, 2&, 3&, 4&, 5&)
    ReDim Preserve myMapExpected(1 To 7)
    
    Dim myOrigResult As Variant
    Dim myMapresult As KvpH
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test13_FilterIt()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = Array(3&, 3&, 4&)
    ReDim Preserve myExpected(1 To 3)
    
    Dim myResult As Variant
    
    'Act:
    myResult = myK.FilterIt(cmpMT(2)).Items
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test14_ReduceIt()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(1&, 2&, 3&, 4&, 5&, 6&, 7&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As LongLong
    myExpected = VBA.CLngLng(3 + 1 + 2 + 3 + 4)
    
    Dim myResult As LongLong
    
    'Act:
    myResult = myK.ReduceIt(rdSum.Deb)
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test15a_KeyByIndex()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 30&
    
    Dim myResult As Variant
    
    'Act:
    myResult = myK.KeyByIndex(3)
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test15b_KeyOf()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 20&
    
    Dim myResult As Variant
    
    'Act:
    myResult = myK.KeyOf("Hello")
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test16a_GetFirst()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 3&
    
    Dim myResult As Variant
    
    'Act:
    myResult = myK.First
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test16b_LetFirst()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 42&
    
    Dim myResult As Variant
    
    'Act:
    myK.First = 42&
    myResult = myK.First
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test16c_GetLast()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 4&
    
    Dim myResult As Variant
    
    'Act:
    myResult = myK.Last
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test16d_LetLast()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 42&
    
    Dim myResult As Variant
    
    'Act:
    myK.Last = 42&
    myResult = myK.Last
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test16e_GetFirstKey()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 10&
    
    Dim myResult As Variant
    
    'Act:
    myResult = myK.FirstKey
    
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


'@TestMethod("VBALib.KvpH")
Private Sub Test16f_LastKey()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myK As KvpH
    Set myK = KvpH.Deb.AddPairs(SeqA(10&, 20&, 30&, 40&, 50&, 60&, 70&), SeqA(3&, "Hello", True, 1&, 2&, 3&, 4&))
   
    Dim myExpected As Variant
    myExpected = 70&
    
    Dim myResult As Variant
    
    'Act:
    myResult = myK.LastKey
    
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
