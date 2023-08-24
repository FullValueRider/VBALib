Attribute VB_Name = "TestSeqH"
'@IgnoreModule
'@TestModule
'@Folder("Tests")
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
    ' Set Assert = CreateObject("Rubberduck.AssertClass")
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

Public Sub SeqHTests()
 
    #If twinbasic Then
        Debug.Print CurrentProcedureName;
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName;
    #End If

    Test01_SeqObj
    
    Test02a_InitByLong_10FirstIndex_LastIndex
    Test02b_InitByString
    Test02c_InitByForEachArray
    Test02d_InitByForEachArrayList
    Test02e_InitByForEachCollection
    Test02f_InitByDictionary
    
    Test03a_WriteItem
    
    Test04a_Add_MultipleItems
    
    Test06a_AddRange_String
    Test06b_AddRange_Array
    Test06c_AddRange_Collection
    Test06d_AddRange_ArrayList
    Test06e_AddRange_Dictionary
    
    Test07a_Insert_SingleItems
    Test07b_Insert_MultipleItems
    
    Test08a_InsertAtRange_String
    Test08b_InsertAtRange_Array
    Test08c_InsertAtRange_Collection
    Test08d_InsertAtRange_ArrayList
    Test08e_InsertAtRange_Dictionary
    
    Test09a0_Remove_SingleItem
    
    Test09a_RemoveAt_SingleItem
    Test09b_RemoveAt_ThreeItems
    
    Test10a_Remove_SingleItems
    
    Test11a_RemoveRange_SingleItem
    Test11b_RemoveRange_ThreeItems
    
    Test12a_RemoveIndexesRange_ThreeItems
    
    Test13a_RemoveAll_DefaultAll
    Test13b_RemoveAll_Default_42AndHello
    Test13c_Reset
    Test13d_Clear
    
    Test14a_Fill
    
    Test15a_Slice
    Test15b_SliceToEnd
    Test15c_SliceRunOnly
    Test15d_Slice_Start3_End9_step2
    Test15e_Slice_Start3_End9_step2_ToCollection
    Test15f_Slice_Start3_End9_step2_ToArray
    
    Test16a_Head
    Test16b_Head_3Items
    Test16c_HeadZeroItems
    Test16d_HeadFullSeq
    
    Test17a_Tail
    Test17b_Tail_3Items
    Test17c_TailFullItems
    Test17d_TailZeroSeq
    
    Test18a_KnownIndexes_Available
    Test18b_KnownIndexes_Unavailable
    
    Test19a_KnownValues_Available
    
    Test20a_IndexOf_WholeSeq_Present
    Test20b_IndexOf_WholeSeq_NotPresent
    Test20c_IndexOf_SubSeq_Present
    Test20d_IndexOf_SubSeq_NotPresent
    
    Test21a_LastIndexOf_WholeSeq_Present
    Test21b_LastIndexOf_WholeSeq_NotPresent
    Test21c_LastIndexOf_SubSeq_Present
    Test21d_LastIndexOf_SubSeq_NotPresent
    
    Test22a_Push
    Test22b_PushRange
    
    Test23a_Pop
    Test23b_PopRange
    Test23c_PopRange_ExceedsHost
    Test23d_PopRange_NegativeRun
    
    Test24a_Enqueue
    Test24b_EnqueueRange
    
    Test25a_Dequeue
    Test25b_DeqeueRange
    Test25c_DequeueRange_ExceedsHost
    
    Test26a_Sort
    Test26b_Sorted
    
    Test27a_Reverse
    Test27b_Reversed
    
    Test28a_Unique
    Test28b_Unique_SingleItem
    Test28c_Unique_NoItems
    
    Test29a_SetOfCommon
    Test29b_SetOfHostOnly
    Test29c_SetOfParamOnly
    Test29d_SetOfNotCommon
    Test29e_SetOfUnique
    
    Test30a_Swap
    
    Debug.Print vbTab, vbTab, vbTab, "Testing completed"

End Sub


'@TestMethod("SeqH")
Private Sub Test01_SeqObj()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Set mySeq = SeqH.Deb
    Dim myExpected As Variant
    myExpected = Array(True, "SeqH", "SeqH")
    
    Dim myResult(0 To 2) As Variant
    
    'Act:
    myResult(0) = VBA.IsObject(mySeq)
    myResult(1) = VBA.TypeName(mySeq)
    myResult(2) = mySeq.TypeName
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


'@TestMethod("SeqH")
Private Sub Test02a_InitByLong_10FirstIndex_LastIndex()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Set mySeq = SeqH(1000)
    Dim myExpected As Variant
    myExpected = -1&
    ' SeqH is hash based.  the initialisation number sets the size of the hashslots required.
    ' It does not preallocate items containing empty as this would just load a single hashslot
    
    Dim myResult As Variant
    
    'Act:
    myResult = mySeq.ToArray
    'Assert:
    AssertExactAreEqual -1&, mySeq.FirstIndex, myProcedureName
    AssertExactAreEqual -1&, mySeq.LastIndex, myProcedureName
    AssertExactAreEqual -1&, mySeq.Count, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqH")
Private Sub Test02b_InitByString()
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
    myExpected = Array("H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 5)
    
    Dim myResult As Variant
    
    Dim mySeq As SeqH
    
    'Act:
    Set mySeq = SeqH("Hello")
    
    myResult = mySeq.ToArray
    
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


'@TestMethod("SeqH")
Private Sub Test02c_InitByForEachArray()
    '    #If twinbasic Then
    '        myProcedureName = myComponentName & ":" & CurrentProcedureName
    '       myComponentName = CurrentComponentName
    '    #Else
    '        myProcedureName = errex.livecallstack.Modulename & ":" & ErrEx.LiveCallstack.ProcedureName
    '        myComponentName = ErrEx.LiveCallstack.ModuleName
'    #End If
    '      on error GoTo TestFail
    '
    '    'Arrange:
    '    Dim myArray(1 To 3, 1 To 3) As Variant
    '    Dim myCount As Long
    '    myCount = 1
    '    Dim myFirst As Long
    '    For myFirst = 1 To 3
    '
    '        Dim mySecond As Long
    '        For mySecond = 1 To 3
    '            myArray(myFirst, mySecond) = myCount
    '            myCount = myCount + 1
    '        Next
    '    Next
    '
    '    Dim myExpected As Variant
    '    myExpected = Array(1&, 4&, 7&, 2&, 5&, 8&, 3&, 6&, 9&)
    '    ReDim Preserve myExpected(1 To 9)
    '
    '    Dim myResult As Variant
    '
    '    Dim mySeq As SeqH
    '
    '    'Act:
    '    ' because we want to add an array rather than a forwarded
    '    ' param array we must encapsule the array in an array
    '    Set mySeq = SeqH(myArray)
    '
    '    myResult = mySeq.ToArray
    '
    '    'Assert:
    '    AssertExactSequenceEquals myExpected, myResult, myProcedureName
    'TestExit:
    '    '@Ignore UnhandledOnErrorResumeNext
    '    on error Resume Next
    '
    '    Exit Sub
    'TestFail:
    '    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    '    Resume TestExit
End Sub


'@TestMethod("SeqH")
Private Sub Test02d_InitByForEachArrayList()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myAL As Object
    Set myAL = CreateObject("System.Collections.Arraylist")
    
    With myAL
        .Add 1
        .Add 4
        .Add 7
        .Add 2
        .Add 5
        .Add 8
        .Add 3
        .Add 6
        .Add 9
    End With
    
    Dim myExpected As Variant
    myExpected = Array(1, 4, 7, 2, 5, 8, 3, 6, 9)
    ReDim Preserve myExpected(1 To 9)
    
    Dim myResult As Variant
    
    Dim mySeq As SeqH
    
    'Act:
    Set mySeq = SeqH(myAL)
    
    myResult = mySeq.ToArray
    
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


'@TestMethod("SeqH")
Private Sub Test02e_InitByForEachCollection()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myC As Collection
    Set myC = New Collection
    
    With myC
        .Add 1
        .Add 4
        .Add 7
        .Add 2
        .Add 5
        .Add 8
        .Add 3
        .Add 6
        .Add 9
    End With
    
    Dim myExpected As Variant
    myExpected = Array(1, 4, 7, 2, 5, 8, 3, 6, 9)
    ReDim Preserve myExpected(1 To 9)
    
    Dim myResult As Variant
    
    Dim mySeq As SeqH
    
    'Act:
    Set mySeq = SeqH(myC)
    
    myResult = mySeq.ToArray
    
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


'@TestMethod("SeqH")
Private Sub Test02f_InitByDictionary()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim myD As Scripting.Dictionary
    Set myD = New Scripting.Dictionary
    
    With myD
        .Add "Hello", "World"
        .Add "Ten", 10&
        .Add "Thing", 3.142
        
    End With
    
    Dim myExpected As Variant
    myExpected = Array("Hello", "World", "Ten", 10&, "Thing", 3.142)
    ReDim Preserve myExpected(1 To 6)
    
    Dim myResult As Variant
    ReDim myResult(1 To 6)
    
    Dim mySeq As SeqH
    
    'Act:
    Set mySeq = SeqH(myD)
    Dim myTmp As Variant
    myTmp = mySeq.ToArray
    
    myResult(1) = myTmp(1)(0)
    myResult(2) = myTmp(1)(1)
    myResult(3) = myTmp(2)(0)
    myResult(4) = myTmp(2)(1)
    myResult(5) = myTmp(3)(0)
    myResult(6) = myTmp(3)(1)
        
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


'@TestMethod("SeqH")
Private Sub Test03a_WriteItem()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Set mySeq = SeqH.Deb
    mySeq.Add 1
    mySeq.Add 2
    mySeq.Add 3
    
    mySeq.Item(1) = 42
    mySeq.Item(2) = "Hello"
    mySeq.Item(3) = 3.142
    Dim myExpected As Variant
    myExpected = Array(True, True, True)
    
    Dim myResult As Variant
    ReDim myResult(0 To 2)
    'Act:
    myResult(0) = mySeq.Item(1) = 42
    myResult(1) = mySeq.Item(2) = "Hello"
    myResult(2) = mySeq.Item(3) = "3.142"
   
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


'@TestMethod("SeqH")
Private Sub Test04a_Add_MultipleItems()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, 42, "Hello", 3.142)
    ReDim Preserve myExpected(1 To 8)
      
    Dim myResult As Variant
   
    'Act:
    Set mySeq = SeqH.Deb
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    myResult = mySeq.AddItems(42, "Hello", 3.142).ToArray
    
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


'@TestMethod("SeqH")
Private Sub Test06a_AddRange_String()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 10)
   
    Dim myResult As Variant
   
    'Act:
    Set mySeq = SeqH.Deb
        mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty

    myResult = mySeq.AddRange("Hello").ToArray
    
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


'@TestMethod("SeqH")
Private Sub Test06b_AddRange_Array()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 10)
   
    Dim myResult As Variant
   
    'Act:
    Set mySeq = SeqH.Deb
        mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty

    myResult = mySeq.AddRange(Array("H", "e", "l", "l", "o")).ToArray
    
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


'@TestMethod("SeqH")
Private Sub Test06c_AddRange_Collection()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 10)
      
    Dim myResult As Variant
   
    Dim myC As Collection
    Set myC = New Collection
    With myC
        .Add "H"
        .Add "e"
        .Add "l"
        .Add "l"
        .Add "o"
    End With
    
    'Act:
    Set mySeq = SeqH.Deb
        mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty

    myResult = mySeq.AddRange(myC).ToArray
    
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


'@TestMethod("SeqH")
Private Sub Test06d_AddRange_ArrayList()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 10)
   
    Dim myResult As Variant
   
    Dim myAL As Object
    Set myAL = CreateObject("System.Collections.Arraylist")
    With myAL
        .Add "H"
        .Add "e"
        .Add "l"
        .Add "l"
        .Add "o"
    End With
    
    'Act:
    Set mySeq = SeqH.Deb
        mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty

    myResult = mySeq.AddRange(myAL).ToArray
    
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


'@TestMethod("SeqH")
Private Sub Test06e_AddRange_Dictionary()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "Hello1", "There2", "World3")
    ReDim Preserve myExpected(1 To 8)
   
    Dim myResult As Variant
   
    Dim myD As KvpH
    Set myD = KvpH.Deb
    With myD
        .Add "Hello", 1
        .Add "There", 2
        .Add "World", 3
    End With
'    Debug.Print
'    myD.PrintByHash
'    Debug.Print
'    myD.PrintByOrder
    
    'Act:
    Set mySeq = SeqH.Deb.Fill(Empty, 5)
    
    mySeq.AddRange myD
    
'    Debug.Print
'    myD.PrintByHash
'    Debug.Print
'    myD.PrintByOrder
    
    myResult = mySeq.ToArray
    myResult(6) = myResult(6)(0) & VBA.CStr(myResult(6)(1))
    myResult(7) = myResult(7)(0) & VBA.CStr(myResult(7)(1))
    myResult(8) = myResult(8)(0) & VBA.CStr(myResult(8)(1))
    
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


'@TestMethod("SeqH")
Private Sub Test07a_Insert_SingleItems()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10, 20, "Hello", 30, 42&, 40, 3.142, 50)
    ReDim Preserve myExpected(1 To 8)
    Dim myExpected2 As Variant
    myExpected2 = Array(3&, 5&, 7&)
      
    Dim myResult As Variant
    Dim myResult2 As Variant
    ReDim myResult2(0 To 2)
    
    'Act:
    Set mySeq = SeqH(10, 20, 30, 40, 50)
    myResult2(0) = mySeq.InsertAt(3, "Hello")
    myResult2(1) = mySeq.InsertAt(5, 42&)
    myResult2(2) = mySeq.InsertAt(7, 3.142)
    
    myResult = mySeq.ToArray
    'Assert:
    AssertExactSequenceEquals myExpected, myResult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName
   
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqH")
Private Sub Test07b_Insert_MultipleItems()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)
   
    Dim myResult As Variant
   
    'Act:
    Set mySeq = SeqH.Deb
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.InsertAtItems 3, "Hello", 42&, 3.142

    myResult = mySeq.ToArray
    
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


''@TestMethod("SeqH")
'Private Sub Test09a_InsertItems()
'    #If twinbasic Then
'        myProcedureName = myComponentName & ":" & CurrentProcedureName
'       myComponentName = CurrentComponentName
'    #Else
'        myProcedureName = errex.livecallstack.Modulename & ":" & ErrEx.LiveCallstack.ProcedureName
'        myComponentName = ErrEx.LiveCallstack.ModuleName
'    #End If
'      on error GoTo TestFail
'
'    'Arrange:
'    Dim mySeq As SeqH
'    Dim myExpected As Variant
'    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
'    ReDim Preserve myExpected(1 To 8)
'    Dim myExpected2 As Variant
'    myExpected2 = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
'    ReDim Preserve myExpected2(1 To 8)
'    Dim myResult As Variant
'    Dim myResult2 As Variant
'
'
'    'Act:
'
'    Set mySeq = SeqH
'    myResult2 = mySeq.InsertItems(3, "Hello", 42&, 3.142).ToArray
'
'
'    myResult = mySeq.ToArray
'    'Assert:
'    AssertExactSequenceEquals myExpected, myResult,myProcedureName
'    AssertExactSequenceEquals myExpected2, myResult2
'
'TestExit:
'    '@Ignore UnhandledOnErrorResumeNext
'    on error Resume Next
'
'    Exit Sub
'TestFail:
'    AssertFail myCOmponentName, myProcedureName," raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub


'@TestMethod("SeqH")
Private Sub Test08a_InsertAtRange_String()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "H", "e", "l", "l", "o", Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 10)
    
    Dim myResult As Variant
  
    'Act:
    Set mySeq = SeqH.Deb
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.InsertAtRange 3, "Hello"
   
    myResult = mySeq.ToArray
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


'@TestMethod("SeqH")
Private Sub Test08b_InsertAtRange_Array()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)
   
    Dim myResult As Variant
    
    'Act:
    Set mySeq = SeqH.Deb
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.InsertAtRange 3, Array("Hello", 42&, 3.142)
    
    myResult = mySeq.ToArray
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


'@TestMethod("SeqH")
Private Sub Test08c_InsertAtRange_Collection()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)
    
    Dim myResult As Variant
   
    
    Dim myC As Collection
    Set myC = New Collection
    With myC
        .Add "Hello"
        .Add 42&
        .Add 3.142
    
    End With
    
    'Act:
    Set mySeq = SeqH.Deb
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.InsertAtRange 3, myC
    
    myResult = mySeq.ToArray
    
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


'@TestMethod("SeqH")
Private Sub Test08d_InsertAtRange_ArrayList()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)
    
    Dim myResult As Variant
    
    
    Dim myAL As Object
    Set myAL = CreateObject("System.collections.arraylist")
    With myAL
        .Add "Hello"
        .Add 42&
        .Add 3.142
    
    End With
    
    'Act:
    Set mySeq = SeqH.Deb
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.Add Empty
    mySeq.InsertAtRange 3, myAL
   
    myResult = mySeq.ToArray
    
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


'@TestMethod("SeqH")
Private Sub Test08e_InsertAtRange_Dictionary()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10, 20, "Hello1", "There2", "World3", 30, 40, 50)
    ReDim Preserve myExpected(1 To 8)
    
    Dim myResult As Variant
    
    Dim myD As KvpA
    Set myD = KvpA.Deb
    With myD
        .Add "Hello", 1
        .Add "There", 2
        .Add "World", 3
    
    End With
    
    'Act:
    Set mySeq = SeqH(10, 20, 30, 40, 50)
    mySeq.InsertAtRange 3, myD
    myResult = mySeq.ToArray
   
    myResult(3) = myResult(3)(0) & VBA.CStr(myResult(3)(1))
    myResult(4) = myResult(4)(0) & VBA.CStr(myResult(4)(1))
    myResult(5) = myResult(5)(0) & VBA.CStr(myResult(5)(1))
    
    
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


'@TestMethod("SeqH")
Private Sub Test09a0_Remove_SingleItem()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)

    Dim myResult As Variant
    Set mySeq = SeqH(Empty, Empty, Empty, 42, Empty, Empty)

    'Act:
    mySeq.Remove 42

    myResult = mySeq.ToArray

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


'@TestMethod("SeqH")
Private Sub Test09a_RemoveAt_SingleItem()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)

    Dim myResult As Variant
    Set mySeq = SeqH(Empty, Empty, Empty, 42, Empty, Empty)

    'Act:
    mySeq.RemoveAt 4

    myResult = mySeq.ToArray

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


'
'@TestMethod("SeqH")
Private Sub Test09b_RemoveAt_ThreeItems()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10, 20, 30, 40, 50)
    ReDim Preserve myExpected(1 To 5)

    Dim myResult As Variant
    Set mySeq = SeqH(10, 42, 20, 30, 42, 40, 50, 42)

    'Act:
    mySeq.RemoveIndexes 8, 2, 5

    myResult = mySeq.ToArray

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


'@TestMethod("SeqH")
Private Sub Test10a_Remove_SingleItems()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, 42, "Hello", "Hello", Empty, Empty, 42, Empty, Empty)
    ReDim Preserve myExpected(1 To 11)

    Dim myResult As Variant
    Set mySeq = SeqH(Empty, 42, Empty, Empty, 42, "Hello", "Hello", "Hello", Empty, 3.142, Empty, 42, Empty, Empty)

    'Act:
    mySeq.RemoveRange SeqH(42, 3.142, "Hello")

    myResult = mySeq.ToArray

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


'@TestMethod("SeqH")
Private Sub Test11a_RemoveRange_SingleItem()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)

    Dim myResult As Variant
    Set mySeq = SeqH(Empty, Empty, Empty, 42, Empty, Empty)

    'Act:
    mySeq.RemoveRange SeqH.Deb.AddItems(42)

    myResult = mySeq.ToArray

    'Assert:
    AssertExactSequenceEquals myExpected, myResult, myProcedureName

TestExit:
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqH")
Private Sub Test11b_RemoveRange_ThreeItems()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)

    Dim myResult As Variant
    Set mySeq = SeqH(Empty, Empty, Empty, 42, 42, 42, Empty, Empty)

    'Act:
    mySeq.RemoveRange SeqH(42, 42, 42)

    myResult = mySeq.ToArray

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


'
'@TestMethod("SeqH")
Private Sub Test12a_RemoveIndexesRange_ThreeItems()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)

    Dim myResult As Variant
    Set mySeq = SeqH(Empty, Empty, Empty, 42, 42, 42, Empty, Empty)

    'Act:
    mySeq.RemoveIndexesRange SeqH(4, 5, 6)

    myResult = mySeq.ToArray

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


'@TestMethod("SeqH")
Private Sub Test13a_RemoveAll_DefaultAll()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = -1&

    Dim myResult As Variant
    Set mySeq = SeqH(Empty, Empty, Empty, 42, 42, 42, Empty, Empty)

    'Act:
    mySeq.RemoveAll
    myResult = mySeq.Count

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


'@TestMethod("SeqH")
Private Sub Test13b_RemoveAll_Default_42AndHello()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim myExpected(1 To 5)

    Dim myResult As Variant
    Set mySeq = SeqH(Empty, "Hello", Empty, "Hello", "Hello", Empty, 42, 42, 42, Empty, Empty)

    'Act:
    mySeq.RemoveAll "Hello", 42
    myResult = mySeq.ToArray

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


'@TestMethod("SeqH")
Private Sub Test13c_Reset()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = -1&

    Dim myResult As Variant
    Set mySeq = SeqH(Empty, Empty, Empty, 42, 42, 42, Empty, Empty)

    'Act:
    mySeq.Reset
    myResult = mySeq.Count

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


'@TestMethod("SeqH")
Private Sub Test13d_Clear()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = -1&

    Dim myResult As Variant
    Set mySeq = SeqH(Empty, Empty, Empty, 42, 42, 42, Empty, Empty)

    'Act:
    mySeq.Clear
    myResult = mySeq.Count

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


'@TestMethod("SeqH")
Private Sub Test14a_Fill()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(True, True, True)
    ReDim Preserve myExpected(1 To 3)

    Dim myResult As Variant
    ReDim myResult(1 To 3)
    Set mySeq = SeqH(Empty, Empty, Empty)

    'Act:
    mySeq.Fill 42, 10
    myResult(1) = mySeq.Count = 13
    myResult(2) = mySeq.Item(4) = 42&
    myResult(3) = mySeq.Item(13) = 42&

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


'@TestMethod("SeqH")
Private Sub Test15a_Slice()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(3&, 4&, 5&)
    ReDim Preserve myExpected(1 To 3)

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Slice(3, 3).ToArray

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


'@TestMethod("SeqH")
Private Sub Test15b_SliceToEnd()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 8)

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Slice(3).ToArray

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


'@TestMethod("SeqH")
Private Sub Test15c_SliceRunOnly()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&, 4&)
    ReDim Preserve myExpected(1 To 4)

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Slice(ipRun:=4).ToArray

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


'@TestMethod("SeqH")
Private Sub Test15d_Slice_Start3_End9_step2()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(3&, 5&, 7&, 9&)
    ReDim Preserve myExpected(1 To 4)

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Slice(3, 7, 2).ToArray

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


'@TestMethod("SeqH")
Private Sub Test15e_Slice_Start3_End9_step2_ToCollection()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(3&, 5&, 7&, 9&)
    ReDim Preserve myExpected(1 To 4)

    Dim myResult As Variant
    ReDim myResult(1 To 4)

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    Dim myC As Collection
    Set myC = mySeq.Slice(3, 7, 2).ToCollection
    myResult(1) = myC.Item(1)
    myResult(2) = myC.Item(2)
    myResult(3) = myC.Item(3)
    myResult(4) = myC.Item(4)

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


'@TestMethod("SeqH")
Private Sub Test15f_Slice_Start3_End9_step2_ToArray()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(3&, 5&, 7&, 9&)
    ReDim Preserve myExpected(1 To 4)

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Slice(3, 7, 2).ToArray

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


'@TestMethod("SeqH")
Private Sub Test16a_Head()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(1&)
    ReDim Preserve myExpected(1 To 1)

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Head.ToArray

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


'@TestMethod("SeqH")
Private Sub Test16b_Head_3Items()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&)
    ReDim Preserve myExpected(1 To 3)

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Head(3).ToArray

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


'@TestMethod("SeqH")
Private Sub Test16c_HeadZeroItems()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = -1&

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Head(-2).Count

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


'@TestMethod("SeqH")
Private Sub Test16d_HeadFullSeq()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 10)

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Head(42).ToArray

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


'@TestMethod("SeqH")
Private Sub Test17a_Tail()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 9)

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Tail.ToArray

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


'@TestMethod("SeqH")
Private Sub Test17b_Tail_3Items()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 7)

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Tail(3).ToArray

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


'@TestMethod("SeqH")
Private Sub Test17c_TailFullItems()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = -1&

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Tail(42).Count

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


'@TestMethod("SeqH")
Private Sub Test17d_TailZeroSeq()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 10)

    Dim myResult As Variant

    Set mySeq = SeqH(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myResult = mySeq.Tail(-2).ToArray

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


'@TestMethod("SeqH")
Private Sub Test18a_KnownIndexes_Available()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 9&, 10&)
    ReDim Preserve myExpected(1 To 4)

    Dim myResult As Variant
    ReDim myResult(1 To 4)

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult(1) = mySeq.FirstIndex
    myResult(2) = mySeq.FBOIndex
    myResult(3) = mySeq.LBOIndex
    myResult(4) = mySeq.LastIndex

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


'@TestMethod("SeqH")
Private Sub Test18b_KnownIndexes_Unavailable()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(-1&, -1&, -1&, -1&)
    ReDim Preserve myExpected(1 To 4)

    Dim myResult As Variant
    ReDim myResult(1 To 4)

    Set mySeq = SeqH.Deb

    'Act:
    myResult(1) = mySeq.FirstIndex
    myResult(2) = mySeq.FBOIndex
    myResult(3) = mySeq.LBOIndex
    myResult(4) = mySeq.LastIndex

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


'@TestMethod("SeqH")
Private Sub Test19a_KnownValues_Available()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 90&, 100&)
    ReDim Preserve myExpected(1 To 4)

    Dim myResult As Variant
    ReDim myResult(1 To 4)

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult(1) = mySeq.First
    myResult(2) = mySeq.FBO
    myResult(3) = mySeq.LBO
    myResult(4) = mySeq.Last

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


'
'@TestMethod("SeqH")
Private Sub Test20a_IndexOf_WholeSeq_Present()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = 5&

    Dim myResult As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.IndexOf(50&)

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


'@TestMethod("SeqH")
Private Sub Test20b_IndexOf_WholeSeq_NotPresent()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = -1&

    Dim myResult As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.IndexOf(55&)

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


'@TestMethod("SeqH")
Private Sub Test20c_IndexOf_SubSeq_Present()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = 5&

    Dim myResult As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.IndexOf(50&, 4, 4)

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


'@TestMethod("SeqH")
Private Sub Test20d_IndexOf_SubSeq_NotPresent()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = -1&

    Dim myResult As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.IndexOf(20&, 4, 4)

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


'@TestMethod("SeqH")
Private Sub Test21a_LastIndexOf_WholeSeq_Present()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = 5&

    Dim myResult As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.LastIndexOf(50&)

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


'@TestMethod("SeqH")
Private Sub Test21b_LastIndexOf_WholeSeq_NotPresent()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = -1&

    Dim myResult As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.LastIndexOf(55&)

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


'@TestMethod("SeqH")
Private Sub Test21c_LastIndexOf_SubSeq_Present()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = 5&

    Dim myResult As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.LastIndexOf(50&, 4, 4)

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


'@TestMethod("SeqH")
Private Sub Test21d_LastIndexOf_SubSeq_NotPresent()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = -1&

    Dim myResult As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.LastIndexOf(20&, 4, 4)

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


'@TestMethod("SeqH")
Private Sub Test22a_Push()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&, 1000&)
    ReDim Preserve myExpected(1 To 11)

    Dim myResult As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.Push(1000&).ToArray

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


'@TestMethod("SeqH")
Private Sub Test22b_PushRange()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&, 11&, 12&, 13&, 14&, 15&)
    ReDim Preserve myExpected(1 To 15)

    Dim myResult As Variant

    Dim myArray As Variant
    myArray = Array(11&, 12&, 13&, 14&, 15&)
    Set mySeq = SeqH(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.PushRange(myArray).ToArray

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


'@TestMethod("SeqH")
Private Sub Test23a_Pop()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = 100&

    Dim myExpected2 As Variant
    myExpected2 = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&)
    ReDim Preserve myExpected2(1 To 9)

    Dim myResult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.Pop
    myResult2 = mySeq.ToArray

    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqH")
Private Sub Test23b_PopRange()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(100&, 90&, 80&, 70&)
    ReDim Preserve myExpected(1 To 4)

    Dim myExpected2 As Variant
    myExpected2 = Array(10&, 20&, 30&, 50&, 50&, 60&)
    ReDim Preserve myExpected2(1 To 6)

    Dim myResult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.PopRange(4).ToArray
    myResult2 = mySeq.ToArray

    'Assert:
    AssertExactSequenceEquals myExpected, myResult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqH")
Private Sub Test23c_PopRange_ExceedsHost()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(100&, 90&, 80&, 70&, 60&, 50&, 40&, 30&, 20&, 10&)
    ReDim Preserve myExpected(1 To 10)
    Dim myExpected2 As Variant
    myExpected2 = -1&

    Dim myResult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.PopRange(25).ToArray
    myResult2 = mySeq.Count

    'Assert:
    AssertExactSequenceEquals myExpected, myResult, myProcedureName
    AssertExactAreEqual myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqH")
Private Sub Test23d_PopRange_NegativeRun()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = -1&

    Dim myExpected2 As Variant
    myExpected2 = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected2(1 To 10)

    Dim myResult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.PopRange(-2).Count
    myResult2 = mySeq.ToArray

    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqH")
Private Sub Test24a_Enqueue()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&, 1000&)
    ReDim Preserve myExpected(1 To 11)

    Dim myResult As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.enQueue(1000&).ToArray

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


'@TestMethod("SeqH")
Private Sub Test24b_EnqueueRange()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&, 11&, 12&, 13&, 14&, 15&)
    ReDim Preserve myExpected(1 To 15)

    Dim myResult As Variant

    Dim myArray As Variant
    myArray = Array(11&, 12&, 13&, 14&, 15&)
    Set mySeq = SeqH(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.EnqueueRange(myArray).ToArray

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


'@TestMethod("SeqH")
Private Sub Test25a_Dequeue()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = 10&

    Dim myExpected2 As Variant
    myExpected2 = Array(20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected2(1 To 9)

    Dim myResult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.Dequeue
    myResult2 = mySeq.ToArray

    'Assert:
    AssertExactAreEqual myExpected, myResult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqH")
Private Sub Test25b_DeqeueRange()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&)
    ReDim Preserve myExpected(1 To 4)

    Dim myExpected2 As Variant
    myExpected2 = Array(50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected2(1 To 6)

    Dim myResult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.DequeueRange(4).ToArray
    myResult2 = mySeq.ToArray

    'Assert:
    AssertExactSequenceEquals myExpected, myResult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqH")
Private Sub Test25c_DequeueRange_ExceedsHost()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 10)

    Dim myExpected2 As Variant
    myExpected2 = -1&


    Dim myResult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.DequeueRange(25).ToArray
    myResult2 = mySeq.Count

    'Assert:
    AssertExactSequenceEquals myExpected, myResult, myProcedureName
    AssertExactAreEqual myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqH")
Private Sub Test26a_Sort()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 10)

    Dim myResult As Variant

    Set mySeq = SeqH(30&, 70&, 40&, 50&, 60&, 80&, 20&, 90&, 10&, 100&)

    'Act:
    myResult = mySeq.Sort.ToArray

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


'@TestMethod("SeqH")
Private Sub Test26b_Sorted()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 10)

   

    Dim myResult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqH(30&, 70&, 40&, 50&, 60&, 80&, 20&, 90&, 10&, 100&)

    'Act:
    myResult = mySeq.Sorted.ToArray
    myResult2 = mySeq.ToArray

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


'@TestMethod("SeqH")
Private Sub Test27a_Reverse()
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
    myExpected = Array(100&, 90&, 80&, 70&, 60&, 50&, 40&, 30&, 20&, 10&)
    ReDim Preserve myExpected(1 To 10)

    Dim mySeq As SeqH
    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    Dim myResult As Variant
    'Act:
    myResult = mySeq.Reverse.ToArray

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


'@TestMethod("SeqH")
Private Sub Test27b_Reversed()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(100&, 90&, 80&, 70&, 60&, 50&, 40&, 30&, 20&, 10&)
    ReDim Preserve myExpected(1 To 10)

    Dim myExpected2 As Variant
    myExpected2 = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected2(1 To 10)

    Dim myResult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myResult = mySeq.Reverse.ToArray
    myResult2 = mySeq.Reversed.ToArray

    'Assert:
    AssertExactSequenceEquals myExpected, myResult, myProcedureName
    AssertExactSequenceEquals myExpected, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqH")
Private Sub Test28a_Unique()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 10)

    Dim myResult As Variant

    Set mySeq = SeqH(10&, 100&, 20&, 30&, 40&, 50&, 30&, 30&, 60&, 100&, 70&, 100&, 80&, 90&, 100&)

    'Act:
    ' The array needs to be sorted because unique copies the first item encountered
    myResult = mySeq.Dedup.Sorted.ToArray

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


'@TestMethod("SeqH")
Private Sub Test28b_Unique_SingleItem()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&)
    ReDim Preserve myExpected(1 To 1)

    Dim myResult As Variant

    Set mySeq = SeqH.Deb.AddItems(10&)

    'Act:
    ' The array needs to be sorted because unique copies the first item encountered
    myResult = mySeq.Dedup.Sorted.ToArray

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


'@TestMethod("SeqH")
Private Sub Test28c_Unique_NoItems()
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = -1&

    Dim myResult As Variant

    Set mySeq = SeqH.Deb

    'Act:
    ' The array needs to be sorted because unique copies the first item encountered
    myResult = mySeq.Count

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


'@TestMethod("SeqH")
Private Sub Test29a_SetOfCommon()

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
    myExpected = Array(50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 6)

    Dim myResult As Variant
    Dim myLS As SeqH: Set myLS = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    Dim myRS As SeqH: Set myRS = SeqH(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)

    'Act:
    Set myResult = myLS.SetOf(m_Common, myRS)
    myResult = myResult.ToArray

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


'@TestMethod("SeqH")
Private Sub Test29b_SetOfHostOnly()

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
    myExpected = Array(110&, 120&, 130&, 140&)
    ReDim Preserve myExpected(1 To 4)

    Dim myResult As Variant

    Dim myLS As SeqH:     Set myLS = SeqH(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)
    Dim myRS As SeqH: Set myRS = SeqH(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    'Act:
    myResult = myLS.SetOf(m_HostOnly, myRS).ToArray

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


'@TestMethod("SeqH")
Private Sub Test29c_SetOfParamOnly()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&)
    ReDim Preserve myExpected(1 To 4)

    Dim myResult As Variant

    Set mySeq = SeqH(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)

    'Act:
    Set myResult = mySeq.SetOf(m_ParamOnly, SeqH(Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)))
    myResult = myResult.ToArray

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


'@TestMethod("SeqH")
Private Sub Test29d_SetOfNotCommon()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 110&, 120&, 130&, 140&)
    ReDim Preserve myExpected(1 To 8)

    Dim myResult As Variant

    Set mySeq = SeqH(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)

    'Act:  Again we need to sort The result SeqH to get the matching array
    myResult = mySeq.SetOf(m_NotCommon, SeqH(Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&))).Sorted.ToArray

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


'@TestMethod("SeqH")
Private Sub Test29e_SetOfUnique()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)
    ReDim Preserve myExpected(1 To 14)

    Dim myResult As Variant

    Set mySeq = SeqH(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)

    'Act:  Again we need to sort The result SeqH to get the matching array
    myResult = mySeq.SetOf(e_SetoF.m_Unique, SeqH(Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&))).Sorted.ToArray

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


'@TestMethod("SeqH")
Private Sub Test30a_Swap()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqH
    Dim myExpected As Variant
    myExpected = Array(140&, 130&, 120&, 110&, 100&, 90&, 80&, 70&, 60&, 50&)
    ReDim Preserve myExpected(1 To 10)

    Dim myResult As Variant

    Set mySeq = SeqH(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)

    'Act:
    mySeq.Swap 1, 10
    mySeq.Swap 2, 9
    mySeq.Swap 3, 8
    mySeq.Swap 4, 7
    mySeq.Swap 5, 6

    myResult = mySeq.ToArray

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


