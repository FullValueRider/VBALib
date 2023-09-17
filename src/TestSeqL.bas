Attribute VB_Name = "TestSeqL"
'@IgnoreModule
'@TestModule
'@Folder("Tests")
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

Public Sub SeqLTests()
 
    #If TWINBASIC Then
        Debug.Print CurrentProcedureName,
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName,
    #End If

    Test01_SeqObj
    
    Test02a_InitByLong_10FirstIndex_LastIndex
    Test02b_InitByString
    Test02c_InitByForEachArray
    Test02d_InitByForEachArrayList
    Test02e_InitByForEachCollection
    Test02f_InitByDictionary
    
    Test03a_WriteReadItems
    
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
    
    'Test09a0_Remove_SingleItem
    
    Test09a_RemoveAt_SingleItem
    Test09b_RemoveAt_ThreeItems
    
    Test10a_Remove_SingleItem
    
    ' Test11a_RemoveRange_SingleItem
    ' Test11b_RemoveRange_ThreeItems
    
    Test12a_RemoveRange_SingleItem
    
    Test13a_RemoveAll_DefaultAll
    Test13b_RemoveAll_42AndHello
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
    Test20c_IndexOf_SubSeq_ItemIsPresent
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
    
    Debug.Print vbTab, vbTab, "Testing completed"

End Sub


'@TestMethod("SeqL")
Private Sub Test01_SeqObj()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail
    
    'Arrange:
    Dim mySeq As SeqL
    Set mySeq = SeqL.Deb
    Dim myExpected As Variant
    myExpected = Array(True, "SeqL", "SeqL")
    
    Dim myresult(0 To 2) As Variant
    
    'Act:
    myresult(0) = VBA.IsObject(mySeq)
    myresult(1) = VBA.TypeName(mySeq)
    myresult(2) = mySeq.TypeName
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


'@TestMethod("SeqL")
Private Sub Test02a_InitByLong_10FirstIndex_LastIndex()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Set mySeq = SeqL.Deb.Fill(Empty, 10)
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 10)
    Dim myresult As Variant

    'Act:
    myresult = mySeq.ToArray
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    AssertExactAreEqual 1&, mySeq.FirstIndex, myProcedureName
    AssertExactAreEqual 10&, mySeq.Lastindex, myProcedureName
    AssertExactAreEqual 10&, mySeq.Count, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test02b_InitByString()
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
    myExpected = Array("H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 5)

    Dim myresult As Variant

    Dim mySeq As SeqL

    'Act:
    Set mySeq = SeqL.Deb("Hello")

    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test02c_InitByForEachArray()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim myArray(1 To 3, 1 To 3) As Variant
    Dim myCount As Long
    myCount = 1
    Dim myFirst As Long
    For myFirst = 1 To 3

        Dim mySecond As Long
        For mySecond = 1 To 3
            myArray(myFirst, mySecond) = myCount
            myCount = myCount + 1
        Next
    Next

    Dim myExpected As Variant
    myExpected = Array(1&, 4&, 7&, 2&, 5&, 8&, 3&, 6&, 9&)
    ReDim Preserve myExpected(1 To 9)

    Dim myresult As Variant

    Dim mySeq As SeqL

    'Act:
    Set mySeq = SeqL.Deb(myArray)

    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test02d_InitByForEachArrayList()
    #If TWINBASIC Then
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

    Dim myresult As Variant

    Dim mySeq As SeqL

    'Act:
    Set mySeq = SeqL.Deb(myAL)

    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test02e_InitByForEachCollection()
    #If TWINBASIC Then
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

    Dim myresult As Variant

    Dim mySeq As SeqL

    'Act:
    Set mySeq = SeqL.Deb(myC)

    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test02f_InitByDictionary()
    #If TWINBASIC Then
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

    Dim myresult As Variant
    ReDim myresult(1 To 6)

    Dim mySeq As SeqL

    'Act:
    Set mySeq = SeqL.Deb(myD)
    Dim myTmp As Variant
    myTmp = mySeq.ToArray

    myresult(1) = myTmp(1)(0)
    myresult(2) = myTmp(1)(1)
    myresult(3) = myTmp(2)(0)
    myresult(4) = myTmp(2)(1)
    myresult(5) = myTmp(3)(0)
    myresult(6) = myTmp(3)(1)

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


'@TestMethod("SeqL")
Private Sub Test03a_WriteReadItems()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Set mySeq = SeqL.Deb.Fill(Empty, 10)
    mySeq.Item(1) = 42
    mySeq.Item(2) = "Hello"
    mySeq.Item(3) = 3.142
    Dim myExpected As Variant
    myExpected = Array(True, True, True)

    Dim myresult As Variant
    ReDim myresult(0 To 2)
    'Act:
    myresult(0) = mySeq.Item(1) = 42
    myresult(1) = mySeq.Item(2) = "Hello"
    myresult(2) = mySeq.Item(3) = "3.142"

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


'@TestMethod("SeqL")
Private Sub Test04a_Add_MultipleItems()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, 42, "Hello", 3.142)
    ReDim Preserve myExpected(1 To 8)

    Dim myresult As Variant

    'Act:
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    myresult = mySeq.AddItems(42, "Hello", 3.142).ToArray

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


'@TestMethod("SeqL")
Private Sub Test06a_AddRange_String()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 10)

    Dim myresult As Variant

    'Act:
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    myresult = mySeq.AddRange("Hello").ToArray

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


'@TestMethod("SeqL")
Private Sub Test06b_AddRange_Array()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 10)

    Dim myresult As Variant

    'Act:
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    myresult = mySeq.AddRange(Array("H", "e", "l", "l", "o")).ToArray

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


'@TestMethod("SeqL")
Private Sub Test06c_AddRange_Collection()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 10)

    Dim myresult As Variant

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
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    myresult = mySeq.AddRange(myC).ToArray

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


'@TestMethod("SeqL")
Private Sub Test06d_AddRange_ArrayList()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "H", "e", "l", "l", "o")
    ReDim Preserve myExpected(1 To 10)

    Dim myresult As Variant

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
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    myresult = mySeq.AddRange(myAL).ToArray

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


'@TestMethod("SeqL")
Private Sub Test06e_AddRange_Dictionary()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty, "Hello1", "There2", "World3")
    ReDim Preserve myExpected(1 To 8)

    Dim myresult As Variant

    Dim myD As KvpC
    Set myD = KvpC.Deb
    With myD
        .Add "Hello", 1
        .Add "There", 2
        .Add "World", 3
    End With

    'Act:
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    myresult = mySeq.AddRange(myD).ToArray
    myresult(6) = myresult(6)(0) & VBA.CStr(myresult(6)(1))
    myresult(7) = myresult(7)(0) & VBA.CStr(myresult(7)(1))
    myresult(8) = myresult(8)(0) & VBA.CStr(myresult(8)(1))

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


'@TestMethod("SeqL")
Private Sub Test07a_Insert_SingleItems()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", Empty, 42&, Empty, 3.142, Empty)
    ReDim Preserve myExpected(1 To 8)
    Dim myExpected2 As Variant
    myExpected2 = Array(3&, 5&, 7&)

    Dim myresult As Variant
    Dim myResult2 As Variant
    ReDim myResult2(0 To 2)

    'Act:
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    myResult2(0) = mySeq.InsertAt(3, "Hello")
    myResult2(1) = mySeq.InsertAt(5, 42&)
    myResult2(2) = mySeq.InsertAt(7, 3.142)

    myresult = mySeq.ToArray
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test07b_Insert_MultipleItems()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)

    Dim myresult As Variant

    'Act:
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    
    mySeq.InsertAtItems 3, "Hello", 42&, 3.142

    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test09a_InsertAtItems()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)
    Dim myExpected2 As Variant
    myExpected2 = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected2(1 To 8)
    Dim myresult As Variant
    Dim myResult2 As Variant


    'Act:

    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    myResult2 = mySeq.InsertAtItems(3, "Hello", 42&, 3.142).ToArray


    myresult = mySeq.ToArray
    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test08a_InsertAtRange_String()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL

    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "H", "e", "l", "l", "o", Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 10)

    Dim myresult As Variant

    'Act:
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    mySeq.InsertAtRange 3, "Hello"

    myresult = mySeq.ToArray
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


'@TestMethod("SeqL")
Private Sub Test08b_InsertAtRange_Array()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)

    Dim myresult As Variant

    'Act:
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    mySeq.InsertAtRange 3, Array("Hello", 42&, 3.142)

    myresult = mySeq.ToArray
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


'@TestMethod("SeqL")
Private Sub Test08c_InsertAtRange_Collection()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)

    Dim myresult As Variant


    Dim myC As Collection
    Set myC = New Collection
    With myC
        .Add "Hello"
        .Add 42&
        .Add 3.142

    End With

    'Act:
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    mySeq.InsertAtRange 3, myC

    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test08d_InsertAtRange_ArrayList()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello", 42&, 3.142, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)

    Dim myresult As Variant


    Dim myAL As Object
    Set myAL = CreateObject("System.collections.arraylist")
    With myAL
        .Add "Hello"
        .Add 42&
        .Add 3.142

    End With

    'Act:
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    mySeq.InsertAtRange 3, myAL

    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test08e_InsertAtRange_Dictionary()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, "Hello1", "There2", "World3", Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 8)

    Dim myresult As Variant

    Dim myD As KvpC
    Set myD = KvpC.Deb
    With myD
        .Add "Hello", 1
        .Add "There", 2
        .Add "World", 3

    End With

    'Act:
    Set mySeq = SeqL.Deb.Fill(Empty, 5)
    mySeq.InsertAtRange 3, myD
    myresult = mySeq.ToArray

    myresult(3) = myresult(3)(0) & VBA.CStr(myresult(3)(1))
    myresult(4) = myresult(4)(0) & VBA.CStr(myresult(4)(1))
    myresult(5) = myresult(5)(0) & VBA.CStr(myresult(5)(1))


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


'@TestMethod("SeqL")
Private Sub Test09a_RemoveAt_SingleItem()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)

    Dim myresult As Variant
    Set mySeq = SeqL.Deb(Array(Empty, Empty, Empty, 42, Empty, Empty))

    'Act:
    mySeq.RemoveAt 4

    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test09b_RemoveAt_ThreeItems()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)

    Dim myresult As Variant
    Set mySeq = SeqL.Deb(Array(Empty, 42, Empty, Empty, 42, Empty, Empty, 42))

    'Act:
    mySeq.RemoveIndexes 8, 2, 5

    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test10a_Remove_SingleItem()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, 42, "Hello", "Hello", Empty, Empty, 42, Empty, Empty)
    ReDim Preserve myExpected(1 To 11)

    Dim myExpectedIndex As Long
    myExpectedIndex = 2&
    
    Dim myresult As Variant
    Set mySeq = SeqL(Empty, 42, Empty, Empty, 42, "Hello", "Hello", Empty, Empty, 42, Empty, Empty)
    Dim myResultIndex As Long
    'Act:
    myResultIndex = mySeq.Remove(42&)

    myresult = mySeq.ToArray

    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    AssertExactAreEqual myExpectedIndex, myResultIndex, myProcedureName
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test11a_RemoveIndexesRange_SingleItem()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)

    Dim myresult As Variant
    Set mySeq = SeqL.Deb(Array(Empty, Empty, Empty, 42, Empty, Empty))

    'Act:
    mySeq.RemoveRange SeqL.Deb.AddItems(42).ToArray

    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test11b_RemoveIndexesRange_ThreeItems()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, Empty, Empty)
    ReDim Preserve myExpected(1 To 5)

    Dim myresult As Variant
    Set mySeq = SeqL.Deb(Array(Empty, Empty, Empty, 42, 42, 42, Empty, Empty))

    'Act:
    mySeq.RemoveIndexesRange SeqL(5, 4, 6).ToArray

    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test12a_RemoveRange_SingleItem()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(Empty, Empty, Empty, 42, 42, Empty, Empty)
    ReDim Preserve myExpected(1 To 7)

    Dim myresult As Variant
    Set mySeq = SeqL(Empty, Empty, Empty, 42, 42, 42, Empty, Empty)

    'Act:
    mySeq.RemoveRange SeqL.Deb.AddItems(42).ToArray

    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test13a_RemoveAll_DefaultAll()
    
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = -1&

    Dim myresult As Variant
    Set mySeq = SeqL(Empty, Empty, Empty, 42, 42, 42, Empty, Empty)

    'Act:
    mySeq.RemoveAll
    myresult = mySeq.Count

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


'@TestMethod("SeqL")
Private Sub Test13b_RemoveAll_42AndHello()
    
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    'on error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(1, 2, 3, 4, 5)
    ReDim Preserve myExpected(1 To 5)

    Dim myresult As Variant
    Set mySeq = SeqL(1, "Hello", 2, "Hello", "Hello", 3, 42, 42, 42, 4, 5)

    'Act:
    mySeq.RemoveAll "Hello", 42
    myresult = mySeq.ToArray

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


'@TestMethod("SeqL")
Private Sub Test13c_Reset()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = -1&

    Dim myresult As Variant
    Set mySeq = SeqL.Deb(Array(Empty, Empty, Empty, 42, 42, 42, Empty, Empty))

    'Act:
    mySeq.Reset
    myresult = mySeq.Count

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


'@TestMethod("SeqL")
Private Sub Test13d_Clear()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = -1&

    Dim myresult As Variant
    Set mySeq = SeqL.Deb(Array(Empty, Empty, Empty, 42, 42, 42, Empty, Empty))

    'Act:
    mySeq.Clear
    myresult = mySeq.Count

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


'@TestMethod("SeqL")
Private Sub Test14a_Fill()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(True, True, True)
    ReDim Preserve myExpected(1 To 3)

    Dim myresult As Variant
    ReDim myresult(1 To 3)
    Set mySeq = SeqL.Deb(Array(Empty, Empty, Empty))

    'Act:
    mySeq.Fill 42, 10
    myresult(1) = mySeq.Count = 13
    myresult(2) = mySeq.Item(4) = 42&
    myresult(3) = mySeq.Item(13) = 42&

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


'@TestMethod("SeqL")
Private Sub Test15a_Slice()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(3&, 4&, 5&)
    ReDim Preserve myExpected(1 To 3)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Slice(3, 3).ToArray

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


'@TestMethod("SeqL")
Private Sub Test15b_SliceToEnd()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 8)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Slice(3).ToArray

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


'@TestMethod("SeqL")
Private Sub Test15c_SliceRunOnly()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&, 4&)
    ReDim Preserve myExpected(1 To 4)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Slice(ipRun:=4).ToArray

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


'@TestMethod("SeqL")
Private Sub Test15d_Slice_Start3_End9_step2()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(3&, 5&, 7&, 9&)
    ReDim Preserve myExpected(1 To 4)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Slice(3, 7, 2).ToArray

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


'@TestMethod("SeqL")
Private Sub Test15e_Slice_Start3_End9_step2_ToCollection()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(3&, 5&, 7&, 9&)
    ReDim Preserve myExpected(1 To 4)

    Dim myresult As Variant
    ReDim myresult(1 To 4)

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    Dim myC As Collection
    Set myC = mySeq.Slice(3, 7, 2).ToCollection
    myresult(1) = myC.Item(1)
    myresult(2) = myC.Item(2)
    myresult(3) = myC.Item(3)
    myresult(4) = myC.Item(4)

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


'@TestMethod("SeqL")
Private Sub Test15f_Slice_Start3_End9_step2_ToArray()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(3&, 5&, 7&, 9&)
    ReDim Preserve myExpected(1 To 4)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Slice(3, 7, 2).ToArray

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


'@TestMethod("SeqL")
Private Sub Test16a_Head()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(1&)
    ReDim Preserve myExpected(1 To 1)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Head.ToArray

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


'@TestMethod("SeqL")
Private Sub Test16b_Head_3Items()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&)
    ReDim Preserve myExpected(1 To 3)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Head(3).ToArray

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


'@TestMethod("SeqL")
Private Sub Test16c_HeadZeroItems()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = -1&

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Head(-2).Count

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


'@TestMethod("SeqL")
Private Sub Test16d_HeadFullSeq()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 10)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Head(42).ToArray

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


'@TestMethod("SeqL")
Private Sub Test17a_Tail()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 9)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Tail.ToArray

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


'@TestMethod("SeqL")
Private Sub Test17b_Tail_3Items()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 7)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Tail(3).ToArray

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


'@TestMethod("SeqL")
Private Sub Test17c_TailFullItems()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = -1&

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Tail(42).Count

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


'@TestMethod("SeqL")
Private Sub Test17d_TailZeroSeq()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)
    ReDim Preserve myExpected(1 To 10)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(1&, 2&, 3&, 4&, 5&, 6&, 7&, 8&, 9&, 10&)

    'Act:
    myresult = mySeq.Tail(-2).ToArray

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


'@TestMethod("SeqL")
Private Sub Test18a_KnownIndexes_Available()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, 9&, 10&)
    ReDim Preserve myExpected(1 To 4)

    Dim myresult As Variant
    ReDim myresult(1 To 4)

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult(1) = mySeq.FirstIndex
    myresult(2) = mySeq.FBOIndex
    myresult(3) = mySeq.LBOIndex
    myresult(4) = mySeq.Lastindex

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


'@TestMethod("SeqL")
Private Sub Test18b_KnownIndexes_Unavailable()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(1&, 2&, -1&, -1&)
    ReDim Preserve myExpected(1 To 4)

    Dim myresult As Variant
    ReDim myresult(1 To 4)

    Set mySeq = SeqL.Deb

    'Act:
    myresult(1) = mySeq.FirstIndex
    myresult(2) = mySeq.FBOIndex
    myresult(3) = mySeq.LBOIndex
    myresult(4) = mySeq.Lastindex

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


'@TestMethod("SeqL")
Private Sub Test19a_KnownValues_Available()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 90&, 100&)
    ReDim Preserve myExpected(1 To 4)

    Dim myresult As Variant
    ReDim myresult(1 To 4)

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult(1) = mySeq.First
    myresult(2) = mySeq.FBO
    myresult(3) = mySeq.LBO
    myresult(4) = mySeq.Last

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


'@TestMethod("SeqL")
Private Sub Test20a_IndexOf_WholeSeq_Present()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = 5&

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.IndexOf(50&)

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


'@TestMethod("SeqL")
Private Sub Test20b_IndexOf_WholeSeq_NotPresent()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = -1&

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.IndexOf(55&)

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


'@TestMethod("SeqL")
Private Sub Test20c_IndexOf_SubSeq_ItemIsPresent()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = 5&

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.IndexOf(50&, 4, 4)

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


'@TestMethod("SeqL")
Private Sub Test20d_IndexOf_SubSeq_NotPresent()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = -1&

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.IndexOf(20&, 4, 4)

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


'@TestMethod("SeqL")
Private Sub Test21a_LastIndexOf_WholeSeq_Present()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = 5&

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.LastIndexOf(50&)

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


'@TestMethod("SeqL")
Private Sub Test21b_LastIndexOf_WholeSeq_NotPresent()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = -1&

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.LastIndexOf(55&)

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


'@TestMethod("SeqL")
Private Sub Test21c_LastIndexOf_SubSeq_Present()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = 5&

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.LastIndexOf(50&, 4, 4)

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


'@TestMethod("SeqL")
Private Sub Test21d_LastIndexOf_SubSeq_NotPresent()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = -1&

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.LastIndexOf(20&, 4, 4)

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


'@TestMethod("SeqL")
Private Sub Test22a_Push()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&, 1000&)
    ReDim Preserve myExpected(1 To 11)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.Push(1000&).ToArray

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


'@TestMethod("SeqL")
Private Sub Test22b_PushRange()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&, 11&, 12&, 13&, 14&, 15&)
    ReDim Preserve myExpected(1 To 15)

    Dim myresult As Variant

    Dim myArray As Variant
    myArray = Array(11&, 12&, 13&, 14&, 15&)
    Set mySeq = SeqL.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.PushRange(myArray).ToArray

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


'@TestMethod("SeqL")
Private Sub Test23a_Pop()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = 100&

    Dim myExpected2 As Variant
    myExpected2 = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&)
    ReDim Preserve myExpected2(1 To 9)

    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.Pop
    myResult2 = mySeq.ToArray

    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test23b_PopRange()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(100&, 90&, 80&, 70&)
    ReDim Preserve myExpected(1 To 4)

    Dim myExpected2 As Variant
    myExpected2 = Array(10&, 20&, 30&, 50&, 50&, 60&)
    ReDim Preserve myExpected2(1 To 6)

    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.PopRange(4).ToArray
    myResult2 = mySeq.ToArray

    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test23c_PopRange_ExceedsHost()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(100&, 90&, 80&, 70&, 60&, 50&, 40&, 30&, 20&, 10&)
    ReDim Preserve myExpected(1 To 10)

    Dim myExpected2 As Variant
    myExpected2 = -1&

    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.PopRange(25).ToArray
    myResult2 = mySeq.Count

    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    AssertExactAreEqual myExpected2 = 0, myResult2 = 0, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test23d_PopRange_NegativeRun()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = -1&

    Dim myExpected2 As Variant
    myExpected2 = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected2(1 To 10)

    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.PopRange(-2).Count
    myResult2 = mySeq.ToArray

    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test24a_Enqueue()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&, 1000&)
    ReDim Preserve myExpected(1 To 11)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.enQueue(1000&).ToArray

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


'@TestMethod("SeqL")
Private Sub Test24b_EnqueueRange()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&, 11&, 12&, 13&, 14&, 15&)
    ReDim Preserve myExpected(1 To 15)

    Dim myresult As Variant

    Dim myArray As Variant
    myArray = Array(11&, 12&, 13&, 14&, 15&)
    Set mySeq = SeqL.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.EnqueueRange(myArray).ToArray

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


'@TestMethod("SeqL")
Private Sub Test25a_Dequeue()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = 10&

    Dim myExpected2 As Variant
    myExpected2 = Array(20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected2(1 To 9)

    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 50&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.Dequeue
    myResult2 = mySeq.ToArray

    'Assert:
    AssertExactAreEqual myExpected, myresult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test25b_DeqeueRange()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&)
    ReDim Preserve myExpected(1 To 4)

    Dim myExpected2 As Variant
    myExpected2 = Array(50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected2(1 To 6)

    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.DequeueRange(4).ToArray
    myResult2 = mySeq.ToArray

    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    AssertExactSequenceEquals myExpected2, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test25c_DequeueRange_ExceedsHost()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 10)

    Dim myExpected2 As Variant
    myExpected2 = 0&


    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqL(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.DequeueRange(25).ToArray
    myResult2 = mySeq.Count

    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    AssertExactAreEqual 0, 0, myProcedureName   'Todo:

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test26a_Sort()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 10)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(30&, 70&, 40&, 50&, 60&, 80&, 20&, 90&, 10&, 100&)

    'Act:
    myresult = mySeq.Sort.ToArray

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


'@TestMethod("SeqL")
Private Sub Test26b_Sorted()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 10)

    Dim myExpected2 As Variant
    myExpected2 = Array(30&, 70&, 40&, 50&, 60&, 80&, 20&, 90&, 10&, 100&)
    ReDim Preserve myExpected2(1 To 10)

    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqL.Deb(30&, 70&, 40&, 50&, 60&, 80&, 20&, 90&, 10&, 100&)

    'Act:
    myresult = mySeq.Sorted.ToArray
    myResult2 = mySeq.ToArray

    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    AssertExactSequenceEquals myExpected, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test27a_Reverse()
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
    myExpected = Array(100&, 90&, 80&, 70&, 60&, 50&, 40&, 30&, 20&, 10&)
    ReDim Preserve myExpected(1 To 10)

    Dim mySeq As SeqL
    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    Dim myresult As Variant
    'Act:
    Set myresult = mySeq.Reverse
    myresult = myresult.ToArray

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


'@TestMethod("SeqL")
Private Sub Test27b_Reversed()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(100&, 90&, 80&, 70&, 60&, 50&, 40&, 30&, 20&, 10&)
    ReDim Preserve myExpected(1 To 10)

    Dim myExpected2 As Variant
    myExpected2 = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected2(1 To 10)

    Dim myresult As Variant
    Dim myResult2 As Variant

    Set mySeq = SeqL.Deb(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.Reverse.ToArray
    myResult2 = mySeq.Reversed.ToArray

    'Assert:
    AssertExactSequenceEquals myExpected, myresult, myProcedureName
    AssertExactSequenceEquals myExpected, myResult2, myProcedureName

TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next

    Exit Sub
TestFail:
    AssertFail myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("SeqL")
Private Sub Test28a_Unique()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 10)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(10&, 100&, 20&, 30&, 40&, 50&, 30&, 30&, 60&, 100&, 70&, 100&, 80&, 90&, 100&)

    'Act:
    myresult = mySeq.Dedup.Sort.ToArray

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


'@TestMethod("SeqL")
Private Sub Test28b_Unique_SingleItem()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&)
    ReDim Preserve myExpected(1 To 1)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb.AddItems(10&)

    'Act:
    
    myresult = mySeq.Dedup.ToArray

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


'@TestMethod("SeqL")
Private Sub Test28c_Unique_NoItems()
    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Boolean
    myExpected = False

    Dim myresult As Variant

    Set mySeq = SeqL.Deb

    'Act:
    myresult = mySeq.IsUnique
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


'@TestMethod("SeqL")
Private Sub Test29a_SetOfCommon()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(50&, 60&, 70&, 80&, 90&, 100&)
    ReDim Preserve myExpected(1 To 6)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)

    'Act:
    myresult = mySeq.SetOf(m_Common, SeqL(Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&))).ToArray

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


'@TestMethod("SeqL")
Private Sub Test29b_SetOfHostOnly()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(110&, 120&, 130&, 140&)
    ReDim Preserve myExpected(1 To 4)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)

    'Act:
    myresult = mySeq.SetOf(m_HostOnly, SeqL(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&)).ToArray

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


'@TestMethod("SeqL")
Private Sub Test29c_SetOfParamOnly()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&)
    ReDim Preserve myExpected(1 To 4)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)

    'Act:
    myresult = mySeq.SetOf(m_ParamOnly, SeqL(Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&))).ToArray

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


'@TestMethod("SeqL")
Private Sub Test29d_SetOfNotCommon()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 110&, 120&, 130&, 140&)
    ReDim Preserve myExpected(1 To 8)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)

    'Act:  Again we need to sort The result SeqL to get the matching array
    myresult = mySeq.SetOf(m_NotCommon, SeqL(Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&))).Sorted.ToArray

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


'@TestMethod("SeqL")
Private Sub Test29e_SetOfUnique()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)
    ReDim Preserve myExpected(1 To 14)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)

    'Act:  Again we need to sort The result SeqL to get the matching array
    myresult = mySeq.SetOf(e_SetoF.m_Unique, SeqL(Array(10&, 20&, 30&, 40&, 50&, 60&, 70&, 80&, 90&, 100&))).Sorted.ToArray

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


'@TestMethod("SeqL")
Private Sub Test30a_Swap()

    #If TWINBASIC Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    On Error GoTo TestFail

    'Arrange:
    Dim mySeq As SeqL
    Dim myExpected As Variant
    myExpected = Array(140&, 130&, 120&, 110&, 100&, 90&, 80&, 70&, 60&, 50&)
    ReDim Preserve myExpected(1 To 10)

    Dim myresult As Variant

    Set mySeq = SeqL.Deb(50&, 60&, 70&, 80&, 90&, 100&, 110&, 120&, 130&, 140&)

    'Act:
    mySeq.Swap 1, 10
    mySeq.Swap 2, 9
    mySeq.Swap 3, 8
    mySeq.Swap 4, 7
    mySeq.Swap 5, 6

    myresult = mySeq.ToArray

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


