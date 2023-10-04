Attribute VB_Name = "TestStrs"
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
    ' BUT only runs if the Testing is via rubberduck unit tests
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

Public Sub StrsTests()
 
    #If twinbasic Then
        Debug.Print CurrentProcedureName, vbTab, vbTab,
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName, vbTab, vbTab,
    #End If

    Test01a_strs_BinToNum_Byte
    Test01b_strs_BinToNum_Integer
    Test01c_strs_BinToNum_Long
    Test01d_strs_BinToNum_LongLong
    
    Test02a_Strs_Dedup_Default
    Test02b_Strs_Dedup
    
    Test03a_Strs_Trimmer_Default
    Test03b_Strs_Trimmer_Spaces
    
    Test04a_Strs_PadRight_Default
    Test04b_Strs_PadRight_hash
    Test04c_Strs_PadRight_abcd
    
    Test05a_Strs_PadLeft_Default
    Test05b_Strs_PadLeft_hash
    Test05c_Strs_PadLeft_abcd
    
    Test06a_Strs_CountOf_char
    Test06b_Strs_CountOf_subStr
    
    Test07a_SubStr_Default
    Test07b_SubStr_ab
    
    Test08a_Repeat_Default
    Test08b_Repeat_Hello
    
    Test09a_Replacer
    Test09b_Replacer_SpecifiedPair
    Test09c_Replacer_NestedPairs
    
    Test10a_MultiReplacer
    Test10b_MultiReplacer_SpecifiedPair
    Test10c_MultiReplacer_NestedPairs
    
    Test11a_ToAscB
    Test11b_ToUnicodeBytes
    Test11c_ToUnicodeIntegers
    
    Test12a_Sort
    
    Test13a_Inc_NoCarryInc
    Test13b_Inc_LastCharNotIncrementable
    Test13c_Inc_FullROllover
    Test13d_Inc_NonIncMidString
    
    Debug.Print "Testing completed"

End Sub


'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'@TestMethod("Strs")
Private Sub Test01a_strs_BinToNum_Byte()

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
    myExpected = Array(CByte(0), CByte(3), CByte(128), CByte(255))
    ReDim Preserve myExpected(1 To 4)
    
    Dim myResult As Variant
    ReDim myResult(1 To 4)

    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult(1) = Strs.BinToNum("00")
    myResult(2) = Strs.BinToNum("0011")
    myResult(3) = Strs.BinToNum("10000000")
    myResult(4) = Strs.BinToNum("11111111")
    
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


'@TestMethod("Strs")
Private Sub Test01b_strs_BinToNum_Integer()

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
    myExpected = Array(CInt(0), CInt(3), CInt(-3), CInt(3), CInt(128), CInt(255))
    ReDim Preserve myExpected(1 To 6)
    
    Dim myResult As Variant
    ReDim myResult(1 To 6)

    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult(1) = Strs.BinToNum("0_0000_0000")
    myResult(2) = Strs.BinToNum("00_0000_0011")
    myResult(3) = Strs.BinToNum("1000_0000_0000_0011")
    myResult(4) = Strs.BinToNum("0000_0000_0000_0011")
    myResult(5) = Strs.BinToNum("0000_0000_1000_0000")
    myResult(6) = Strs.BinToNum("0000_0000_1111_1111")
    
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


'@TestMethod("Strs")
Private Sub Test01c_strs_BinToNum_Long()

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
    myExpected = Array(CLng(0), CLng(3), CLng(-3), CLng(3), CLng(128), CLng(255))
    ReDim Preserve myExpected(1 To 6)
    
    Dim myResult As Variant
    ReDim myResult(1 To 6)

    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult(1) = Strs.BinToNum("0_0000_0000_0000_0000_0000_0000")
    myResult(2) = Strs.BinToNum("00_0000_0000_0000_0000_0000_0011")
    myResult(3) = Strs.BinToNum("1000_0000_0000_0000_0000_0000_0000_0011")
    myResult(4) = Strs.BinToNum("0000_0000_0000_0000_0000_0000_0000_0011")
    myResult(5) = Strs.BinToNum("0000_0000_0000_0000_0000_0000_1000_0000")
    myResult(6) = Strs.BinToNum("0000_0000_0000_0000_0000_0000_1111_1111")
    
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


'@TestMethod("Strs")
Private Sub Test01d_strs_BinToNum_LongLong()

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
    myExpected = Array(CLngLng(0), CLngLng(3), CLngLng(-3), CLngLng(3), CLngLng(128), CLngLng(255))
    ReDim Preserve myExpected(1 To 6)
    
    Dim myResult As Variant
    ReDim myResult(1 To 6)

    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myResult(1) = Strs.BinToNum("0_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000")
    myResult(2) = Strs.BinToNum("00_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0011")
    myResult(3) = Strs.BinToNum("1000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0011")
    myResult(4) = Strs.BinToNum("0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0011")
    myResult(5) = Strs.BinToNum("0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_1000_0000")
    myResult(6) = Strs.BinToNum("0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_1111_1111")
    
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


'@TestMethod("Strs")
Private Sub Test02a_Strs_Dedup_Default()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As String
    myExpected = "Hello World"
    
    Dim myResult As String
   
   
    'Act:
   
    myResult = Strs.Dedup("Hello   World")
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


'@TestMethod("Strs")
Private Sub Test02b_Strs_Dedup()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As String
    myExpected = "Helo World"
    
    Dim myResult As String
   
   
    'Act:
   
    myResult = Strs.Dedup("Heeellllo   Worlld", Strs.ToChars("el "))
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


'@TestMethod("Strs")
Private Sub Test03a_Strs_Trimmer_Default()
    
    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As String
    myExpected = "Hello World"
    
    Dim myResult As String
   
   
    'Act:
   
    myResult = Strs.Trimmer(" ,  ;   Hello World ,,,  ;")
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


'@TestMethod("Strs")
Private Sub Test03b_Strs_Trimmer_Spaces()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If

    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As String
    myExpected = "Hello World,"
    
    Dim myResult As String
   
   
    'Act:
   
    myResult = Strs.Trimmer("     Hello World,     ", " ")
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


'@TestMethod("Strs")
Private Sub Test04a_Strs_PadRight_Default()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello World                     "
    
    Dim myResult As String
   
    'Act:
    myResult = Strs.PadRight("Hello World", 32)
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


'@TestMethod("Strs")
Private Sub Test04b_Strs_PadRight_hash()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello World#####################"
    
    Dim myResult As String
   
    'Act:
    myResult = Strs.PadRight("Hello World", 32, "#")
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


'@TestMethod("Strs")
Private Sub Test04c_Strs_PadRight_abcd()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello Worldabcdabcdabcdabcdabcda"
    
    Dim myResult As String
   
    'Act:
    myResult = Strs.PadRight("Hello World", 32, "abcd")
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


'@TestMethod("Strs")
Private Sub Test05a_Strs_PadLeft_Default()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "                     Hello World"
    
    Dim myResult As String
   
    'Act:
    myResult = Strs.PadLeft("Hello World", 32)
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


'@TestMethod("Strs")
Private Sub Test05b_Strs_PadLeft_hash()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "#####################Hello World"
    
    Dim myResult As String
   
    'Act:
    myResult = Strs.PadLeft("Hello World", 32, "#")
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


'@TestMethod("Strs")
Private Sub Test05c_Strs_PadLeft_abcd()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "dabcdabcdabcdabcdabcdHello World"
    
    Dim myResult As String
   
    'Act:
    myResult = Strs.PadLeft("Hello World", 32, "abcd")
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


'@TestMethod("Strs")
Private Sub Test06a_Strs_CountOf_char()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 3
    
    Dim myResult As Long
   
    'Act:
    myResult = Strs.Countof("Hello World", "l")
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


'@TestMethod("Strs")
Private Sub Test06b_Strs_CountOf_subStr()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 2
    
    Dim myResult As Long
   
    'Act:
    myResult = Strs.Countof("Hello Worldel", "el")
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


'@TestMethod("Strs")
Private Sub Test07a_SubStr_Default()

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
    myExpected = Split("Hello,There,World", ",")
    ReDim Preserve myExpected(1 To 3)
    
    
    Dim myResult As Variant
    'Act:
    myResult = Strs.ToSubStr("Hello,There,World")
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


'@TestMethod("Strs")
Private Sub Test07b_SubStr_ab()

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
    myExpected = Split("Hello_ab_There_ab_World", "_ab_")
    ReDim Preserve myExpected(1 To 3)
    
    Dim myResult As Variant
    'Act:
    myResult = Strs.ToSubStr("Hello_ab_There_ab_World", "_ab_")
   
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


'@TestMethod("Strs")
Private Sub Test08a_Repeat_Default()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "        "
   
    Dim myResult As String
    'Act:
    myResult = Strs.Repeat(" ", 8)
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


'@TestMethod("Strs")
Private Sub Test08b_Repeat_Hello()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "HelloHelloHelloHelloHelloHelloHelloHello"
   
    Dim myResult As String
    'Act:
    myResult = Strs.Repeat("Hello", 8)
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


'@TestMethod("Strs")
Private Sub Test09a_Replacer()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "HelloWorld"
   
    Dim myResult As String
    'Act:
    myResult = Strs.Replacer("    He llo   Worl   d ")
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


'@TestMethod("Strs")
Private Sub Test09b_Replacer_SpecifiedPair()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "aaaaaHealloaaaWorlaaada"
   
    Dim myResult As String
    'Act:
    myResult = Strs.Replacer("     He llo   Worl   d ", Chars.twSpace, "a")
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


'@TestMethod("Strs")
Private Sub Test09c_Replacer_NestedPairs()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "HelloWorld"
   
    Dim myResult As String
    'Act:
    myResult = Strs.Replacer("HelloaaaaapppppWorld", "ap", Chars.twNullStr)
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


'@TestMethod("Strs")
Private Sub Test10a_MultiReplacer()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "HelloWorld"
   
    Dim myResult As String
    'Act:
    myResult = Strs.MultiReplacer("    He llo   Worl   d ")
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


'@TestMethod("Strs")
Private Sub Test10b_MultiReplacer_SpecifiedPair()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "aaaaaHealloaaaWorlaaada"
   
    Dim myResult As String
    'Act:
    myResult = Strs.MultiReplacer("     He llo   Worl   d ", Array(Array(Chars.twSpace, "a")))
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


'@TestMethod("Strs")
Private Sub Test10c_MultiReplacer_NestedPairs()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "HeLLoWorLd"
   
    Dim myResult As String
    'Act:
    myResult = Strs.MultiReplacer("HelloaaaaapppppWorld", Array(Array("ap", Chars.twNullStr), Array("l", "L")))
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


'@TestMethod("Strs")
Private Sub Test11a_ToAscB()

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
    myExpected = Array(AscB("H"), AscB("e"), AscB("l"), AscB("l"), AscB("o"))
    ReDim Preserve myExpected(1 To 5)
    Dim myResult As Variant
    'Act:
    myResult = Strs.ToAscB("Hello")
    ReDim Preserve myResult(1 To 5)
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


'@TestMethod("Strs")
Private Sub Test11b_ToUnicodeBytes()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myTmp() As Byte
    myTmp = "Hello"
    Dim myExpected() As Variant
    ReDim Preserve myExpected(1 To 10)
    Dim myIndex As Long
    For myIndex = 0 To 9
        myExpected(myIndex + 1) = myTmp(myIndex)
    Next
    
    Dim myResult As Variant
    'Act:
    myResult = Strs.ToUnicodeBytes("Hello")
    ReDim Preserve myResult(1 To 10)
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


'@TestMethod("Strs")
Private Sub Test11c_ToUnicodeIntegers()

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
    myExpected = Array(AscW("H"), AscW("e"), AscW("l"), AscW("l"), AscW("o"))
    ReDim Preserve myExpected(1 To 5)
    
    
    Dim myResult As Variant
    'Act:
    myResult = Strs.ToUnicodeIntegers("Hello")
    ReDim Preserve myResult(1 To 5)
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


'@TestMethod("Strs")
Private Sub Test12a_Sort()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello"
        
    Dim myResult As String
    'Act:
    myResult = Strs.Sort("oleHl")
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


'@TestMethod("Strs")
Private Sub Test13a_Inc_NoCarryInc()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Helm0"
        
    Dim myString As String
    myString = "Hellz"
    
    Dim myResult As String
    
    'Act:
    myResult = Strs.Inc(myString)
    
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


'@TestMethod("Strs")
Private Sub Test13b_Inc_LastCharNotIncrementable()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello/1"
    
    Dim myString As String
    myString = "Hello/"
    
    Dim myResult As String
    
    'Act:
    myResult = Strs.Inc(myString)
    
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


'@TestMethod("Strs")
Private Sub Test13c_Inc_FullROllover()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "100000"
        
    
    Dim myString As String
    myString = "zzzzz"
    
    Dim myResult As String
    
    'Act:
    myResult = Strs.Inc(myString)
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


'@TestMethod("Strs")
Private Sub Test13d_Inc_NonIncMidString()

    #If twinbasic Then
        myProcedureName = myComponentName & ":" & CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ModuleName & ":" & ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hel/100"
        
    Dim myString As String
    myString = "Hel/zz"
    
    Dim myResult As String
    
    'Act:
    myResult = Strs.Inc(myString)
    
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


