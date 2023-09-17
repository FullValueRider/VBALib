Attribute VB_Name = "TestStrs"
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
 
    #If TWINBASIC Then
        Debug.Print CurrentProcedureName; vbTab, vbTab, vbTab, vbTab,
    #Else
        GlobalAssert
        VBATesting = True
        Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab, vbTab,
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
    myExpected = Array(CByte(0), CByte(3), CByte(128), CByte(255))
    ReDim Preserve myExpected(1 To 4)
    
    Dim myresult As Variant
    ReDim myresult(1 To 4)

    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult(1) = Strs.BinToNum("00")
    myresult(2) = Strs.BinToNum("0011")
    myresult(3) = Strs.BinToNum("10000000")
    myresult(4) = Strs.BinToNum("11111111")
    
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


'@TestMethod("Strs")
Private Sub Test01b_strs_BinToNum_Integer()

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
    myExpected = Array(CInt(0), CInt(3), CInt(-3), CInt(3), CInt(128), CInt(255))
    ReDim Preserve myExpected(1 To 6)
    
    Dim myresult As Variant
    ReDim myresult(1 To 6)

    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult(1) = Strs.BinToNum("0_0000_0000")
    myresult(2) = Strs.BinToNum("00_0000_0011")
    myresult(3) = Strs.BinToNum("1000_0000_0000_0011")
    myresult(4) = Strs.BinToNum("0000_0000_0000_0011")
    myresult(5) = Strs.BinToNum("0000_0000_1000_0000")
    myresult(6) = Strs.BinToNum("0000_0000_1111_1111")
    
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


'@TestMethod("Strs")
Private Sub Test01c_strs_BinToNum_Long()

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
    myExpected = Array(CLng(0), CLng(3), CLng(-3), CLng(3), CLng(128), CLng(255))
    ReDim Preserve myExpected(1 To 6)
    
    Dim myresult As Variant
    ReDim myresult(1 To 6)

    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult(1) = Strs.BinToNum("0_0000_0000_0000_0000_0000_0000")
    myresult(2) = Strs.BinToNum("00_0000_0000_0000_0000_0000_0011")
    myresult(3) = Strs.BinToNum("1000_0000_0000_0000_0000_0000_0000_0011")
    myresult(4) = Strs.BinToNum("0000_0000_0000_0000_0000_0000_0000_0011")
    myresult(5) = Strs.BinToNum("0000_0000_0000_0000_0000_0000_1000_0000")
    myresult(6) = Strs.BinToNum("0000_0000_0000_0000_0000_0000_1111_1111")
    
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


'@TestMethod("Strs")
Private Sub Test01d_strs_BinToNum_LongLong()

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
    myExpected = Array(CLngLng(0), CLngLng(3), CLngLng(-3), CLngLng(3), CLngLng(128), CLngLng(255))
    ReDim Preserve myExpected(1 To 6)
    
    Dim myresult As Variant
    ReDim myresult(1 To 6)

    
    'Act:  Again we need to sort The result SeqC to get the matching array
    myresult(1) = Strs.BinToNum("0_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000")
    myresult(2) = Strs.BinToNum("00_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0011")
    myresult(3) = Strs.BinToNum("1000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0011")
    myresult(4) = Strs.BinToNum("0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0011")
    myresult(5) = Strs.BinToNum("0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_1000_0000")
    myresult(6) = Strs.BinToNum("0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_0000_1111_1111")
    
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


'@TestMethod("Strs")
Private Sub Test02a_Strs_Dedup_Default()

    #If TWINBASIC Then
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
    
    Dim myresult As String
   
   
    'Act:
   
    myresult = Strs.Dedup("Hello   World")
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


'@TestMethod("Strs")
Private Sub Test02b_Strs_Dedup()

    #If TWINBASIC Then
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
    
    Dim myresult As String
   
   
    'Act:
   
    myresult = Strs.Dedup("Heeellllo   Worlld", Strs.ToChars("el "))
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


'@TestMethod("Strs")
Private Sub Test03a_Strs_Trimmer_Default()
    
    #If TWINBASIC Then
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
    
    Dim myresult As String
   
   
    'Act:
   
    myresult = Strs.Trimmer(" ,  ;   Hello World ,,,  ;")
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


'@TestMethod("Strs")
Private Sub Test03b_Strs_Trimmer_Spaces()

    #If TWINBASIC Then
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
    
    Dim myresult As String
   
   
    'Act:
   
    myresult = Strs.Trimmer("     Hello World,     ", " ")
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


'@TestMethod("Strs")
Private Sub Test04a_Strs_PadRight_Default()

    #If TWINBASIC Then
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
    
    Dim myresult As String
   
    'Act:
    myresult = Strs.PadRight("Hello World", 32)
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


'@TestMethod("Strs")
Private Sub Test04b_Strs_PadRight_hash()

    #If TWINBASIC Then
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
    
    Dim myresult As String
   
    'Act:
    myresult = Strs.PadRight("Hello World", 32, "#")
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


'@TestMethod("Strs")
Private Sub Test04c_Strs_PadRight_abcd()

    #If TWINBASIC Then
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
    
    Dim myresult As String
   
    'Act:
    myresult = Strs.PadRight("Hello World", 32, "abcd")
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


'@TestMethod("Strs")
Private Sub Test05a_Strs_PadLeft_Default()

    #If TWINBASIC Then
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
    
    Dim myresult As String
   
    'Act:
    myresult = Strs.PadLeft("Hello World", 32)
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


'@TestMethod("Strs")
Private Sub Test05b_Strs_PadLeft_hash()

    #If TWINBASIC Then
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
    
    Dim myresult As String
   
    'Act:
    myresult = Strs.PadLeft("Hello World", 32, "#")
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


'@TestMethod("Strs")
Private Sub Test05c_Strs_PadLeft_abcd()

    #If TWINBASIC Then
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
    
    Dim myresult As String
   
    'Act:
    myresult = Strs.PadLeft("Hello World", 32, "abcd")
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


'@TestMethod("Strs")
Private Sub Test06a_Strs_CountOf_char()

    #If TWINBASIC Then
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
    
    Dim myresult As Long
   
    'Act:
    myresult = Strs.Countof("Hello World", "l")
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


'@TestMethod("Strs")
Private Sub Test06b_Strs_CountOf_subStr()

    #If TWINBASIC Then
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
    
    Dim myresult As Long
   
    'Act:
    myresult = Strs.Countof("Hello Worldel", "el")
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


'@TestMethod("Strs")
Private Sub Test07a_SubStr_Default()

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
    myExpected = Split("Hello,There,World", ",")
    ReDim Preserve myExpected(1 To 3)
    
    
    Dim myresult As Variant
    'Act:
    myresult = Strs.ToSubStr("Hello,There,World")
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


'@TestMethod("Strs")
Private Sub Test07b_SubStr_ab()

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
    myExpected = Split("Hello_ab_There_ab_World", "_ab_")
    ReDim Preserve myExpected(1 To 3)
    
    Dim myresult As Variant
    'Act:
    myresult = Strs.ToSubStr("Hello_ab_There_ab_World", "_ab_")
   
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


'@TestMethod("Strs")
Private Sub Test08a_Repeat_Default()

    #If TWINBASIC Then
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
   
    Dim myresult As String
    'Act:
    myresult = Strs.Repeat(" ", 8)
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


'@TestMethod("Strs")
Private Sub Test08b_Repeat_Hello()

    #If TWINBASIC Then
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
   
    Dim myresult As String
    'Act:
    myresult = Strs.Repeat("Hello", 8)
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


'@TestMethod("Strs")
Private Sub Test09a_Replacer()

    #If TWINBASIC Then
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
   
    Dim myresult As String
    'Act:
    myresult = Strs.Replacer("    He llo   Worl   d ")
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


'@TestMethod("Strs")
Private Sub Test09b_Replacer_SpecifiedPair()

    #If TWINBASIC Then
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
   
    Dim myresult As String
    'Act:
    myresult = Strs.Replacer("     He llo   Worl   d ", Chars.twSpace, "a")
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


'@TestMethod("Strs")
Private Sub Test09c_Replacer_NestedPairs()

    #If TWINBASIC Then
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
   
    Dim myresult As String
    'Act:
    myresult = Strs.Replacer("HelloaaaaapppppWorld", "ap", Chars.twNullStr)
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


'@TestMethod("Strs")
Private Sub Test10a_MultiReplacer()

    #If TWINBASIC Then
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
   
    Dim myresult As String
    'Act:
    myresult = Strs.MultiReplacer("    He llo   Worl   d ")
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


'@TestMethod("Strs")
Private Sub Test10b_MultiReplacer_SpecifiedPair()

    #If TWINBASIC Then
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
   
    Dim myresult As String
    'Act:
    myresult = Strs.MultiReplacer("     He llo   Worl   d ", Array(Array(Chars.twSpace, "a")))
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


'@TestMethod("Strs")
Private Sub Test10c_MultiReplacer_NestedPairs()

    #If TWINBASIC Then
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
   
    Dim myresult As String
    'Act:
    myresult = Strs.MultiReplacer("HelloaaaaapppppWorld", Array(Array("ap", Chars.twNullStr), Array("l", "L")))
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


'@TestMethod("Strs")
Private Sub Test11a_ToAscB()

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
    myExpected = Array(AscB("H"), AscB("e"), AscB("l"), AscB("l"), AscB("o"))
    ReDim Preserve myExpected(1 To 5)
    Dim myresult As Variant
    'Act:
    myresult = Strs.ToAscB("Hello")
    ReDim Preserve myresult(1 To 5)
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


'@TestMethod("Strs")
Private Sub Test11b_ToUnicodeBytes()

    #If TWINBASIC Then
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
    
    Dim myresult As Variant
    'Act:
    myresult = Strs.ToUnicodeBytes("Hello")
    ReDim Preserve myresult(1 To 10)
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


'@TestMethod("Strs")
Private Sub Test11c_ToUnicodeIntegers()

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
    myExpected = Array(AscW("H"), AscW("e"), AscW("l"), AscW("l"), AscW("o"))
    ReDim Preserve myExpected(1 To 5)
    
    
    Dim myresult As Variant
    'Act:
    myresult = Strs.ToUnicodeIntegers("Hello")
    ReDim Preserve myresult(1 To 5)
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


'@TestMethod("Strs")
Private Sub Test12a_Sort()

    #If TWINBASIC Then
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
        
    Dim myresult As String
    'Act:
    myresult = Strs.Sort("oleHl")
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


'@TestMethod("Strs")
Private Sub Test13a_Inc_NoCarryInc()

    #If TWINBASIC Then
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
    
    Dim myresult As String
    
    'Act:
    myresult = Strs.Inc(myString)
    
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


'@TestMethod("Strs")
Private Sub Test13b_Inc_LastCharNotIncrementable()

    #If TWINBASIC Then
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
    
    Dim myresult As String
    
    'Act:
    myresult = Strs.Inc(myString)
    
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


'@TestMethod("Strs")
Private Sub Test13c_Inc_FullROllover()

    #If TWINBASIC Then
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
    
    Dim myresult As String
    
    'Act:
    myresult = Strs.Inc(myString)
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


'@TestMethod("Strs")
Private Sub Test13d_Inc_NonIncMidString()

    #If TWINBASIC Then
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
    
    Dim myresult As String
    
    'Act:
    myresult = Strs.Inc(myString)
    
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


