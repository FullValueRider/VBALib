Attribute VB_Name = "TestStrs"
'@TestModule
'@Folder("Tests")
'@IgnoreModule


Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub


'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub


'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod("Strs")
Private Sub Test32a_strs_BinToNum_Byte()

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
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test32b_strs_BinToNum_Integer()

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
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test32c_strs_BinToNum_Long()

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
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test32c_strs_BinToNum_LongLong()

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
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub




'@TestMethod("Strs")
Private Sub Test36a_Strs_Dedup_Default()

    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As String
    myExpected = "Hello World"
    
    Dim myresult As String
   
   
    'Act:
   
    myresult = Strs.Dedup("Hello   World")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test36b_Strs_Dedup()

    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As String
    myExpected = "Helo World"
    
    Dim myresult As String
   
   
    'Act:
   
    myresult = Strs.Dedup("Heeellllo   Worlld", SeqC("e", "l", " "))
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test37a_Strs_Trimmer_Default()

    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As String
    myExpected = "Hello World"
    
    Dim myresult As String
   
   
    'Act:
   
    myresult = Strs.Trimmer(" ,  ;   Hello World ,,,  ;")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test37b_Strs_Trimmer_Spaces()

    On Error GoTo TestFail
    
    'Arrange:
    
    Dim myExpected As String
    myExpected = "Hello World,"
    
    Dim myresult As String
   
   
    'Act:
   
    myresult = Strs.Trimmer("     Hello World,     ", " ")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test38a_Strs_PadRight_Default()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello World                     "
    
    Dim myresult As String
   
    'Act:
    myresult = Strs.PadRight("Hello World", 32)
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test38b_Strs_PadRight_hash()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello World#####################"
    
    Dim myresult As String
   
    'Act:
    myresult = Strs.PadRight("Hello World", 32, "#")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test38c_Strs_PadRight_abcd()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello Worldabcdabcdabcdabcdabcda"
    
    Dim myresult As String
   
    'Act:
    myresult = Strs.PadRight("Hello World", 32, "abcd")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test39a_Strs_PadLeft_Default()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "                     Hello World"
    
    Dim myresult As String
   
    'Act:
    myresult = Strs.PadLeft("Hello World", 32)
    'Assert:
    
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test39b_Strs_PadLeft_hash()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "#####################Hello World"
    
    Dim myresult As String
   
    'Act:
    myresult = Strs.PadLeft("Hello World", 32, "#")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test39c_Strs_PadLeft_abcd()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "dabcdabcdabcdabcdabcdHello World"
    
    Dim myresult As String
   
    'Act:
    myresult = Strs.PadLeft("Hello World", 32, "abcd")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test40a_Strs_CountOf_char()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 3
    
    Dim myresult As Long
   
    'Act:
    myresult = Strs.CountOf("Hello World", "l")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test40b_Strs_CountOf_subStr()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Long
    myExpected = 2
    
    Dim myresult As Long
   
    'Act:
    myresult = Strs.CountOf("Hello Worldel", "el")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test41a_SubStr_Default()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Split("Hello,There,World", ",")
    ReDim Preserve myExpected(1 To 3)
    
    Dim myresult As Variant
    'Act:
    myresult = Strs.ToSubStr("Hello,There,World").ToArray
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test41b_SubStr_ab()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Split("Hello_ab_There_ab_World", "_ab_")
    ReDim Preserve myExpected(1 To 3)
    
    Dim myresult As Variant
    'Act:
    myresult = Strs.ToSubStr("Hello_ab_There_ab_World", "_ab_").ToArray
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test42a_Repeat_Default()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "        "
   
    Dim myresult As String
    'Act:
    myresult = Strs.Repeat(" ", 8)
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test42b_Repeat_Hello()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "HelloHelloHelloHelloHelloHelloHelloHello"
   
    Dim myresult As String
    'Act:
    myresult = Strs.Repeat("Hello", 8)
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test43a_Replacer()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "HelloWorld"
   
    Dim myresult As String
    'Act:
    myresult = Strs.Replacer("    He llo   Worl   d ")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test43b_Replacer_SpecifiedPair()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "aaaaaHealloaaaWorlaaada"
   
    Dim myresult As String
    'Act:
    myresult = Strs.Replacer("     He llo   Worl   d ", Chars.twSpace, "a")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test43c_Replacer_NestedPairs()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "HelloWorld"
   
    Dim myresult As String
    'Act:
    myresult = Strs.Replacer("HelloaaaaapppppWorld", "ap", Chars.twNullStr)
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test44a_MultiReplacer()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "HelloWorld"
   
    Dim myresult As String
    'Act:
    myresult = Strs.MultiReplacer("    He llo   Worl   d ")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test44b_MultiReplacer_SpecifiedPair()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "aaaaaHealloaaaWorlaaada"
   
    Dim myresult As String
    'Act:
    myresult = Strs.MultiReplacer("     He llo   Worl   d ", SeqC.Deb.AddItems(Array(Chars.twSpace, "a")))
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test44c_MultiReplacer_NestedPairs()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "HeLLoWorLd"
   
    Dim myresult As String
    'Act:
    myresult = Strs.MultiReplacer("HelloaaaaapppppWorld", SeqC(Array("ap", Chars.twNullStr), Array("l", "L")))
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test45a_ToAscB()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(AscB("H"), AscB("e"), AscB("l"), AscB("l"), AscB("o"))
    ReDim Preserve myExpected(1 To 5)
    Dim myresult As Variant
    'Act:
    myresult = Strs.ToAscB("Hello").ToArray
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test45b_ToUnicodeBytes()

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
    myresult = Strs.ToUnicodeBytes("Hello").ToArray
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test45c_ToUnicodeBytes()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As Variant
    myExpected = Array(AscW("H"), AscW("e"), AscW("l"), AscW("l"), AscW("o"))
    ReDim Preserve myExpected(1 To 5)
    
    
    Dim myresult As Variant
    'Act:
    myresult = Strs.ToUnicodeIntegers("Hello").ToArray
    'Assert:
    Assert.SequenceEquals myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test46a_Sort()

    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected As String
    myExpected = "Hello"
        
    Dim myresult As String
    'Act:
    myresult = Strs.Sort("oleHl")
    'Assert:
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test47a_Inc_NoCarryInc()

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
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Strs")
Private Sub Test47b_Inc_LastCharNotIncrementable()

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
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test47c_Inc_FullROllover()

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
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Strs")
Private Sub Test47d_Inc_NonIncMidString()

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
    Assert.AreEqual myExpected, myresult
    
TestExit:
    '@Ignore UnhandledOnErrorResumeNext
    On Error Resume Next
    
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
